export function createNetworkTranscriptCaptureRuntime({
  createCdpEventScope,
  createEmptyResponseBodyRecord,
  loadResponseBody,
  refetchBinaryBody,
  reloadPageAndWait,
  runAutomaticTranscriptCaptureFlow,
  sleep,
  printCaptureInstructions,
  logAutomaticAction,
  sanitizeHeaders,
  lowerCaseHeaderMap,
  trimForPreview,
  containsTranscriptSignal,
  shouldTrackDebugTraffic,
  shouldFetchDebugResponseBody,
  summarizeWebSocketPayload,
  shouldTrackNetworkResponse,
  scoreNetworkResponse,
  isLikelyTranscriptResponseSignal,
  formatNetworkResponseForTerminal,
  buildProtectionKeyMap,
  isEncryptedTranscriptUrl,
  getUrlSearchParam,
  finalizeBodyRecordForResponse,
  summarizeCapturedBody,
  scoreTranscriptCandidate,
  isUsableTranscriptCandidate,
  maxTrackedResponses,
  maxDebugResponses,
  maxDebugRequests,
  maxDebugBodies,
  maxDebugWebSockets,
  maxDebugWebSocketFrames,
  maxDebugPostDataPreviewLength,
  maxDebugFramePreviewLength,
  defaultCaptureSettleMs,
}) {
  async function captureStreamNetwork(
    cdp,
    prompt,
    debugEnabled = false,
    captureControl = 'manual',
    { allowManualAssist = true } = {},
  ) {
    await cdp.send('Network.enable', {
      maxTotalBufferSize: 100_000_000,
      maxResourceBufferSize: 10_000_000,
    });

    const eventScope = createCdpEventScope(cdp);
    const trackedResponses = new Map();
    const finishedRequests = new Set();
    const failedRequests = new Map();
    const bodyCache = new Map();
    const captureFeedback = {
      candidateResponseCount: 0,
      transcriptHintCount: 0,
      transcriptHintSettledCount: 0,
      transcriptHintFailedCount: 0,
      firstTranscriptHint: null,
      printedTranscriptHint: false,
      automaticActions: [],
    };
    const transcriptHintRequestIds = new Set();
    const debugState = debugEnabled
      ? {
          requests: new Map(),
          responses: new Map(),
          failures: [],
          webSockets: new Map(),
          webSocketFrames: [],
          eventSourceMessages: [],
        }
      : null;

    function ensureTrackedWebSocket(requestId, url = '') {
      if (!debugState) {
        return null;
      }

      const normalizedRequestId = String(requestId || '');
      if (!normalizedRequestId) {
        return null;
      }

      if (debugState.webSockets.has(normalizedRequestId)) {
        const socket = debugState.webSockets.get(normalizedRequestId);
        if (url && !socket.url) {
          socket.url = url;
        }
        return socket;
      }

      if (
        debugState.webSockets.size >= maxDebugWebSockets ||
        !shouldTrackDebugTraffic({
          url,
          resourceType: 'WebSocket',
          captureAll: true,
        })
      ) {
        return null;
      }

      const socket = {
        requestId: normalizedRequestId,
        url,
        created: false,
        createdAt: null,
        handshakeRequest: null,
        handshakeResponse: null,
        framesSent: 0,
        framesReceived: 0,
        transcriptSignalDetected: false,
        closedAt: null,
      };
      debugState.webSockets.set(normalizedRequestId, socket);
      return socket;
    }

    eventScope.on('Network.responseReceived', (params) => {
      if (trackedResponses.size >= maxTrackedResponses) {
        // Continue recording extended debug traffic even after candidate capture is full.
      } else if (shouldTrackNetworkResponse(params.response, params.type)) {
        const headers = sanitizeHeaders(params.response.headers);
        const lowerHeaders = lowerCaseHeaderMap(params.response.headers);
        const requestId = String(params.requestId);
        const url = String(params.response.url || '');
        const mimeType = String(params.response.mimeType || '');
        const resourceType = String(params.type || '');
        const contentType = String(lowerHeaders['content-type'] || '');
        const score = scoreNetworkResponse({
          url: url.toLowerCase(),
          mimeType: mimeType.toLowerCase(),
          contentType: contentType.toLowerCase(),
          contentDisposition: String(
            lowerHeaders['content-disposition'] || '',
          ).toLowerCase(),
          resourceType,
        });

        trackedResponses.set(requestId, {
          requestId,
          url,
          status: params.response.status,
          mimeType,
          resourceType,
          headers,
          score,
        });
        captureFeedback.candidateResponseCount = trackedResponses.size;

        if (
          isLikelyTranscriptResponseSignal({
            url,
            mimeType,
            contentType,
            resourceType,
            score,
          })
        ) {
          transcriptHintRequestIds.add(requestId);
          captureFeedback.transcriptHintCount += 1;
          if (captureControl === 'automatic') {
            logAutomaticAction(
              captureFeedback,
              debugEnabled,
              'transcript hint response received',
              {
                requestId,
                url,
                resourceType,
                mimeType,
                contentType,
                score,
              },
            );
          }

          if (!captureFeedback.firstTranscriptHint) {
            captureFeedback.firstTranscriptHint = {
              requestId,
              url,
              resourceType,
              mimeType,
              contentType,
              score,
            };
          }

          if (!captureFeedback.printedTranscriptHint) {
            captureFeedback.printedTranscriptHint = true;
            console.log(
              `\nDetected likely transcript network response: ` +
                `${formatNetworkResponseForTerminal({
                  url,
                  resourceType,
                  mimeType,
                  contentType,
                })}`,
            );
            console.log(
              captureControl === 'automatic'
                ? 'Automatic mode will keep waiting for the response to settle.'
                : 'Let the transcript finish loading, then press Enter to continue.',
            );
          }
        }
      }

      if (!debugState) {
        return;
      }

      const requestId = String(params.requestId);
      const url = String(params.response.url || '');
      const resourceType = String(params.type || '');
      if (
        !debugState.responses.has(requestId) &&
        debugState.responses.size >= maxDebugResponses
      ) {
        return;
      }

      if (
        !shouldTrackDebugTraffic({
          url,
          resourceType,
          headers: params.response.headers,
          captureAll: true,
        })
      ) {
        return;
      }

      const headers = sanitizeHeaders(params.response.headers);
      const lowerHeaders = lowerCaseHeaderMap(params.response.headers);
      const existingResponse = debugState.responses.get(requestId);
      debugState.responses.set(requestId, {
        requestId,
        url,
        status: params.response.status,
        statusText: String(params.response.statusText || ''),
        mimeType: String(params.response.mimeType || ''),
        contentType: String(lowerHeaders['content-type'] || ''),
        resourceType,
        headers,
        fromDiskCache: Boolean(params.response.fromDiskCache),
        fromServiceWorker: Boolean(params.response.fromServiceWorker),
        fromPrefetchCache: Boolean(params.response.fromPrefetchCache),
        encodedDataLength: existingResponse?.encodedDataLength ?? null,
        finished: existingResponse?.finished ?? false,
        loadingFailure: existingResponse?.loadingFailure || '',
        remoteIPAddress: String(params.response.remoteIPAddress || ''),
        protocol: String(params.response.protocol || ''),
      });
    });

    eventScope.on('Network.requestWillBeSent', (params) => {
      if (!debugState) {
        return;
      }

      const requestId = String(params.requestId);
      const request = params.request || {};
      const url = String(request.url || '');
      const resourceType = String(params.type || '');
      if (
        !debugState.requests.has(requestId) &&
        debugState.requests.size >= maxDebugRequests
      ) {
        return;
      }

      if (
        !shouldTrackDebugTraffic({
          url,
          resourceType,
          method: String(request.method || ''),
          headers: request.headers,
          postData: String(request.postData || ''),
          captureAll: true,
        })
      ) {
        return;
      }

      debugState.requests.set(requestId, {
        requestId,
        url,
        method: String(request.method || ''),
        resourceType,
        headers: sanitizeHeaders(request.headers),
        hasPostData:
          typeof request.postData === 'string' && request.postData.length > 0,
        postDataPreview:
          typeof request.postData === 'string'
            ? trimForPreview(
                request.postData,
                maxDebugPostDataPreviewLength,
              )
            : '',
        initiatorType: String(params.initiator?.type || ''),
        documentURL: String(params.documentURL || ''),
        timestamp:
          typeof params.timestamp === 'number' ? params.timestamp : null,
        wallTime:
          typeof params.wallTime === 'number' ? params.wallTime : null,
        redirectResponseStatus:
          typeof params.redirectResponse?.status === 'number'
            ? params.redirectResponse.status
            : null,
        redirectResponseUrl: String(params.redirectResponse?.url || ''),
      });
    });

    eventScope.on('Network.loadingFinished', (params) => {
      const requestId = String(params.requestId);
      finishedRequests.add(requestId);
      if (transcriptHintRequestIds.has(requestId)) {
        captureFeedback.transcriptHintSettledCount += 1;
        if (captureControl === 'automatic') {
          logAutomaticAction(
            captureFeedback,
            debugEnabled,
            'transcript hint response finished',
            {
              requestId,
            },
          );
        }
      }

      if (!debugState) {
        return;
      }

      const responseRecord = debugState.responses.get(requestId);
      if (responseRecord) {
        responseRecord.finished = true;
        responseRecord.encodedDataLength =
          typeof params.encodedDataLength === 'number'
            ? params.encodedDataLength
            : null;
      }
    });

    eventScope.on('Network.loadingFailed', (params) => {
      const requestId = String(params.requestId);
      const errorText = params.errorText || 'Request failed';
      failedRequests.set(requestId, errorText);
      if (transcriptHintRequestIds.has(requestId)) {
        captureFeedback.transcriptHintFailedCount += 1;
        if (captureControl === 'automatic') {
          logAutomaticAction(
            captureFeedback,
            debugEnabled,
            'transcript hint response failed',
            {
              requestId,
              errorText,
            },
          );
        }
      }

      if (!debugState) {
        return;
      }

      const responseRecord = debugState.responses.get(requestId);
      if (responseRecord) {
        responseRecord.loadingFailure = errorText;
      }
      debugState.failures.push({
        requestId,
        errorText,
        canceled: Boolean(params.canceled),
        blockedReason: String(params.blockedReason || ''),
        resourceType: String(params.type || ''),
      });
    });

    eventScope.on('Network.webSocketCreated', (params) => {
      const socket = ensureTrackedWebSocket(
        params.requestId,
        String(params.url || ''),
      );
      if (!socket) {
        return;
      }

      socket.created = true;
      socket.createdAt =
        typeof params.timestamp === 'number' ? params.timestamp : null;
    });

    eventScope.on('Network.webSocketWillSendHandshakeRequest', (params) => {
      const socket = ensureTrackedWebSocket(params.requestId);
      if (!socket) {
        return;
      }

      socket.handshakeRequest = {
        timestamp:
          typeof params.timestamp === 'number' ? params.timestamp : null,
        wallTime:
          typeof params.wallTime === 'number' ? params.wallTime : null,
        headers: sanitizeHeaders(params.request?.headers),
      };
    });

    eventScope.on('Network.webSocketHandshakeResponseReceived', (params) => {
      const socket = ensureTrackedWebSocket(params.requestId);
      if (!socket) {
        return;
      }

      socket.handshakeResponse = {
        timestamp:
          typeof params.timestamp === 'number' ? params.timestamp : null,
        status:
          typeof params.response?.status === 'number'
            ? params.response.status
            : null,
        statusText: String(params.response?.statusText || ''),
        headers: sanitizeHeaders(params.response?.headers),
      };
    });

    eventScope.on('Network.webSocketFrameReceived', (params) => {
      if (!debugState) {
        return;
      }

      const socket = ensureTrackedWebSocket(params.requestId);
      if (!socket) {
        return;
      }

      socket.framesReceived += 1;
      const payload = summarizeWebSocketPayload(
        params.response?.payloadData || '',
      );
      socket.transcriptSignalDetected =
        socket.transcriptSignalDetected || payload.transcriptSignalDetected;

      if (debugState.webSocketFrames.length >= maxDebugWebSocketFrames) {
        return;
      }

      debugState.webSocketFrames.push({
        requestId: String(params.requestId),
        direction: 'received',
        timestamp:
          typeof params.timestamp === 'number' ? params.timestamp : null,
        opcode:
          typeof params.response?.opcode === 'number'
            ? params.response.opcode
            : null,
        mask: Boolean(params.response?.mask),
        ...payload,
      });
    });

    eventScope.on('Network.webSocketFrameSent', (params) => {
      if (!debugState) {
        return;
      }

      const socket = ensureTrackedWebSocket(params.requestId);
      if (!socket) {
        return;
      }

      socket.framesSent += 1;
      const payload = summarizeWebSocketPayload(
        params.response?.payloadData || '',
      );
      socket.transcriptSignalDetected =
        socket.transcriptSignalDetected || payload.transcriptSignalDetected;

      if (debugState.webSocketFrames.length >= maxDebugWebSocketFrames) {
        return;
      }

      debugState.webSocketFrames.push({
        requestId: String(params.requestId),
        direction: 'sent',
        timestamp:
          typeof params.timestamp === 'number' ? params.timestamp : null,
        opcode:
          typeof params.response?.opcode === 'number'
            ? params.response.opcode
            : null,
        mask: Boolean(params.response?.mask),
        ...payload,
      });
    });

    eventScope.on('Network.webSocketClosed', (params) => {
      const socket = ensureTrackedWebSocket(params.requestId);
      if (!socket) {
        return;
      }

      socket.closedAt =
        typeof params.timestamp === 'number' ? params.timestamp : null;
    });

    eventScope.on('Network.eventSourceMessageReceived', (params) => {
      if (!debugState) {
        return;
      }

      const requestId = String(params.requestId);
      const requestRecord = debugState.requests.get(requestId);
      const responseRecord = debugState.responses.get(requestId);
      const url = responseRecord?.url || requestRecord?.url || '';
      if (
        !url ||
        !shouldTrackDebugTraffic({
          url,
          resourceType: 'EventSource',
          captureAll: true,
        })
      ) {
        return;
      }

      debugState.eventSourceMessages.push({
        requestId,
        timestamp:
          typeof params.timestamp === 'number' ? params.timestamp : null,
        eventName: String(params.eventName || ''),
        eventId: String(params.eventId || ''),
        dataPreview: trimForPreview(
          String(params.data || ''),
          maxDebugFramePreviewLength,
        ),
        transcriptSignalDetected: containsTranscriptSignal(params.data || ''),
      });
    });

    try {
      if (captureControl === 'automatic') {
        await runAutomaticTranscriptCaptureFlow(
          cdp,
          prompt,
          captureFeedback,
          debugEnabled,
          {
            allowManualAssist,
          },
        );
      } else if (debugEnabled) {
        console.log('Reloading the page with capture armed...');
        await reloadPageAndWait(cdp);
        printCaptureInstructions(debugEnabled);
        await prompt.waitForEnter(
          '\nPress Enter after the transcript panel has loaded...\n',
        );
        await sleep(defaultCaptureSettleMs);
      } else {
        printCaptureInstructions(debugEnabled);
        await prompt.waitForEnter(
          '\nPress Enter after the transcript panel has loaded...\n',
        );
        await sleep(defaultCaptureSettleMs);
      }

      eventScope.stopAll();

      const debugBodyTargets = debugState
        ? Array.from(debugState.responses.values())
            .filter(shouldFetchDebugResponseBody)
            .slice(0, maxDebugBodies)
        : [];
      const bodyTargets = new Map();

      for (const candidate of trackedResponses.values()) {
        bodyTargets.set(candidate.requestId, {
          requestId: candidate.requestId,
          url: candidate.url,
          status: candidate.status,
          mimeType: candidate.mimeType,
          resourceType: candidate.resourceType,
          contentType: String(
            lowerCaseHeaderMap(candidate.headers)['content-type'] || '',
          ),
          finished: finishedRequests.has(candidate.requestId),
          loadingFailure: failedRequests.get(candidate.requestId) || '',
          wasCandidate: true,
        });
      }

      for (const responseRecord of debugBodyTargets) {
        if (bodyTargets.has(responseRecord.requestId)) {
          continue;
        }

        bodyTargets.set(responseRecord.requestId, {
          requestId: responseRecord.requestId,
          url: responseRecord.url,
          status: responseRecord.status,
          mimeType: responseRecord.mimeType,
          resourceType: responseRecord.resourceType,
          contentType: responseRecord.contentType,
          finished: responseRecord.finished,
          loadingFailure: responseRecord.loadingFailure || '',
          wasCandidate: false,
        });
      }

      const bodyRecordsByRequestId = new Map();

      for (const target of bodyTargets.values()) {
        const requestId = target.requestId;
        const failed = target.loadingFailure || failedRequests.get(requestId) || '';
        const finished = target.finished || finishedRequests.has(requestId);
        let bodyRecord;

        if (finished && !failed) {
          bodyRecord = await loadResponseBody(cdp, requestId, bodyCache);
        } else if (failed) {
          bodyRecord = createEmptyResponseBodyRecord(failed);
        } else {
          bodyRecord = createEmptyResponseBodyRecord(
            'Request did not finish before capture ended.',
          );
        }

        bodyRecordsByRequestId.set(requestId, bodyRecord);
      }

      const protectionKeys = buildProtectionKeyMap(
        Array.from(bodyTargets.values()).map((target) => ({
          url: target.url,
          body: bodyRecordsByRequestId.get(target.requestId)?.body || '',
        })),
      );

      for (const target of bodyTargets.values()) {
        if (!isEncryptedTranscriptUrl(target.url)) {
          continue;
        }

        const existingBodyRecord = bodyRecordsByRequestId.get(target.requestId);
        if (!existingBodyRecord || existingBodyRecord.bodyError) {
          continue;
        }

        const keyId = getUrlSearchParam(target.url, 'kid');
        const hasProtectionKey =
          (keyId && protectionKeys.has(keyId)) || protectionKeys.size === 1;
        if (!hasProtectionKey) {
          continue;
        }

        if (
          existingBodyRecord.base64Encoded &&
          existingBodyRecord.rawBodyBase64
        ) {
          continue;
        }

        try {
          const refetchedBodyRecord = await refetchBinaryBody(target.url);
          bodyRecordsByRequestId.set(target.requestId, {
            ...existingBodyRecord,
            ...refetchedBodyRecord,
          });
        } catch (error) {
          bodyRecordsByRequestId.set(target.requestId, {
            ...existingBodyRecord,
            refetchedExternally: false,
            refetchError:
              error instanceof Error
                ? error.message
                : 'Transcript refetch failed.',
          });
        }
      }

      const candidates = [];

      for (const candidate of trackedResponses.values()) {
        const requestId = candidate.requestId;
        const bodyRecord = finalizeBodyRecordForResponse(
          candidate.url,
          bodyRecordsByRequestId.get(requestId) ||
            createEmptyResponseBodyRecord('Response body was not captured.'),
          protectionKeys,
        );
        const parsed = summarizeCapturedBody(bodyRecord.body);
        candidates.push({
          ...candidate,
          body: bodyRecord.body,
          bodyLength: bodyRecord.bodyLength,
          bodyPreview: bodyRecord.bodyPreview,
          bodyError: bodyRecord.bodyError,
          base64Encoded: bodyRecord.base64Encoded,
          rawBodyBase64: bodyRecord.rawBodyBase64,
          rawBodyByteLength: bodyRecord.rawBodyByteLength,
          rawBodyPreviewHex: bodyRecord.rawBodyPreviewHex,
          refetchedExternally: bodyRecord.refetchedExternally,
          refetchError: bodyRecord.refetchError,
          decrypted: bodyRecord.decrypted,
          decryptionKeyId: bodyRecord.decryptionKeyId,
          decryptionError: bodyRecord.decryptionError,
          parsedFormat: parsed.format,
          parsedPath: parsed.path,
          parsedEntryCount: parsed.entryCount,
          parsedEntries: parsed.entries,
        });
      }

      let debug = null;

      if (debugState) {
        const debugBodies = [];

        for (const responseRecord of debugBodyTargets) {
          const bodyRecord = finalizeBodyRecordForResponse(
            responseRecord.url,
            bodyRecordsByRequestId.get(responseRecord.requestId) ||
              createEmptyResponseBodyRecord('Response body was not captured.'),
            protectionKeys,
          );
          const parsed = summarizeCapturedBody(bodyRecord.body);
          debugBodies.push({
            requestId: responseRecord.requestId,
            url: responseRecord.url,
            status: responseRecord.status,
            mimeType: responseRecord.mimeType,
            contentType: responseRecord.contentType,
            resourceType: responseRecord.resourceType,
            bodyLength: bodyRecord.bodyLength,
            bodyPreview: bodyRecord.bodyPreview,
            body: bodyRecord.body,
            bodyError: bodyRecord.bodyError,
            base64Encoded: bodyRecord.base64Encoded,
            rawBodyBase64: bodyRecord.rawBodyBase64,
            rawBodyByteLength: bodyRecord.rawBodyByteLength,
            rawBodyPreviewHex: bodyRecord.rawBodyPreviewHex,
            refetchedExternally: bodyRecord.refetchedExternally,
            refetchError: bodyRecord.refetchError,
            decrypted: bodyRecord.decrypted,
            decryptionKeyId: bodyRecord.decryptionKeyId,
            decryptionError: bodyRecord.decryptionError,
            parsedFormat: parsed.format,
            parsedPath: parsed.path,
            parsedEntryCount: parsed.entryCount,
            parsedEntrySample: parsed.entries.slice(0, 5),
            wasCandidate: trackedResponses.has(responseRecord.requestId),
          });
        }

        debug = {
          requestCount: debugState.requests.size,
          responseCount: debugState.responses.size,
          failureCount: debugState.failures.length,
          webSocketCount: debugState.webSockets.size,
          webSocketFrameCount: debugState.webSocketFrames.length,
          eventSourceMessageCount: debugState.eventSourceMessages.length,
          protectionKeys: Array.from(protectionKeys.values()).map((key) => ({
            kid: key.kid,
            sourceUrl: key.sourceUrl,
            encryptionAlgorithm: key.encryptionAlgorithm,
            encryptionMode: key.encryptionMode,
            keySize: key.keySize,
            padding: key.padding,
          })),
          requests: Array.from(debugState.requests.values()),
          responses: Array.from(debugState.responses.values()),
          failures: debugState.failures,
          responseBodies: debugBodies,
          webSockets: Array.from(debugState.webSockets.values()),
          webSocketFrames: debugState.webSocketFrames,
          eventSourceMessages: debugState.eventSourceMessages,
        };
      }

      candidates.sort((left, right) => {
        const scoreDelta =
          scoreTranscriptCandidate(right) - scoreTranscriptCandidate(left);
        if (scoreDelta !== 0) {
          return scoreDelta;
        }

        if (right.parsedEntryCount !== left.parsedEntryCount) {
          return right.parsedEntryCount - left.parsedEntryCount;
        }

        return right.score - left.score;
      });

      const transcriptMatches = candidates.filter((candidate) =>
        isUsableTranscriptCandidate(candidate),
      );

      const matchedCandidate = transcriptMatches[0] || null;

      return {
        candidates,
        matchedCandidate,
        transcriptMatchCount: transcriptMatches.length,
        captureFeedback: {
          candidateResponseCount: captureFeedback.candidateResponseCount,
          transcriptHintCount: captureFeedback.transcriptHintCount,
          transcriptHintSettledCount: captureFeedback.transcriptHintSettledCount,
          transcriptHintFailedCount: captureFeedback.transcriptHintFailedCount,
          firstTranscriptHint: captureFeedback.firstTranscriptHint,
          automaticActions: captureFeedback.automaticActions.slice(0, 200),
        },
        debug,
      };
    } finally {
      eventScope.stopAll();
      await cdp.send('Network.disable').catch(() => {});
    }
  }

  return {
    captureStreamNetwork,
  };
}
