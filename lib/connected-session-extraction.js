export function createConnectedSessionExtractionRuntime({
  appName,
  buildVersion,
  buildTime,
  platform,
  defaultOutputDir,
  maxSavedCandidates,
  buildTranscriptOutputBasePath,
  buildTranscriptOutputPayload,
  saveTranscriptOutputs,
  buildMarkdownOutput,
  saveNetworkCaptureOutput,
  extractMeetingMetadata,
  extractMeetingMetadataFromCapture,
  mergeMeetingMetadata,
  captureStreamNetwork,
  buildCrawlerExtractionTargetUrls,
  navigatePageAndWait,
  isClearlyWrongMeetingLandingPage,
  errorFactory,
}) {
  function buildNetworkCapturePayload({
    options,
    browser,
    profile,
    debugPort,
    targetPage,
    captureResult,
  }) {
    const visibleCandidates = captureResult.candidates.slice(
      0,
      maxSavedCandidates,
    );

    return {
      app: {
        name: appName,
        version: buildVersion,
        buildTime,
        platform,
      },
      run: {
        browser: browser
          ? {
              key: browser.key,
              name: browser.name,
            }
          : null,
        profile: profile
          ? {
              dirName: profile.dirName,
              displayName: profile.displayName,
              email: profile.email,
            }
          : null,
        targetPage: targetPage
          ? {
              title: targetPage.title,
              url: targetPage.url,
            }
          : null,
        outputFormat: options.outputFormat,
        debugPort,
        capturedAt: new Date().toISOString(),
        debugCaptureEnabled: options.debug,
      },
      summary: {
        candidateCount: captureResult.candidates.length,
        captureFeedback: captureResult.captureFeedback || null,
        transcriptMatches: captureResult.transcriptMatchCount || 0,
        matchedCandidate: captureResult.matchedCandidate
          ? {
              url: captureResult.matchedCandidate.url,
              mimeType: captureResult.matchedCandidate.mimeType,
              resourceType: captureResult.matchedCandidate.resourceType,
              decrypted: captureResult.matchedCandidate.decrypted,
              decryptionKeyId: captureResult.matchedCandidate.decryptionKeyId,
              parsedFormat: captureResult.matchedCandidate.parsedFormat,
              parsedPath: captureResult.matchedCandidate.parsedPath,
              parsedEntryCount: captureResult.matchedCandidate.parsedEntryCount,
            }
          : null,
        debug:
          options.debug && captureResult.debug
            ? {
                requestCount: captureResult.debug.requestCount,
                responseCount: captureResult.debug.responseCount,
                failureCount: captureResult.debug.failureCount,
                webSocketCount: captureResult.debug.webSocketCount,
                webSocketFrameCount: captureResult.debug.webSocketFrameCount,
                eventSourceMessageCount:
                  captureResult.debug.eventSourceMessageCount,
                responseBodyCount: captureResult.debug.responseBodies.length,
                protectionKeyCount: captureResult.debug.protectionKeys.length,
              }
            : null,
      },
      candidates: visibleCandidates.map((candidate) => ({
        requestId: candidate.requestId,
        url: candidate.url,
        status: candidate.status,
        mimeType: candidate.mimeType,
        resourceType: candidate.resourceType,
        score: candidate.score,
        bodyLength: candidate.bodyLength,
        bodyError: candidate.bodyError,
        base64Encoded: candidate.base64Encoded,
        rawBodyByteLength: candidate.rawBodyByteLength,
        rawBodyPreviewHex: candidate.rawBodyPreviewHex,
        refetchedExternally: candidate.refetchedExternally,
        refetchError: candidate.refetchError,
        decrypted: candidate.decrypted,
        decryptionKeyId: candidate.decryptionKeyId,
        decryptionError: candidate.decryptionError,
        parsedFormat: candidate.parsedFormat,
        parsedPath: candidate.parsedPath,
        parsedEntryCount: candidate.parsedEntryCount,
        parsedEntrySample: candidate.parsedEntries.slice(0, 5),
        headers: candidate.headers,
        bodyPreview: candidate.bodyPreview,
        body: options.debug ? candidate.body : undefined,
        rawBodyBase64:
          options.debug && candidate.base64Encoded
            ? candidate.rawBodyBase64
            : undefined,
      })),
      debug: options.debug ? captureResult.debug : undefined,
    };
  }

  function buildMissingTranscriptCaptureMessage(
    captureResult,
    debugEnabled,
    captureControl = 'manual',
  ) {
    const candidateResponseCount =
      captureResult.captureFeedback?.candidateResponseCount || 0;
    const transcriptHintCount =
      captureResult.captureFeedback?.transcriptHintCount || 0;

    let guidance =
      'No transcript payload was observed after network capture started.';

    if (candidateResponseCount === 0) {
      guidance +=
        ' No transcript-related network responses were captured while the terminal was armed.';
    } else if (transcriptHintCount === 0) {
      guidance +=
        ' Some Stream responses were captured, but none looked transcript-specific.';
    } else {
      guidance +=
        ' Transcript-like traffic was detected, but none of the captured bodies parsed into transcript entries.';
    }

    if (debugEnabled) {
      return `${guidance} Review the saved .network.json capture.`;
    }

    if (captureControl === 'automatic') {
      return (
        `${guidance} Rerun with --debug to save a .network.json capture and ` +
        'print the automatic action trace.'
      );
    }

    return (
      `${guidance} Rerun with --debug to save a .network.json capture and ` +
      'auto-reload the page once capture is armed.'
    );
  }

  function buildFallbackMeetingMetadata(targetPage) {
    return {
      title: targetPage?.title || '',
      date: '',
      recordedBy: '',
      sourceUrl: targetPage?.url || '',
      createdBy: '',
      createdByEmail: '',
      createdByTenantId: '',
      sourceApplication: '',
      recordingStartDateTime: '',
      recordingEndDateTime: '',
      sharePointFilePath: '',
      sharePointItemUrl: '',
    };
  }

  function attachCliErrorDetails(error, details) {
    if (error && typeof error === 'object') {
      error.details = {
        ...(error.details || {}),
        ...details,
      };
    }

    return error;
  }

  async function extractTranscriptFromConnectedPage({
    cdp,
    prompt,
    options,
    browser,
    profile,
    debugPort,
    targetPage,
    captureControl = 'manual',
    allowManualAssist = true,
  }) {
    await cdp.send('Runtime.enable');

    console.log(
      captureControl === 'automatic'
        ? 'Capturing transcript-related network responses in automatic mode...'
        : 'Capturing transcript-related network responses...',
    );
    if (options.debug) {
      console.log(
        'Debug capture enabled: saving request/response lifecycle data, ' +
          'candidate bodies, and WebSocket frames.',
      );
    }

    const captureResult = await captureStreamNetwork(
      cdp,
      prompt,
      options.debug,
      captureControl,
      {
        allowManualAssist,
      },
    );
    const pageMetadata = await extractMeetingMetadata(cdp).catch(() =>
      buildFallbackMeetingMetadata(targetPage),
    );
    const captureMetadata = extractMeetingMetadataFromCapture(captureResult);
    const metadata = mergeMeetingMetadata(
      pageMetadata,
      captureMetadata,
      targetPage,
    );
    const outputBasePath = buildTranscriptOutputBasePath({
      defaultName: metadata.title || targetPage?.title || 'meeting',
      outputName: options.outputName,
      outputDir: options.outputDir,
      defaultOutputDir,
    });
    let networkOutputPath = '';

    console.log(
      `Observed ${captureResult.candidates.length} potentially relevant network response` +
        `${captureResult.candidates.length === 1 ? '' : 's'}.`,
    );
    console.log(
      `Parsed ${captureResult.transcriptMatchCount || 0} transcript payload match` +
        `${captureResult.transcriptMatchCount === 1 ? '' : 'es'}.`,
    );

    if (!captureResult.matchedCandidate) {
      if (options.debug) {
        const networkCapturePayload = buildNetworkCapturePayload({
          options,
          browser,
          profile,
          debugPort,
          targetPage,
          captureResult,
        });
        networkOutputPath = saveNetworkCaptureOutput(
          networkCapturePayload,
          outputBasePath,
        );
        console.log(`Saved network capture to: ${networkOutputPath}`);
      }

      throw attachCliErrorDetails(
        errorFactory(
          buildMissingTranscriptCaptureMessage(
            captureResult,
            options.debug,
            captureControl,
          ),
        ),
        {
          networkOutputPath,
          outputBasePath,
        },
      );
    }

    if (options.debug) {
      const networkCapturePayload = buildNetworkCapturePayload({
        options,
        browser,
        profile,
        debugPort,
        targetPage,
        captureResult,
      });
      networkOutputPath = saveNetworkCaptureOutput(
        networkCapturePayload,
        outputBasePath,
      );
    }

    const entries = captureResult.matchedCandidate.parsedEntries;
    const outputPayload = buildTranscriptOutputPayload(metadata, entries);
    const outputPaths = saveTranscriptOutputs({
      payload: outputPayload,
      outputBasePath,
      outputFormat: options.outputFormat,
      renderMarkdown: buildMarkdownOutput,
    });

    console.log(
      `Matched transcript payload: ${captureResult.matchedCandidate.url}`,
    );
    console.log(`Parsed ${entries.length} transcript entries.`);
    for (const outputPath of outputPaths) {
      console.log(`Saved transcript to: ${outputPath}`);
    }
    if (networkOutputPath) {
      console.log(`Saved network capture to: ${networkOutputPath}`);
    }

    return {
      metadata,
      outputPayload,
      outputBasePath,
      outputPaths,
      networkOutputPath,
      matchedCandidateUrl: captureResult.matchedCandidate.url,
      entryCount: entries.length,
      captureSummary: {
        candidateResponseCount: captureResult.candidates.length,
        transcriptMatchCount: captureResult.transcriptMatchCount || 0,
      },
    };
  }

  async function extractCrawlerItemInConnectedSession({
    cdp,
    prompt,
    options,
    browser,
    profile,
    debugPort,
    targetPage,
    item,
    captureControl = 'automatic',
    allowManualAssist = false,
  }) {
    const sessionPrompt = prompt || {
      ask: async () => '',
      waitForEnter: async () => {},
      close() {},
    };
    const candidateTargetUrls = buildCrawlerExtractionTargetUrls(item);
    let extractionResult = null;
    let extractionTargetUrl = '';
    let extractionPageSnapshot = null;
    let lastExtractionError = null;

    for (
      let candidateIndex = 0;
      candidateIndex < candidateTargetUrls.length;
      candidateIndex += 1
    ) {
      const candidateTargetUrl = candidateTargetUrls[candidateIndex];
      console.log(
        `Trying playback URL ${candidateIndex + 1}/${candidateTargetUrls.length}: ${candidateTargetUrl}`,
      );

      const pageSnapshot = await navigatePageAndWait(cdp, candidateTargetUrl);
      if (isClearlyWrongMeetingLandingPage(pageSnapshot, candidateTargetUrl)) {
        lastExtractionError = errorFactory(
          `Navigation landed on an unexpected page instead of the meeting recording: ` +
            `${pageSnapshot.title || pageSnapshot.url || candidateTargetUrl}`,
        );
        console.log(
          `Skipping URL because it landed on "${pageSnapshot.title || pageSnapshot.url}".`,
        );
        continue;
      }

      const currentTargetPage = {
        ...targetPage,
        title: pageSnapshot.title || item?.title || targetPage?.title || '',
        url: pageSnapshot.url || candidateTargetUrl,
      };

      try {
        extractionResult = await extractTranscriptFromConnectedPage({
          cdp,
          prompt: sessionPrompt,
          options,
          browser,
          profile,
          debugPort,
          targetPage: currentTargetPage,
          captureControl,
          allowManualAssist,
        });
        extractionTargetUrl = candidateTargetUrl;
        extractionPageSnapshot = pageSnapshot;
        break;
      } catch (error) {
        lastExtractionError = error;
        if (candidateIndex < candidateTargetUrls.length - 1) {
          console.log(
            'Extraction did not succeed from this URL. Trying the next candidate...',
          );
        }
      }
    }

    if (!extractionResult) {
      throw (
        lastExtractionError ||
        errorFactory(
          'No usable playback URL produced a transcript for this meeting.',
        )
      );
    }

    return {
      ...extractionResult,
      candidateTargetUrls,
      extractionTargetUrl:
        extractionTargetUrl || extractionPageSnapshot?.url || item?.url || '',
      extractionPageSnapshot,
    };
  }

  return {
    extractTranscriptFromConnectedPage,
    extractCrawlerItemInConnectedSession,
  };
}
