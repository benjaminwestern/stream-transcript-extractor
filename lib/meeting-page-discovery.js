export function createMeetingPageDiscoveryRuntime({
  defaultPageNavigationTimeoutMs,
  maxDebugResponses,
  maxDiscoveryResponseBodies,
  createCdpEventScope,
  loadResponseBody,
  navigatePageAndWait,
  getCurrentPageSnapshot,
  evaluate,
  sleep,
  buildMeetingDiscoveryExpression,
  mergeDiscoveredMeetingItems,
  normalizeMeetingUrl,
  normalizeInlineText,
  stripUtf8Bom,
  tryParseJsonLenient,
  getFirstStringValue,
  uniqueNormalizedMeetingUrls,
  normalizeMeetingDownloadUrls,
  isLikelyMeetingItemUrl,
  isLikelyMeetingVideoFileUrl,
  isStreamRelatedUrl,
  lowerCaseHeaderMap,
  errorFactory,
}) {
  function extractUrlsFromText(value) {
    const text = String(value || '');
    const matches = text.match(/https?:\/\/[^\s"'<>\\]+/g) || [];

    return matches
      .map((match) => match.split(/(?:&quot;|&#34;|&apos;|&#39;)/i, 1)[0])
      .map((match) => match.replace(/[),.;]+$/g, ''))
      .filter(Boolean);
  }

  function extractMeetingDiscoveryItemsFromValue(value, sourceUrl = '') {
    const matches = [];
    const seenMatches = new Set();

    function addMatch(candidateUrl, metadata, path) {
      const normalizedUrl = normalizeMeetingUrl(candidateUrl);
      if (!isLikelyMeetingItemUrl(normalizedUrl)) {
        return;
      }

      const key = `${normalizedUrl}|${path}`;
      if (seenMatches.has(key)) {
        return;
      }

      seenMatches.add(key);
      matches.push({
        url: normalizedUrl,
        title: normalizeInlineText(metadata?.title || ''),
        subtitle: normalizeInlineText(metadata?.subtitle || ''),
        identityKey: normalizeInlineText(metadata?.identityKey || ''),
        candidateUrls: [normalizedUrl],
        downloadUrls: isLikelyMeetingVideoFileUrl(normalizedUrl)
          ? [normalizedUrl]
          : [],
        sourceUrl,
        path,
      });
    }

    function walk(current, path, depth = 0) {
      if (depth > 10 || current == null) {
        return;
      }

      if (typeof current === 'string') {
        for (const candidateUrl of extractUrlsFromText(current)) {
          addMatch(candidateUrl, {}, path);
        }
        return;
      }

      if (Array.isArray(current)) {
        current.forEach((entry, index) => {
          walk(entry, `${path}[${index}]`, depth + 1);
        });
        return;
      }

      if (typeof current !== 'object') {
        return;
      }

      const title = getFirstStringValue(current, [
        'title',
        'name',
        'displayName',
        'meetingTitle',
        'subject',
        'fileName',
        'label',
      ]);
      const subtitle = getFirstStringValue(current, [
        'description',
        'sharedBy',
        'sharedByName',
        'createdBy',
        'createdByName',
        'ownerName',
      ]);
      const extension = getFirstStringValue(current, [
        'extension',
        'fileExtension',
        'file_extension',
      ]).toLowerCase();
      const itemType = getFirstStringValue(current, ['type']).toLowerCase();
      const fileType = getFirstStringValue(current, [
        'fileType',
        'file_type',
        'app',
      ]).toLowerCase();
      const isMeetingRecording =
        current?.is_meeting_recording === true ||
        current?.isMeetingRecording === true ||
        String(
          current?.is_meeting_recording ?? current?.isMeetingRecording ?? '',
        ).toLowerCase() === 'true';
      const looksLikeVideoFile =
        ['mp4', 'm4v', 'mov', 'webm'].includes(extension) ||
        itemType === 'video' ||
        fileType === 'stream';
      const identityKey =
        normalizeInlineText(current?.resource_id || current?.resourceId || '') ||
        normalizeInlineText(current?.doc_id || current?.docId || '') ||
        normalizeInlineText(current?.file_id || current?.fileId || '') ||
        normalizeInlineText(current?.sharepoint_info?.unique_id || '') ||
        normalizeInlineText(current?.sharepointIds?.listItemUniqueId || '') ||
        normalizeInlineText(current?.sharepointIds?.driveItemId || '') ||
        normalizeInlineText(current?.sharepointIds?.itemId || '') ||
        normalizeInlineText(current?.onedrive_info?.item_id || '');

      if (isMeetingRecording || looksLikeVideoFile) {
        const candidateKeys = [
          'meeting_catchup_link',
          'meetingCatchupLink',
          'canonicalUrl',
          'canonical_url',
          'webUrl',
          'web_url',
          'url',
          'mru_url',
          'shareUrl',
          'share_url',
          'playerUrl',
          'player_url',
          'downloadUrl',
          'download_url',
        ];
        candidateKeys.forEach((key) => {
          if (typeof current[key] === 'string') {
            addMatch(
              current[key],
              { title, subtitle, identityKey },
              `${path}.${key}`,
            );
          }
        });
        return;
      }

      for (const [key, child] of Object.entries(current)) {
        const childPath = `${path}.${key}`;

        if (typeof child === 'string') {
          if (
            /url|uri|href|link|path|weburl|shareurl|watchurl|playerurl/i.test(key) ||
            isLikelyMeetingItemUrl(child)
          ) {
            addMatch(child, { title, subtitle }, childPath);
          } else {
            for (const candidateUrl of extractUrlsFromText(child)) {
              addMatch(candidateUrl, { title, subtitle }, childPath);
            }
          }
          continue;
        }

        walk(child, childPath, depth + 1);
      }
    }

    walk(value, '$');
    return matches;
  }

  function extractMeetingDiscoveryItemsFromBody(body, sourceUrl = '') {
    const trimmed = stripUtf8Bom(String(body || '')).trim();
    if (!trimmed) {
      return [];
    }

    const jsonValue = tryParseJsonLenient(trimmed);
    if (jsonValue != null) {
      return extractMeetingDiscoveryItemsFromValue(jsonValue, sourceUrl);
    }

    return extractUrlsFromText(trimmed)
      .map((candidateUrl) => ({
        url: normalizeMeetingUrl(candidateUrl),
        title: '',
        subtitle: '',
        sourceUrl,
        path: '$',
      }))
      .filter((item) => isLikelyMeetingItemUrl(item.url));
  }

  function normalizeMeetingComparisonTitle(value) {
    return normalizeInlineText(
      String(value || '')
        .replace(/^stream\s+/i, '')
        .replace(
          /-\d{8}_\d{6}-meeting\s+(?:recording|transcript)\.(?:mp4|m4v|mov|webm)$/i,
          '',
        )
        .replace(/\.(?:mp4|m4v|mov|webm)$/i, '')
        .replace(/\{placeholder\}/gi, 'placeholder')
        .replace(/[|/]+/g, ' ')
        .replace(/[^a-z0-9]+/gi, ' '),
    ).toLowerCase();
  }

  function buildMeetingComparisonKey(value) {
    return normalizeMeetingComparisonTitle(value).replace(/\s+/g, '');
  }

  function parseGraphThumbnailDriveItem(value) {
    const normalizedUrl = normalizeMeetingUrl(value);
    if (!normalizedUrl) {
      return null;
    }

    const match = normalizedUrl.match(
      /^https:\/\/graph\.microsoft\.com\/v1\.0\/drives\/([^/]+)\/items\/([^/]+)\/thumbnails\//i,
    );
    if (!match) {
      return null;
    }

    return {
      driveId: decodeURIComponent(match[1]),
      itemId: decodeURIComponent(match[2]),
      thumbnailUrl: normalizedUrl,
    };
  }

  function getGraphAuthorizationHeader(graphThumbnailCandidates) {
    if (
      !(graphThumbnailCandidates instanceof Map) ||
      graphThumbnailCandidates.size === 0
    ) {
      return '';
    }

    return (
      [...new Set(
        [...graphThumbnailCandidates.values()]
          .map((candidate) => String(candidate?.authorization || '').trim())
          .filter(Boolean),
      )][0] || ''
    );
  }

  function buildGraphItemMetadataUrl(driveId, itemId) {
    const metadataUrl = new URL(
      `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}`,
    );
    metadataUrl.searchParams.set(
      '$select',
      'id,name,webUrl,file,video,sharepointIds,parentReference,@microsoft.graph.downloadUrl',
    );
    return metadataUrl.toString();
  }

  async function fetchGraphItemMetadata(driveId, itemId, authorization) {
    if (!driveId || !itemId || !authorization) {
      return null;
    }

    let response;
    try {
      response = await fetch(buildGraphItemMetadataUrl(driveId, itemId), {
        headers: {
          Authorization: authorization,
          Accept: 'application/json',
        },
      });
    } catch {
      return null;
    }

    if (!response.ok) {
      return null;
    }

    try {
      return await response.json();
    } catch {
      return null;
    }
  }

  async function resolveGraphThumbnailMeetingItems(
    graphThumbnailCandidates,
    {
      visibleTitles = [],
      existingNetworkItems = [],
    } = {},
  ) {
    if (!(graphThumbnailCandidates instanceof Map) || graphThumbnailCandidates.size === 0) {
      return [];
    }

    const authorization = getGraphAuthorizationHeader(graphThumbnailCandidates);
    if (!authorization) {
      return [];
    }

    const visibleQueueByKey = new Map();
    for (const title of visibleTitles) {
      const key = buildMeetingComparisonKey(title);
      if (!key) {
        continue;
      }

      if (!visibleQueueByKey.has(key)) {
        visibleQueueByKey.set(key, []);
      }
      visibleQueueByKey.get(key).push(normalizeInlineText(title));
    }

    const existingNetworkCounts = new Map();
    for (const item of existingNetworkItems) {
      const key = buildMeetingComparisonKey(item?.title || '');
      if (!key) {
        continue;
      }

      existingNetworkCounts.set(
        key,
        (existingNetworkCounts.get(key) || 0) + 1,
      );
    }

    const outstandingVisibleByKey = new Map();
    for (const [key, titles] of visibleQueueByKey.entries()) {
      const alreadyCovered = existingNetworkCounts.get(key) || 0;
      const remainingTitles = titles.slice(alreadyCovered);
      if (remainingTitles.length > 0) {
        outstandingVisibleByKey.set(key, remainingTitles);
      }
    }

    if (outstandingVisibleByKey.size === 0) {
      return [];
    }

    const resolvedGraphItems = [];
    for (const candidate of graphThumbnailCandidates.values()) {
      const payload = await fetchGraphItemMetadata(
        candidate.driveId,
        candidate.itemId,
        authorization,
      );
      if (!payload) {
        continue;
      }

      const itemName = normalizeInlineText(payload?.name || '');
      const key = buildMeetingComparisonKey(itemName);
      if (!key || !outstandingVisibleByKey.has(key)) {
        continue;
      }

      const queue = outstandingVisibleByKey.get(key);
      const title = queue.shift() || itemName;
      if (queue.length === 0) {
        outstandingVisibleByKey.delete(key);
      }

      const targetUrl =
        normalizeMeetingUrl(
          payload?.webUrl ||
            payload?.['@microsoft.graph.downloadUrl'] ||
            '',
        ) || '';
      if (!targetUrl || !isLikelyMeetingItemUrl(targetUrl)) {
        continue;
      }

      const candidateUrls = uniqueNormalizedMeetingUrls([
        payload?.webUrl || '',
        payload?.['@microsoft.graph.downloadUrl'] || '',
        targetUrl,
      ]);
      const downloadUrls = normalizeMeetingDownloadUrls([
        payload?.['@microsoft.graph.downloadUrl'] || '',
        targetUrl,
      ]);

      resolvedGraphItems.push({
        url: targetUrl,
        title,
        subtitle: normalizeInlineText(
          payload?.parentReference?.path || '',
        ),
        identityKey:
          normalizeInlineText(payload?.sharepointIds?.listItemUniqueId || '') ||
          normalizeInlineText(payload?.id || ''),
        candidateUrls,
        downloadUrls,
        sourceUrl: normalizeMeetingUrl(candidate.thumbnailUrl),
        path: '$graphThumbnail',
      });
    }

    return resolvedGraphItems;
  }

  async function resolveDomGraphMeetingItems(
    domItems,
    {
      graphThumbnailCandidates,
      existingNetworkItems = [],
    } = {},
  ) {
    if (!Array.isArray(domItems) || domItems.length === 0) {
      return [];
    }

    const authorization = getGraphAuthorizationHeader(graphThumbnailCandidates);
    if (!authorization) {
      return [];
    }

    const existingNetworkUrls = new Set(
      existingNetworkItems
        .map((item) => normalizeMeetingUrl(item?.url || ''))
        .filter(Boolean),
    );
    const existingNetworkIdentityKeys = new Set(
      existingNetworkItems
        .map((item) => normalizeInlineText(item?.identityKey || ''))
        .filter(Boolean),
    );
    const requestedGraphItems = new Set();
    const resolvedGraphItems = [];

    for (const item of domItems) {
      const existingUrl = normalizeMeetingUrl(item?.url || '');
      const identityKey = normalizeInlineText(item?.identityKey || '');
      const driveId = normalizeInlineText(item?.graphDriveId || '');
      const itemId = normalizeInlineText(item?.graphItemId || '');
      if (!driveId || !itemId) {
        continue;
      }

      if (
        (existingUrl && existingNetworkUrls.has(existingUrl)) ||
        (identityKey && existingNetworkIdentityKeys.has(identityKey))
      ) {
        continue;
      }

      const requestKey = `${driveId}\t${itemId}`;
      if (requestedGraphItems.has(requestKey)) {
        continue;
      }
      requestedGraphItems.add(requestKey);

      const payload = await fetchGraphItemMetadata(
        driveId,
        itemId,
        authorization,
      );
      if (!payload) {
        continue;
      }

      const resolvedUrl =
        normalizeMeetingUrl(
          payload?.webUrl || payload?.['@microsoft.graph.downloadUrl'] || '',
        ) || existingUrl;
      if (
        !resolvedUrl ||
        (!isLikelyMeetingItemUrl(resolvedUrl) &&
          !isLikelyMeetingVideoFileUrl(resolvedUrl))
      ) {
        continue;
      }

      const candidateUrls = uniqueNormalizedMeetingUrls([
        existingUrl,
        payload?.webUrl || '',
        payload?.['@microsoft.graph.downloadUrl'] || '',
        ...(Array.isArray(item?.candidateUrls) ? item.candidateUrls : []),
      ]);
      const downloadUrls = normalizeMeetingDownloadUrls([
        ...(Array.isArray(item?.downloadUrls) ? item.downloadUrls : []),
        payload?.['@microsoft.graph.downloadUrl'] || '',
        resolvedUrl,
      ]);

      resolvedGraphItems.push({
        url: resolvedUrl,
        title:
          normalizeInlineText(item?.title || '') ||
          normalizeInlineText(payload?.name || '') ||
          resolvedUrl,
        subtitle:
          normalizeInlineText(item?.subtitle || '') ||
          normalizeInlineText(payload?.parentReference?.path || ''),
        identityKey:
          identityKey ||
          normalizeInlineText(payload?.sharepointIds?.listItemUniqueId || '') ||
          normalizeInlineText(payload?.id || ''),
        candidateUrls,
        downloadUrls,
        sourceUrl: buildGraphItemMetadataUrl(driveId, itemId),
        path: '$graphDomItem',
      });
      existingNetworkUrls.add(resolvedUrl);
      if (identityKey) {
        existingNetworkIdentityKeys.add(identityKey);
      }
    }

    return resolvedGraphItems;
  }

  function shouldTrackMeetingDiscoveryResponse(response, resourceType) {
    const url = String(response?.url || '').toLowerCase();
    if (!url || !isStreamRelatedUrl(url)) {
      return false;
    }

    const resource = String(resourceType || '');
    if (/Image|Media|Font|Stylesheet/.test(resource)) {
      return false;
    }

    const mimeType = String(response?.mimeType || '').toLowerCase();
    const contentType = String(
      lowerCaseHeaderMap(response?.headers || {})['content-type'] || '',
    ).toLowerCase();
    const contentHint = `${mimeType} ${contentType}`;

    return /json|text|html|xml|javascript/.test(contentHint);
  }

  async function discoverMeetingPages(
    cdp,
    startUrl,
    {
      navigate = true,
      initialPage = null,
      waitBeforeDomDiscoveryMs = 0,
      debugEnabled = false,
    } = {},
  ) {
    await cdp.send('Network.enable', {
      maxTotalBufferSize: 100_000_000,
      maxResourceBufferSize: 10_000_000,
    });
    await cdp
      .send('Network.setCacheDisabled', { cacheDisabled: true })
      .catch(() => {});
    await cdp
      .send('Network.setBypassServiceWorker', { bypass: true })
      .catch(() => {});

    const eventScope = createCdpEventScope(cdp);
    const trackedResponses = new Map();
    const finishedRequests = new Set();
    const failedRequests = new Set();
    const bodyCache = new Map();
    const requestMetadata = new Map();
    const graphThumbnailCandidates = new Map();

    eventScope.on('Network.requestWillBeSent', (params) => {
      requestMetadata.set(String(params.requestId), {
        requestId: String(params.requestId),
        url: String(params.request?.url || ''),
        method: String(params.request?.method || ''),
      });
    });

    eventScope.on('Network.responseReceived', (params) => {
      if (
        trackedResponses.size >= maxDebugResponses ||
        !shouldTrackMeetingDiscoveryResponse(params.response, params.type)
      ) {
        return;
      }

      const requestId = String(params.requestId);
      trackedResponses.set(requestId, {
        requestId,
        url: String(params.response.url || ''),
        status: params.response.status,
        mimeType: String(params.response.mimeType || ''),
        resourceType: String(params.type || ''),
      });
    });

    eventScope.on('Network.requestWillBeSentExtraInfo', (params) => {
      const requestId = String(params.requestId);
      const request = requestMetadata.get(requestId) || {};
      const graphTarget = parseGraphThumbnailDriveItem(request.url || '');
      if (!graphTarget || String(request.method || '').toUpperCase() !== 'GET') {
        return;
      }

      const headers = params.headers || {};
      const authorization = String(
        headers.authorization || headers.Authorization || '',
      ).trim();
      const key = `${graphTarget.driveId}\t${graphTarget.itemId}`;
      if (!graphThumbnailCandidates.has(key)) {
        graphThumbnailCandidates.set(key, {
          ...graphTarget,
          authorization,
        });
        return;
      }

      const existing = graphThumbnailCandidates.get(key);
      if (!existing.authorization && authorization) {
        existing.authorization = authorization;
      }
    });

    eventScope.on('Network.loadingFinished', (params) => {
      finishedRequests.add(String(params.requestId));
    });

    eventScope.on('Network.loadingFailed', (params) => {
      failedRequests.add(String(params.requestId));
    });

    try {
      const pageSnapshot = navigate
        ? await navigatePageAndWait(cdp, startUrl)
        : initialPage ||
          (await getCurrentPageSnapshot(cdp).catch(() => ({
            title: '',
            url: startUrl,
          })));

      if (waitBeforeDomDiscoveryMs > 0) {
        await sleep(waitBeforeDomDiscoveryMs);
      }

      const domDiscovery = await evaluate(
        cdp,
        buildMeetingDiscoveryExpression(),
        defaultPageNavigationTimeoutMs * 2,
      );

      if (domDiscovery?.error) {
        throw errorFactory(domDiscovery.error);
      }

      eventScope.stopAll();

      const networkItems = [];
      const debugResponseBodies = [];
      const bodyTargets = Array.from(trackedResponses.values())
        .filter(
          (responseRecord) =>
            finishedRequests.has(responseRecord.requestId) &&
            !failedRequests.has(responseRecord.requestId),
        )
        .slice(0, maxDiscoveryResponseBodies);

      for (const responseRecord of bodyTargets) {
        const bodyRecord = await loadResponseBody(
          cdp,
          responseRecord.requestId,
          bodyCache,
        );

        if (bodyRecord.bodyError || !bodyRecord.body) {
          if (debugEnabled) {
            debugResponseBodies.push({
              ...responseRecord,
              bodyError: bodyRecord.bodyError,
              bodyPreview: bodyRecord.bodyPreview,
              matchedItems: [],
            });
          }
          continue;
        }

        const extractedItems = extractMeetingDiscoveryItemsFromBody(
          bodyRecord.body,
          responseRecord.url,
        );
        networkItems.push(...extractedItems);

        if (debugEnabled) {
          debugResponseBodies.push({
            ...responseRecord,
            bodyError: bodyRecord.bodyError,
            bodyPreview: bodyRecord.bodyPreview,
            matchedItems: extractedItems,
          });
        }
      }

      const graphResolvedItems = await resolveGraphThumbnailMeetingItems(
        graphThumbnailCandidates,
        {
          visibleTitles: Array.isArray(domDiscovery?.visibleCardTitles)
            ? domDiscovery.visibleCardTitles
            : [],
          existingNetworkItems: networkItems,
        },
      );
      networkItems.push(...graphResolvedItems);

      const domGraphResolvedItems = await resolveDomGraphMeetingItems(
        Array.isArray(domDiscovery?.items) ? domDiscovery.items : [],
        {
          graphThumbnailCandidates,
          existingNetworkItems: networkItems,
        },
      );
      networkItems.push(...domGraphResolvedItems);

      const mergedItems = mergeDiscoveredMeetingItems(
        networkItems,
        Array.isArray(domDiscovery?.items) ? domDiscovery.items : [],
      );

      const totalGraphResolvedItemCount =
        graphResolvedItems.length + domGraphResolvedItems.length;

      return {
        startUrl,
        page: domDiscovery?.page || pageSnapshot,
        meetingsControlLabel: String(domDiscovery?.controlLabel || ''),
        meetingsSelectedBeforeClick: Boolean(domDiscovery?.selectedBeforeClick),
        scrollIterations:
          typeof domDiscovery?.scrollIterations === 'number'
            ? domDiscovery.scrollIterations
            : 0,
        trackedResponseCount: trackedResponses.size,
        networkItemCount: networkItems.length,
        graphResolvedItemCount: totalGraphResolvedItemCount,
        graphThumbnailResolvedItemCount: graphResolvedItems.length,
        domGraphResolvedItemCount: domGraphResolvedItems.length,
        domItemCount: Array.isArray(domDiscovery?.items)
          ? domDiscovery.items.length
          : 0,
        visibleCardCount:
          typeof domDiscovery?.visibleCardCount === 'number'
            ? domDiscovery.visibleCardCount
            : Array.isArray(domDiscovery?.visibleCardTitles)
              ? domDiscovery.visibleCardTitles.length
              : 0,
        visibleCardTitles: Array.isArray(domDiscovery?.visibleCardTitles)
          ? domDiscovery.visibleCardTitles
          : [],
        items: mergedItems,
        debug: debugEnabled
          ? {
              trackedResponses: Array.from(trackedResponses.values()),
              responseBodies: debugResponseBodies,
              graphThumbnailCandidateCount: graphThumbnailCandidates.size,
              graphResolvedItemCount: totalGraphResolvedItemCount,
              graphThumbnailResolvedItemCount: graphResolvedItems.length,
              domGraphResolvedItemCount: domGraphResolvedItems.length,
            }
          : null,
      };
    } finally {
      eventScope.stopAll();
      await cdp
        .send('Network.setBypassServiceWorker', { bypass: false })
        .catch(() => {});
      await cdp
        .send('Network.setCacheDisabled', { cacheDisabled: false })
        .catch(() => {});
      await cdp.send('Network.disable').catch(() => {});
    }
  }

  return {
    discoverMeetingPages,
  };
}
