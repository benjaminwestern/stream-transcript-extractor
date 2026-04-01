import { existsSync, readFileSync } from 'node:fs';

export function createCrawlerStateRuntime({
  stateVersion,
  appName,
  buildVersion,
  buildTime,
  platform,
  defaultOutputDir,
  resolveOutputDirectory,
  normalizeMeetingUrl,
  normalizeMeetingDownloadUrls,
  uniqueNormalizedMeetingUrls,
  sortMeetingUrlsForExtraction,
  normalizeInlineText,
  errorFactory,
}) {
  function uniqueStrings(values) {
    return [
      ...new Set(
        (values || []).map((value) => String(value || '').trim()).filter(Boolean),
      ),
    ];
  }

  function loadCrawlerState(stateFilePath) {
    if (!existsSync(stateFilePath)) {
      return {
        stateVersion,
        items: [],
      };
    }

    try {
      const payload = JSON.parse(readFileSync(stateFilePath, 'utf-8'));
      return {
        stateVersion:
          typeof payload?.stateVersion === 'number'
            ? payload.stateVersion
            : stateVersion,
        items: Array.isArray(payload?.items) ? payload.items : [],
      };
    } catch (error) {
      throw errorFactory(
        `Unable to read crawl state file "${stateFilePath}": ` +
          `${error instanceof Error ? error.message : 'Unknown parse failure.'}`,
      );
    }
  }

  function getCrawlerItemProgress(item) {
    if (item?.isNewThisRun && Number(item?.attemptCount || 0) === 0) {
      return 'new';
    }

    const status = String(item?.lastStatus || '').toLowerCase();
    if (status === 'success') {
      return 'done';
    }
    if (status === 'failed') {
      return 'failed';
    }

    return 'pending';
  }

  function getCrawlerItemLookupUrls(item) {
    return uniqueNormalizedMeetingUrls([
      item?.url || '',
      ...(Array.isArray(item?.candidateUrls) ? item.candidateUrls : []),
    ]);
  }

  function buildCrawlerExtractionTargetUrls(item) {
    return sortMeetingUrlsForExtraction([
      ...(Array.isArray(item?.candidateUrls) ? item.candidateUrls : []),
      item?.url || '',
    ]);
  }

  function buildCrawlerDownloadTargetUrls(item) {
    return normalizeMeetingDownloadUrls([
      ...(Array.isArray(item?.downloadUrls) ? item.downloadUrls : []),
      ...(Array.isArray(item?.candidateUrls) ? item.candidateUrls : []),
      item?.url || '',
    ]);
  }

  function normalizeCrawlerStateItem(item, fallbackDiscoveredAt = '') {
    const candidateUrls = sortMeetingUrlsForExtraction(
      getCrawlerItemLookupUrls(item),
    );
    const url = candidateUrls[0] || normalizeMeetingUrl(item?.url);
    const downloadUrls = normalizeMeetingDownloadUrls([
      ...(Array.isArray(item?.downloadUrls) ? item.downloadUrls : []),
      url,
      ...candidateUrls,
    ]);

    return {
      url,
      identityKey: normalizeInlineText(item?.identityKey || ''),
      title: String(item?.title || ''),
      subtitle: String(item?.subtitle || ''),
      candidateUrls,
      downloadUrls,
      sources: uniqueStrings(item?.sources),
      discoveryPaths: uniqueStrings(item?.discoveryPaths),
      discoverySourceUrls: uniqueStrings(item?.discoverySourceUrls),
      firstDiscoveredAt: String(item?.firstDiscoveredAt || fallbackDiscoveredAt),
      lastDiscoveredAt: String(
        item?.lastDiscoveredAt || item?.firstDiscoveredAt || fallbackDiscoveredAt,
      ),
      discoveryCount: Number.isInteger(item?.discoveryCount) ? item.discoveryCount : 0,
      lastStatus: String(item?.lastStatus || 'pending'),
      attemptCount: Number.isInteger(item?.attemptCount) ? item.attemptCount : 0,
      successCount: Number.isInteger(item?.successCount) ? item.successCount : 0,
      failureCount: Number.isInteger(item?.failureCount) ? item.failureCount : 0,
      lastSelectedAt: String(item?.lastSelectedAt || ''),
      lastAttemptedAt: String(item?.lastAttemptedAt || ''),
      lastSucceededAt: String(item?.lastSucceededAt || ''),
      lastFailedAt: String(item?.lastFailedAt || ''),
      lastError: String(item?.lastError || ''),
      lastEntryCount: Number.isInteger(item?.lastEntryCount) ? item.lastEntryCount : 0,
      lastMatchedCandidateUrl: String(item?.lastMatchedCandidateUrl || ''),
      outputPaths: Array.isArray(item?.outputPaths) ? item.outputPaths.map(String) : [],
      networkOutputPath: String(item?.networkOutputPath || ''),
      isNewThisRun: Boolean(item?.isNewThisRun),
      seenInCurrentDiscovery: Boolean(item?.seenInCurrentDiscovery),
    };
  }

  function mergeCrawlerStateItems(existingItems, discoveredItems, discoveredAt) {
    const existingByUrl = new Map();
    const existingByIdentityKey = new Map();
    const normalizedExistingItems = [];

    for (const item of existingItems) {
      const normalized = normalizeCrawlerStateItem(item);
      if (!normalized.url) {
        continue;
      }

      normalized.isNewThisRun = false;
      normalized.seenInCurrentDiscovery = false;
      normalizedExistingItems.push(normalized);

      for (const lookupUrl of getCrawlerItemLookupUrls(normalized)) {
        if (!existingByUrl.has(lookupUrl)) {
          existingByUrl.set(lookupUrl, normalized);
        }
      }

      if (normalized.identityKey && !existingByIdentityKey.has(normalized.identityKey)) {
        existingByIdentityKey.set(normalized.identityKey, normalized);
      }
    }

    const mergedItems = [];
    const matchedExistingItems = new Set();
    const seenDiscoveryKeys = new Set();

    for (const discoveryItem of discoveredItems) {
      const identityKey = normalizeInlineText(discoveryItem?.identityKey || '');
      const discoveryLookupUrls = getCrawlerItemLookupUrls(discoveryItem);
      const discoveryKey = identityKey || discoveryLookupUrls[0] || '';
      if (!discoveryKey || seenDiscoveryKeys.has(discoveryKey)) {
        continue;
      }
      seenDiscoveryKeys.add(discoveryKey);

      let existingItem = identityKey
        ? existingByIdentityKey.get(identityKey) || null
        : null;
      if (!existingItem) {
        for (const lookupUrl of discoveryLookupUrls) {
          existingItem = existingByUrl.get(lookupUrl) || null;
          if (existingItem) {
            break;
          }
        }
      }

      if (existingItem) {
        matchedExistingItems.add(existingItem);
      }

      const mergedItem = normalizeCrawlerStateItem(
        {
          ...existingItem,
          ...discoveryItem,
          identityKey: identityKey || existingItem?.identityKey || '',
          title: discoveryItem?.title || existingItem?.title || '',
          subtitle: discoveryItem?.subtitle || existingItem?.subtitle || '',
          candidateUrls: sortMeetingUrlsForExtraction([
            ...(existingItem?.candidateUrls || []),
            ...discoveryLookupUrls,
          ]),
          downloadUrls: normalizeMeetingDownloadUrls([
            ...(existingItem?.downloadUrls || []),
            ...(Array.isArray(discoveryItem?.downloadUrls)
              ? discoveryItem.downloadUrls
              : []),
            ...(Array.isArray(discoveryItem?.candidateUrls)
              ? discoveryItem.candidateUrls
              : []),
            discoveryItem?.url || '',
          ]),
          sources: uniqueStrings([
            ...(existingItem?.sources || []),
            ...(discoveryItem?.sources || []),
          ]),
          discoveryPaths: uniqueStrings([
            ...(existingItem?.discoveryPaths || []),
            ...(discoveryItem?.discoveryPaths || []),
          ]),
          discoverySourceUrls: uniqueStrings([
            ...(existingItem?.discoverySourceUrls || []),
            ...(discoveryItem?.discoverySourceUrls || []),
          ]),
          firstDiscoveredAt: existingItem?.firstDiscoveredAt || discoveredAt,
          lastDiscoveredAt: discoveredAt,
          discoveryCount: Math.max(0, Number(existingItem?.discoveryCount || 0)) + 1,
          lastStatus: existingItem?.lastStatus || 'pending',
          isNewThisRun: !existingItem,
          seenInCurrentDiscovery: true,
        },
        discoveredAt,
      );

      mergedItems.push(mergedItem);
    }

    for (const item of normalizedExistingItems) {
      if (matchedExistingItems.has(item)) {
        continue;
      }

      mergedItems.push({
        ...item,
        isNewThisRun: false,
        seenInCurrentDiscovery: false,
      });
    }

    return mergedItems;
  }

  function buildCrawlerStateSummary(items) {
    const summary = {
      totalItemCount: items.length,
      seenInCurrentDiscoveryCount: 0,
      newItemCount: 0,
      pendingItemCount: 0,
      failedItemCount: 0,
      successItemCount: 0,
    };

    for (const item of items) {
      const progress = getCrawlerItemProgress(item);
      if (item?.seenInCurrentDiscovery) {
        summary.seenInCurrentDiscoveryCount += 1;
      }

      if (progress === 'new') {
        summary.newItemCount += 1;
      } else if (progress === 'failed') {
        summary.failedItemCount += 1;
      } else if (progress === 'done') {
        summary.successItemCount += 1;
      } else {
        summary.pendingItemCount += 1;
      }
    }

    return summary;
  }

  function buildCrawlerStatePayload({
    options,
    browser,
    profile,
    stateFilePath,
    items,
    updatedAt,
  }) {
    return {
      app: {
        name: appName,
        version: buildVersion,
        buildTime,
        platform,
      },
      stateVersion,
      updatedAt,
      stateFilePath,
      options: {
        startUrl: options.startUrl,
        outputDir: resolveOutputDirectory(options.outputDir, {
          defaultOutputDir,
        }),
        outputFormat: options.outputFormat,
        debug: options.debug,
      },
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
      summary: buildCrawlerStateSummary(items),
      items: items.map((item) => {
        const normalizedItem = normalizeCrawlerStateItem(item);
        return {
          ...normalizedItem,
          isNewThisRun: undefined,
          seenInCurrentDiscovery: undefined,
        };
      }),
    };
  }

  return {
    loadCrawlerState,
    getCrawlerItemProgress,
    mergeCrawlerStateItems,
    buildCrawlerStateSummary,
    buildCrawlerStatePayload,
    buildCrawlerExtractionTargetUrls,
    buildCrawlerDownloadTargetUrls,
  };
}
