export function createMeetingDiscoveryRuntime({
  normalizeMeetingUrl,
  uniqueNormalizedMeetingUrls,
  normalizeMeetingDownloadUrls,
  sortMeetingUrlsForExtraction,
  normalizeInlineText,
  scoreMeetingTargetUrl,
}) {
  function mergeDiscoveredMeetingItems(networkItems = [], domItems = []) {
    const mergedItems = new Map();

    function buildCandidateUrls(item) {
      return uniqueNormalizedMeetingUrls([
        item?.url || '',
        ...(Array.isArray(item?.candidateUrls) ? item.candidateUrls : []),
      ]);
    }

    function buildDownloadUrls(item) {
      return normalizeMeetingDownloadUrls([
        ...(Array.isArray(item?.downloadUrls) ? item.downloadUrls : []),
        ...buildCandidateUrls(item),
      ]);
    }

    function buildCanonicalUrl(item) {
      const sortedCandidateUrls = sortMeetingUrlsForExtraction(
        buildCandidateUrls(item),
      );
      return sortedCandidateUrls[0] || normalizeMeetingUrl(item?.url || '');
    }

    function buildMergeLookupUrls(item) {
      return uniqueNormalizedMeetingUrls([
        normalizeMeetingUrl(item?.url || ''),
        buildCanonicalUrl(item),
      ]);
    }

    function findExistingMergedKey(lookupUrls, identityKey = '') {
      if (identityKey) {
        for (const [existingKey, existingItem] of mergedItems.entries()) {
          if (existingItem.identityKey === identityKey) {
            return existingKey;
          }
        }
      }

      for (const [existingKey, existingItem] of mergedItems.entries()) {
        const existingLookupUrls = uniqueNormalizedMeetingUrls([
          existingItem.url,
          ...(Array.isArray(existingItem.mergeLookupUrls)
            ? existingItem.mergeLookupUrls
            : []),
        ]);
        if (lookupUrls.some((lookupUrl) => existingLookupUrls.includes(lookupUrl))) {
          return existingKey;
        }
      }

      return '';
    }

    function updateMergedItemUrl(mergedItem) {
      const candidateUrls = uniqueNormalizedMeetingUrls(mergedItem.candidateUrls || []);
      mergedItem.candidateUrls = candidateUrls;

      let preferredUrl = mergedItem.url;
      let preferredUrlScore = scoreMeetingTargetUrl(preferredUrl);
      for (const candidateUrl of candidateUrls) {
        const candidateScore = scoreMeetingTargetUrl(candidateUrl);
        if (candidateScore > preferredUrlScore) {
          preferredUrl = candidateUrl;
          preferredUrlScore = candidateScore;
        }
      }

      mergedItem.url = preferredUrl || candidateUrls[0] || mergedItem.url;
      mergedItem.preferredUrlScore = preferredUrlScore;
      mergedItem.mergeLookupUrls = uniqueNormalizedMeetingUrls([
        ...(Array.isArray(mergedItem.mergeLookupUrls)
          ? mergedItem.mergeLookupUrls
          : []),
        mergedItem.url,
      ]);
    }

    function ensureMergedItem(item) {
      const identityKey = normalizeInlineText(item?.identityKey || '');
      const lookupUrls = buildMergeLookupUrls(item);
      const candidateUrls = buildCandidateUrls(item);
      const normalizedUrl = lookupUrls[0] || normalizeMeetingUrl(item?.url || '');
      if (!normalizedUrl) {
        return null;
      }

      let mergedKey = findExistingMergedKey(lookupUrls, identityKey);
      if (!mergedKey) {
        mergedKey = identityKey || normalizedUrl;
      }

      if (!mergedItems.has(mergedKey)) {
        mergedItems.set(mergedKey, {
          url: normalizedUrl,
          identityKey,
          title: '',
          subtitle: '',
          candidateUrls: candidateUrls.length > 0 ? candidateUrls : [normalizedUrl],
          downloadUrls: buildDownloadUrls(item),
          mergeLookupUrls: lookupUrls,
          sources: [],
          discoveryPaths: [],
          discoverySourceUrls: [],
          domOrder: Number.POSITIVE_INFINITY,
          preferredUrlScore: scoreMeetingTargetUrl(normalizedUrl),
        });
      }

      const mergedItem = mergedItems.get(mergedKey);
      if (!mergedItem.identityKey && identityKey) {
        mergedItem.identityKey = identityKey;
      }
      mergedItem.candidateUrls = uniqueNormalizedMeetingUrls([
        ...(mergedItem.candidateUrls || []),
        ...candidateUrls,
      ]);
      mergedItem.downloadUrls = normalizeMeetingDownloadUrls([
        ...(mergedItem.downloadUrls || []),
        ...buildDownloadUrls(item),
      ]);
      mergedItem.mergeLookupUrls = uniqueNormalizedMeetingUrls([
        ...(Array.isArray(mergedItem.mergeLookupUrls)
          ? mergedItem.mergeLookupUrls
          : []),
        ...lookupUrls,
      ]);
      updateMergedItemUrl(mergedItem);

      return mergedItem;
    }

    domItems.forEach((item, index) => {
      const mergedItem = ensureMergedItem(item);
      if (!mergedItem) {
        return;
      }

      const title = normalizeInlineText(item?.title || '');
      const subtitle = normalizeInlineText(item?.subtitle || '');
      mergedItem.title = title || mergedItem.title;
      mergedItem.subtitle = subtitle || mergedItem.subtitle;
      mergedItem.candidateUrls = uniqueNormalizedMeetingUrls([
        ...(mergedItem.candidateUrls || []),
        ...(Array.isArray(item?.candidateUrls) ? item.candidateUrls : []),
      ]);
      mergedItem.downloadUrls = normalizeMeetingDownloadUrls([
        ...(mergedItem.downloadUrls || []),
        ...(Array.isArray(item?.downloadUrls) ? item.downloadUrls : []),
      ]);
      mergedItem.domOrder = Math.min(mergedItem.domOrder, index);
      mergedItem.sources = [...new Set([...mergedItem.sources, 'dom'])];
    });

    networkItems.forEach((item) => {
      const mergedItem = ensureMergedItem(item);
      if (!mergedItem) {
        return;
      }

      const title = normalizeInlineText(item?.title || '');
      const subtitle = normalizeInlineText(item?.subtitle || '');
      if (!mergedItem.title) {
        mergedItem.title = title;
      }
      if (!mergedItem.subtitle) {
        mergedItem.subtitle = subtitle;
      }
      mergedItem.candidateUrls = uniqueNormalizedMeetingUrls([
        ...(mergedItem.candidateUrls || []),
        ...(Array.isArray(item?.candidateUrls) ? item.candidateUrls : []),
      ]);
      mergedItem.downloadUrls = normalizeMeetingDownloadUrls([
        ...(mergedItem.downloadUrls || []),
        ...(Array.isArray(item?.downloadUrls) ? item.downloadUrls : []),
      ]);
      mergedItem.sources = [...new Set([...mergedItem.sources, 'network'])];

      const path = String(item?.path || '').trim();
      if (path && !mergedItem.discoveryPaths.includes(path)) {
        mergedItem.discoveryPaths.push(path);
      }

      const sourceUrl = normalizeMeetingUrl(item?.sourceUrl || '');
      if (sourceUrl && !mergedItem.discoverySourceUrls.includes(sourceUrl)) {
        mergedItem.discoverySourceUrls.push(sourceUrl);
      }
    });

    return [...mergedItems.values()]
      .sort((left, right) => {
        if (left.domOrder !== right.domOrder) {
          return left.domOrder - right.domOrder;
        }

        return (left.title || left.url).localeCompare(right.title || right.url);
      })
      .map((item) => {
        const candidateUrls = sortMeetingUrlsForExtraction(
          item.candidateUrls || [item.url],
        );
        return {
          url: candidateUrls[0] || item.url,
          identityKey: item.identityKey,
          title: item.title || item.url,
          subtitle: item.subtitle,
          candidateUrls,
          downloadUrls: normalizeMeetingDownloadUrls(item.downloadUrls || []),
          sources: item.sources,
          discoveryPaths: item.discoveryPaths,
          discoverySourceUrls: item.discoverySourceUrls,
        };
      });
  }

  return {
    mergeDiscoveredMeetingItems,
  };
}
