export function buildMeetingDiscoveryExpression({
  settleMs,
  defaultCrawlStartUrl,
}) {
  const serializedSettleMs = JSON.stringify(settleMs);
  const serializedDefaultCrawlStartUrl = JSON.stringify(
    String(defaultCrawlStartUrl || '').toLowerCase(),
  );

  return `
    (async () => {
      const settleMs = ${serializedSettleMs};

      function wait(ms) {
        return new Promise((resolveWait) => setTimeout(resolveWait, ms));
      }

      function normalizeText(value) {
        return String(value || '').replace(/\\s+/g, ' ').trim();
      }

      function isVisible(element) {
        if (!(element instanceof HTMLElement)) {
          return false;
        }

        const style = window.getComputedStyle(element);
        const rect = element.getBoundingClientRect();
        return (
          style.display !== 'none' &&
          style.visibility !== 'hidden' &&
          rect.width > 8 &&
          rect.height > 8
        );
      }

      function normalizeHref(value) {
        try {
          return new URL(String(value || ''), location.href).href;
        } catch {
          return '';
        }
      }

      function normalizeBoolean(value) {
        if (typeof value === 'boolean') {
          return value;
        }

        const normalized = normalizeText(value).toLowerCase();
        return normalized === 'true' || normalized === '1' || normalized === 'yes';
      }

      function isLikelyMeetingHref(value) {
        const href = normalizeHref(value).toLowerCase();
        if (!href) {
          return false;
        }

        if (
          href === ${serializedDefaultCrawlStartUrl} ||
          /[?&]home=1(?:&|$)/.test(href)
        ) {
          return false;
        }

        if (!/(stream|sharepoint|m365\\.cloud\\.microsoft)/.test(href)) {
          if (!/(meetingcatchupportal|cortana\\.ai)/.test(href)) {
            return false;
          }
        }

        return /meetingcatchupportal|meetingcatchup|stream\\.aspx|recording|watch|oneplayer|clip|\\/recordings?\\/.*\\.(mp4|m4v|mov|webm)(?:$|[?#])/.test(
          href,
        );
      }

      function scoreMeetingHref(value) {
        const href = normalizeHref(value).toLowerCase();
        if (!href) {
          return 0;
        }

        if (/meetingcatchupportal|meetingcatchup/.test(href)) {
          return 4;
        }

        if (/\\/recordings?\\/.*\\.(mp4|m4v|mov|webm)(?:$|[?#])/.test(href)) {
          return 3;
        }

        if (/stream\\.aspx|watch|oneplayer|clip/.test(href)) {
          return 2;
        }

        return isLikelyMeetingHref(href) ? 1 : 0;
      }

      function pickBestMeetingHref(values) {
        let bestHref = '';
        let bestScore = 0;

        for (const value of values) {
          const href = normalizeHref(value);
          const score = scoreMeetingHref(href);
          if (!href || score <= 0) {
            continue;
          }

          if (!bestHref || score > bestScore) {
            bestHref = href;
            bestScore = score;
          }
        }

        return bestHref;
      }

      function looksActive(element) {
        const attributeState =
          String(
            element.getAttribute('aria-selected') ||
              element.getAttribute('aria-pressed') ||
              '',
          ).toLowerCase() === 'true';
        if (attributeState) {
          return true;
        }

        return /active|selected|current/.test(
          [
            element.className,
            element.getAttribute('data-is-active') || '',
            element.getAttribute('data-active') || '',
          ]
            .join(' ')
            .toLowerCase(),
        );
      }

      function buildControlLabel(element) {
        return normalizeText(
          element.getAttribute('aria-label') ||
            element.textContent ||
            element.innerText ||
            '',
        );
      }

      function findMeetingsControl() {
        const candidates = [
          ...document.querySelectorAll('button, [role="tab"], [role="button"], a'),
        ].filter(isVisible);

        const matches = candidates
          .map((element) => ({
            element,
            label: buildControlLabel(element),
          }))
          .filter((candidate) => /^meetings$/i.test(candidate.label));

        matches.sort((left, right) => {
          const leftScore = looksActive(left.element) ? 1 : 0;
          const rightScore = looksActive(right.element) ? 1 : 0;
          return rightScore - leftScore;
        });

        return matches[0] || null;
      }

      function clickElement(element) {
        element.dispatchEvent(
          new MouseEvent('mousedown', { bubbles: true, cancelable: true }),
        );
        element.dispatchEvent(
          new MouseEvent('mouseup', { bubbles: true, cancelable: true }),
        );
        element.click();
      }

      function getTextLines(element) {
        return String(element?.innerText || '')
          .split(/\\r?\\n/)
          .map((line) => normalizeText(line))
          .filter(Boolean);
      }

      function getReactAttachedValue(element, prefixes) {
        if (!element || typeof element !== 'object') {
          return null;
        }

        for (const key of Object.keys(element)) {
          if (prefixes.some((prefix) => key.startsWith(prefix))) {
            return element[key];
          }
        }

        return null;
      }

      function findReactItemFromFiber(rootFiber) {
        if (!rootFiber || typeof rootFiber !== 'object') {
          return null;
        }

        const seen = new Set();
        const stack = [rootFiber];

        while (stack.length > 0) {
          const fiber = stack.pop();
          if (!fiber || typeof fiber !== 'object' || seen.has(fiber)) {
            continue;
          }

          seen.add(fiber);

          const propsCandidates = [fiber.pendingProps, fiber.memoizedProps];
          for (const props of propsCandidates) {
            if (props && typeof props.item === 'object' && props.item) {
              return props.item;
            }
          }

          if (fiber.child) {
            stack.push(fiber.child);
          }
          if (fiber.sibling) {
            stack.push(fiber.sibling);
          }
        }

        return null;
      }

      function findReactItemForElement(element) {
        const directProps = getReactAttachedValue(element, [
          '__reactProps$',
          '__reactProps',
        ]);
        if (directProps && typeof directProps.item === 'object' && directProps.item) {
          return directProps.item;
        }

        const fiber = getReactAttachedValue(element, [
          '__reactFiber$',
          '__reactFiber',
          '__reactInternalInstance$',
        ]);
        return findReactItemFromFiber(fiber);
      }

      function findReactItemForCard(card) {
        const descendants = [card, ...card.querySelectorAll('*')];
        const maxNodesToInspect = 80;
        for (let index = 0; index < descendants.length && index < maxNodesToInspect; index += 1) {
          const item = findReactItemForElement(descendants[index]);
          if (item && typeof item === 'object') {
            return item;
          }
        }

        return null;
      }

      function firstNormalizedString(values) {
        for (const value of values) {
          const normalized = normalizeText(value);
          if (normalized) {
            return normalized;
          }
        }

        return '';
      }

      function buildCardIdentityKey(item) {
        return firstNormalizedString([
          item?.resourceId,
          item?.resource_id,
          item?.docId,
          item?.doc_id,
          item?.fileId,
          item?.file_id,
          item?.sharepointIds?.listItemUniqueId,
          item?.sharepointIds?.driveItemId,
          item?.sharepointIds?.itemId,
          item?.sharepoint_info?.unique_id,
          item?.onedrive_info?.item_id,
        ]);
      }

      function buildCardGraphDriveId(item) {
        return firstNormalizedString([
          item?.driveId,
          item?.drive_id,
        ]);
      }

      function buildCardGraphItemId(item) {
        return firstNormalizedString([
          item?.docId,
          item?.doc_id,
          item?.sharepointIds?.driveItemId,
          item?.sharepointIds?.itemId,
        ]);
      }

      function buildCardUrlCandidates(item) {
        return [
          item?.meetingCatchupLink,
          item?.meetingCatchUpLink,
          item?.meeting_catchup_link,
          item?.canonicalUrl,
          item?.canonical_url,
          item?.webUrl,
          item?.web_url,
          item?.url,
          item?.mru_url,
          item?.shareUrl,
          item?.share_url,
          item?.playerUrl,
          item?.player_url,
          item?.downloadUrl,
          item?.download_url,
        ]
          .map((value) => normalizeHref(value))
          .filter(Boolean);
      }

      function sortExtractionCandidates(values) {
        const uniqueValues = [...new Set((values || []).map((value) => normalizeHref(value)).filter(Boolean))];
        function scoreExtractionHref(value) {
          const href = normalizeHref(value).toLowerCase();
          if (!href) {
            return 0;
          }

          if (/stream\\.aspx|watch|oneplayer|clip/.test(href)) {
            return 4;
          }

          if (/\\.(mp4|m4v|mov|webm)(?:$|[?#])/.test(href)) {
            return 3;
          }

          if (/meetingcatchupportal|meetingcatchup/.test(href)) {
            return 2;
          }

          return isLikelyMeetingHref(href) ? 1 : 0;
        }

        uniqueValues.sort((left, right) => scoreExtractionHref(right) - scoreExtractionHref(left));
        return uniqueValues;
      }

      function stripStreamPrefix(value) {
        return normalizeText(String(value || '').replace(/^stream\\s+/i, ''));
      }

      function buildCardTitle(card, item, textLines) {
        return firstNormalizedString([
          item?.title,
          item?.name,
          item?.displayName,
          item?.meetingTitle,
          item?.subject,
          item?.fileName,
          stripStreamPrefix(card.getAttribute('aria-label') || ''),
          textLines[0],
        ]);
      }

      function buildCardSubtitle(item, textLines, title) {
        const structuredSubtitle = firstNormalizedString([
          item?.sharedBy,
          item?.sharedByName,
          item?.createdBy,
          item?.createdByName,
          item?.ownerName,
          item?.description,
        ]);
        if (structuredSubtitle) {
          return structuredSubtitle;
        }

        const normalizedTitle = normalizeText(title).toLowerCase();
        const remainingLines = textLines.filter((line) => {
          const normalizedLine = normalizeText(line).toLowerCase();
          return normalizedLine && normalizedLine !== normalizedTitle;
        });
        return remainingLines.slice(0, 2).join(' · ');
      }

      function isMeetingRecordingItem(item, urlCandidates) {
        const extension = firstNormalizedString([
          item?.extension,
          item?.fileExtension,
          item?.file_extension,
        ]).toLowerCase();
        const itemType = firstNormalizedString([item?.type]).toLowerCase();
        const fileType = firstNormalizedString([
          item?.fileType,
          item?.file_type,
          item?.app,
        ]).toLowerCase();
        const isMeetingRecording = normalizeBoolean(
          item?.isMeetingRecording ?? item?.is_meeting_recording ?? '',
        );

        if (isMeetingRecording) {
          return true;
        }

        if (['mp4', 'm4v', 'mov', 'webm'].includes(extension)) {
          return true;
        }

        if (itemType === 'video' || fileType === 'stream') {
          return true;
        }

        return urlCandidates.some((href) => isLikelyMeetingHref(href));
      }

      function findRenderedCardsRoot() {
        return (
          document.querySelector('[class*="BaseEdgeworthControl-module__cards__"]') ||
          document.querySelector('[class*="cards__"]') ||
          null
        );
      }

      function getRenderedCards(container) {
        const primaryCards = [
          ...container.querySelectorAll('.fui-Card[role="group"]'),
        ].filter(isVisible);

        if (primaryCards.length > 0) {
          return primaryCards;
        }

        return [...container.querySelectorAll('.fui-Card')].filter((card) => {
          if (!isVisible(card)) {
            return false;
          }

          const ariaLabel = normalizeText(card.getAttribute('aria-label') || '');
          return /^stream\\s+/i.test(ariaLabel);
        });
      }

      function collectRenderedCardItems(container) {
        const discovered = [];
        const seenItems = new Set();
        const visibleCardTitles = [];
        const cards = getRenderedCards(container);

        for (const card of cards) {
          const reactItem = findReactItemForCard(card);
          const textLines = getTextLines(card);
          const title = buildCardTitle(card, reactItem || {}, textLines);
          if (title) {
            visibleCardTitles.push(title);
          }

          const urlCandidates = reactItem ? buildCardUrlCandidates(reactItem) : [];
          const url = pickBestMeetingHref(urlCandidates);
          if (!url) {
            continue;
          }

          if (!isMeetingRecordingItem(reactItem || {}, urlCandidates)) {
            continue;
          }

          const identityKey = buildCardIdentityKey(reactItem || {});
          const dedupeKey = identityKey || url;
          if (seenItems.has(dedupeKey)) {
            continue;
          }

          seenItems.add(dedupeKey);
          discovered.push({
            url,
            title: title || url,
            subtitle: buildCardSubtitle(reactItem || {}, textLines, title),
            identityKey,
            candidateUrls: sortExtractionCandidates(urlCandidates),
            graphDriveId: buildCardGraphDriveId(reactItem || {}),
            graphItemId: buildCardGraphItemId(reactItem || {}),
          });
        }

        return {
          items: discovered,
          visibleCardTitles,
          visibleCardCount: cards.length,
        };
      }

      function collectAnchorItems() {
        const discovered = [];
        const seenUrls = new Set();
        const anchors = [...document.querySelectorAll('a[href]')].filter(isVisible);

        for (const anchor of anchors) {
          if (!isLikelyMeetingHref(anchor.href)) {
            continue;
          }

          const url = normalizeHref(anchor.href);
          if (!url || seenUrls.has(url)) {
            continue;
          }

          const card =
            anchor.closest('article, li, [role="listitem"], [data-item-index]') ||
            anchor.parentElement ||
            anchor;
          const textLines = String(card?.innerText || anchor.innerText || '')
            .split(/\\r?\\n/)
            .map((line) => normalizeText(line))
            .filter(Boolean);
          const title =
            normalizeText(anchor.getAttribute('aria-label')) ||
            normalizeText(
              card?.querySelector('[title]')?.getAttribute('title') || '',
            ) ||
            normalizeText(anchor.querySelector('img')?.getAttribute('alt') || '') ||
            textLines[0] ||
            normalizeText(anchor.textContent);
          const subtitle = textLines.slice(1, 3).join(' · ');

          discovered.push({
            url,
            title,
            subtitle,
          });
          seenUrls.add(url);
        }

        return discovered;
      }

      function getScrollMetrics(container) {
        if (container === document.scrollingElement) {
          return {
            top: window.scrollY,
            height: window.innerHeight,
            maxTop: Math.max(
              document.scrollingElement.scrollHeight - window.innerHeight,
              0,
            ),
          };
        }

        return {
          top: container.scrollTop,
          height: container.clientHeight,
          maxTop: Math.max(container.scrollHeight - container.clientHeight, 0),
        };
      }

      function setScrollTop(container, top) {
        if (container === document.scrollingElement) {
          window.scrollTo(0, top);
          return;
        }

        container.scrollTop = top;
      }

      function countAnchors(container) {
        return container.querySelectorAll('a[href]').length;
      }

      function countVisibleCards(container) {
        return getRenderedCards(container).length;
      }

      function findScrollContainer() {
        const renderedCardsRoot = findRenderedCardsRoot();
        let current = renderedCardsRoot;
        while (current) {
          if (
            current instanceof HTMLElement &&
            current.scrollHeight > current.clientHeight + 120
          ) {
            return current;
          }

          current = current.parentElement;
        }

        const candidates = [
          document.scrollingElement,
          ...document.querySelectorAll('main, [role="main"], section, div'),
        ].filter(
          (element) =>
            element instanceof HTMLElement &&
            element.scrollHeight > element.clientHeight + 120,
        );

        candidates.sort(
          (left, right) =>
            countVisibleCards(right) - countVisibleCards(left) ||
            countAnchors(right) - countAnchors(left),
        );
        return candidates[0] || document.scrollingElement || document.body;
      }

      const meetingsControl = findMeetingsControl();
      if (!meetingsControl) {
        return {
          error: 'Could not find the Meetings filter in Stream.',
          items: [],
          controlLabel: '',
          selectedBeforeClick: false,
          scrollIterations: 0,
        };
      }

      const controlElement = meetingsControl.element;
      const selectedBeforeClick = looksActive(controlElement);
      if (!selectedBeforeClick) {
        clickElement(controlElement);
        await wait(settleMs * 2);
      }

      const container = findScrollContainer();
      const initialRenderedDiscovery = collectRenderedCardItems(
        findRenderedCardsRoot() || container,
      );
      const visibleCardTitles = [...initialRenderedDiscovery.visibleCardTitles];
      let stablePasses = 0;
      let scrollIterations = 0;

      while (scrollIterations < 80 && stablePasses < 4) {
        const renderedRoot = findRenderedCardsRoot() || container;
        const beforeRendered = collectRenderedCardItems(renderedRoot);
        const beforeItems = collectAnchorItems();
        const beforeSurfaceCount = Math.max(
          beforeRendered.visibleCardCount,
          beforeRendered.items.length,
          beforeItems.length,
          countAnchors(container),
          countVisibleCards(container),
        );
        const metrics = getScrollMetrics(container);
        const nextTop = Math.min(
          metrics.top + Math.max(Math.floor(metrics.height * 0.85), 320),
          metrics.maxTop,
        );
        const moved = nextTop > metrics.top + 4;

        setScrollTop(container, nextTop);
        await wait(settleMs);

        const afterRendered = collectRenderedCardItems(renderedRoot);
        const afterItems = collectAnchorItems();
        visibleCardTitles.length = 0;
        visibleCardTitles.push(...afterRendered.visibleCardTitles);
        const afterSurfaceCount = Math.max(
          afterRendered.visibleCardCount,
          afterRendered.items.length,
          afterItems.length,
          countAnchors(container),
          countVisibleCards(container),
        );
        const noNewItems = afterSurfaceCount <= beforeSurfaceCount;
        const atBottom = nextTop >= metrics.maxTop - 4;

        if ((!moved && noNewItems) || (atBottom && noNewItems)) {
          stablePasses += 1;
        } else {
          stablePasses = 0;
        }

        scrollIterations += 1;
      }

      const renderedRoot = findRenderedCardsRoot() || container;
      const renderedDiscovery = collectRenderedCardItems(renderedRoot);
      const anchorItems = collectAnchorItems();

      return {
        error: '',
        controlLabel: buildControlLabel(controlElement),
        selectedBeforeClick,
        scrollIterations,
        items: [...renderedDiscovery.items, ...anchorItems],
        visibleCardCount: renderedDiscovery.visibleCardCount,
        visibleCardTitles: renderedDiscovery.visibleCardTitles,
        page: {
          title: document.title.replace(' - Microsoft Stream', '').trim(),
          url: location.href,
        },
      };
    })()
  `;
}
