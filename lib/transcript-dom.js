export function buildTranscriptExtractionExpression({
  extractionSettings,
  debugEnabled = false,
}) {
  const settings = JSON.stringify(extractionSettings);
  const debugMode = JSON.stringify(Boolean(debugEnabled));

  return `
    (async () => {
      const settings = ${settings};
      const debugEnabled = ${debugMode};
      const selectorGroups = [
        '.ms-List-cell',
        '[role="listitem"]',
        '[data-list-index]',
        '[id^="sub-entry-"]',
      ];
      const debug = debugEnabled
        ? {
            settings,
            scrollContainer: null,
            selectorMatchCounts: {},
            strategyCounts: {
              cells: 0,
              text: 0,
            },
            passes: [],
            headerSamples: [],
            textSamples: {
              first: null,
              last: null,
            },
            final: null,
          }
        : null;
      const debugHeaderKeys = debugEnabled ? new Set() : null;

      function normalizeText(value) {
        return String(value || '')
          .replace(/\\u200b/g, '')
          .replace(/\\s+/g, ' ')
          .trim();
      }

      function trimForDebug(value, maxLength = 160) {
        const text = normalizeText(value);
        if (text.length <= maxLength) {
          return text;
        }

        return \`\${text.slice(0, maxLength - 3)}...\`;
      }

      function pushDebugSample(array, value, maxItems = 40) {
        if (!debugEnabled || array.length >= maxItems) {
          return;
        }

        array.push(value);
      }

      function pushDebugHeader(kind, speaker, timestamp, line) {
        if (!debugEnabled) {
          return;
        }

        const key = [kind, speaker, timestamp].join('|');
        if (debugHeaderKeys.has(key)) {
          return;
        }

        debugHeaderKeys.add(key);
        pushDebugSample(
          debug.headerSamples,
          {
            kind,
            speaker: trimForDebug(speaker, 120),
            timestamp,
            line: trimForDebug(line, 220),
          },
          60,
        );
      }

      function isTimestamp(value) {
        return /^\\d{1,2}:\\d{2}(:\\d{2})?$/.test(normalizeText(value));
      }

      function isLifecycleMessage(value) {
        return /(?:started|stopped) transcription/i.test(value);
      }

      function looksLikeSpeakerName(value, knownSpeakers = null) {
        const text = normalizeText(value);
        if (!text || isTimestamp(text) || text.length > 80) {
          return false;
        }

        if (knownSpeakers && knownSpeakers.has(text)) {
          return true;
        }

        if (/[?!,:;]/.test(text)) {
          return false;
        }

        if (/[.]$/.test(text) && !/^(?:[A-Z]\\.?|[A-Z][a-z]+\\.)$/.test(text)) {
          return false;
        }

        const connectorWords = new Set([
          'al',
          'bin',
          'da',
          'de',
          'del',
          'della',
          'di',
          'dos',
          'du',
          'la',
          'le',
          'van',
          'von',
        ]);
        const words = text.split(' ').filter(Boolean);

        if (words.length === 0 || words.length > 6) {
          return false;
        }

        let matchedWordCount = 0;

        for (const rawWord of words) {
          const word = rawWord.replace(/[()]/g, '');
          const lowerWord = word.toLowerCase();

          if (!word) {
            continue;
          }

          if (connectorWords.has(lowerWord)) {
            continue;
          }

          if (
            /^[A-Z][a-z]+(?:['-][A-Z]?[a-z]+)*$/.test(word) ||
            /^[A-Z]{2,}$/.test(word) ||
            /^[A-Z]\\.?$/.test(word)
          ) {
            matchedWordCount += 1;
            continue;
          }

          return false;
        }

        if (matchedWordCount === 0) {
          return false;
        }

        if (words.length === 1) {
          return /^[A-Z][A-Za-z0-9'().-]{1,29}$/.test(text);
        }

        return matchedWordCount >= Math.max(1, words.length - 1);
      }

      function looksLikeSpeakerNameAfterTimestamp(value, knownSpeakers) {
        const text = normalizeText(value);
        if (!looksLikeSpeakerName(text, knownSpeakers)) {
          return false;
        }

        return (
          (knownSpeakers && knownSpeakers.has(text)) ||
          text.split(' ').filter(Boolean).length >= 2
        );
      }

      function parseSpeakerTimestampLine(value, knownSpeakers = null) {
        const text = normalizeText(value);
        if (!text) {
          return null;
        }

        const speakerThenTimestampMatch = text.match(
          /^(.*?)\\s+(\\d{1,2}:\\d{2}(?::\\d{2})?)$/,
        );
        if (speakerThenTimestampMatch) {
          const speaker = normalizeText(speakerThenTimestampMatch[1]);
          const timestamp = normalizeText(speakerThenTimestampMatch[2]);

          if (looksLikeSpeakerName(speaker, knownSpeakers)) {
            pushDebugHeader('speaker-then-timestamp', speaker, timestamp, text);
            return { speaker, timestamp };
          }
        }

        const timestampThenSpeakerMatch = text.match(
          /^(\\d{1,2}:\\d{2}(?::\\d{2})?)\\s+(.*)$/,
        );
        if (timestampThenSpeakerMatch) {
          const timestamp = normalizeText(timestampThenSpeakerMatch[1]);
          const speaker = normalizeText(timestampThenSpeakerMatch[2]);

          if (looksLikeSpeakerNameAfterTimestamp(speaker, knownSpeakers)) {
            pushDebugHeader('timestamp-then-speaker', speaker, timestamp, text);
            return { speaker, timestamp };
          }
        }

        return null;
      }

      function isSpeakerTimestampBoundary(line, nextLine, knownSpeakers) {
        return (
          !isTimestamp(line) &&
          looksLikeSpeakerName(line, knownSpeakers) &&
          isTimestamp(nextLine)
        );
      }

      function buildEntryKey(entry) {
        return [
          entry.speaker,
          entry.timestamp,
          entry.text.slice(0, settings.dedupePrefixLength),
        ].join('|');
      }

      let scrollEl = null;
      for (const el of document.querySelectorAll('*')) {
        const overflowY = window.getComputedStyle(el).overflowY;
        const sampleText = normalizeText(el.textContent).slice(0, 500);
        const isScrollable =
          (overflowY === 'auto' || overflowY === 'scroll') &&
          el.scrollHeight > el.clientHeight + 20 &&
          el.clientHeight > 100;

        if (isScrollable && /\\d{1,2}:\\d{2}/.test(sampleText)) {
          scrollEl = el;
          break;
        }
      }

      if (!scrollEl) {
        return {
          entries: [],
          scrollCount: 0,
          lastTimestamp: '',
          error: 'No scrollable transcript container found.',
          debug,
        };
      }

      if (debugEnabled) {
        debug.scrollContainer = {
          tagName: scrollEl.tagName,
          className: trimForDebug(scrollEl.className?.toString?.() || '', 240),
          role: scrollEl.getAttribute('role') || '',
          ariaLabel: scrollEl.getAttribute('aria-label') || '',
          childElementCount: scrollEl.childElementCount,
          clientHeight: scrollEl.clientHeight,
          scrollHeight: scrollEl.scrollHeight,
        };
        debug.selectorMatchCounts = Object.fromEntries(
          selectorGroups.map((selector) => [
            selector,
            scrollEl.querySelectorAll(selector).length,
          ]),
        );
      }

      function collectUniqueElements(elements) {
        return Array.from(new Set(elements.filter(Boolean)));
      }

      function collectCandidateCells() {
        for (const selector of selectorGroups) {
          const cells = collectUniqueElements([
            ...scrollEl.querySelectorAll(selector),
          ]).filter((cell) => cell instanceof HTMLElement && cell !== scrollEl);

          if (cells.length > 0) {
            return {
              selectedSelector: selector,
              cells,
            };
          }
        }

        return {
          selectedSelector: '',
          cells: [],
        };
      }

      function extractTextFromContainer(container, ignoredValues = new Set()) {
        const walker = document.createTreeWalker(
          container,
          NodeFilter.SHOW_TEXT,
          null,
        );
        const parts = [];
        let node = walker.nextNode();

        while (node) {
          const parent = node.parentElement;
          const parentClassName = parent?.className?.toString?.() || '';
          const text = normalizeText(node.textContent);

          const isIgnoredText =
            !text ||
            ignoredValues.has(text) ||
            isTimestamp(text) ||
            parentClassName.includes('screenReaderFriendly') ||
            parentClassName.includes('eventSpeakerName') ||
            parentClassName.includes('itemDisplayName') ||
            parentClassName.includes('itemHeader') ||
            parent?.getAttribute('aria-hidden') === 'true';

          if (!isIgnoredText) {
            parts.push(text);
          }

          node = walker.nextNode();
        }

        return normalizeText(parts.join(' '));
      }

      function extractVisibleEntriesFromCells() {
        const { selectedSelector, cells } = collectCandidateCells();
        if (debugEnabled) {
          debug.strategyCounts.cells += 1;
        }

        const results = [];
        let currentSpeaker = '';
        let currentTimestamp = '';

        for (const cell of cells) {
          const speakerEl =
            cell.querySelector('[class*="itemDisplayName"]') ||
            cell.querySelector('[class*="eventSpeakerName"]') ||
            cell.querySelector('[data-tid*="speaker"]');
          const speakerText = normalizeText(speakerEl?.textContent);
          if (speakerText && !isTimestamp(speakerText)) {
            currentSpeaker = speakerText;
          }

          const timestampEls = collectUniqueElements([
            ...cell.querySelectorAll('[aria-hidden="true"]'),
            ...cell.querySelectorAll('time'),
          ]);
          for (const timestampEl of timestampEls) {
            const maybeTimestamp = normalizeText(timestampEl.textContent);
            if (isTimestamp(maybeTimestamp)) {
              currentTimestamp = maybeTimestamp;
              break;
            }
          }

          const textContainer =
            cell.querySelector('[class*="entryText"]') ||
            cell.querySelector('[id^="sub-entry-"]') ||
            cell;
          const ignoredValues = new Set(
            [speakerText, currentSpeaker, currentTimestamp].filter(Boolean),
          );
          const text = extractTextFromContainer(textContainer, ignoredValues);

          if (!text || isLifecycleMessage(text) || !currentSpeaker) {
            continue;
          }

          results.push({
            speaker: currentSpeaker,
            timestamp: currentTimestamp,
            text,
          });
        }

        return {
          strategy: 'cells',
          selectedSelector,
          candidateCellCount: cells.length,
          lineCount: 0,
          entries: results,
        };
      }

      function extractVisibleEntriesFromText() {
        const lines = String(scrollEl.innerText || scrollEl.textContent || '')
          .split(/\\r?\\n/)
          .map(normalizeText)
          .filter(Boolean);
        if (debugEnabled) {
          debug.strategyCounts.text += 1;
          const sample = {
            lineCount: lines.length,
            firstLines: lines.slice(0, 12).map((line) => trimForDebug(line)),
            lastLines: lines.slice(-12).map((line) => trimForDebug(line)),
          };

          if (!debug.textSamples.first) {
            debug.textSamples.first = sample;
          }
          debug.textSamples.last = sample;
        }

        const results = [];
        const knownSpeakers = new Set();
        let currentSpeaker = '';
        let index = 0;

        while (index < lines.length) {
          const line = lines[index];
          const nextLine = lines[index + 1] || '';
          const combinedHeader = parseSpeakerTimestampLine(
            line,
            knownSpeakers,
          );

          if (combinedHeader) {
            currentSpeaker = combinedHeader.speaker;
            knownSpeakers.add(currentSpeaker);

            const textLines = [];
            let cursor = index + 1;

            while (cursor < lines.length) {
              const candidate = lines[cursor];
              const followingLine = lines[cursor + 1] || '';

              if (parseSpeakerTimestampLine(candidate, knownSpeakers)) {
                break;
              }

              if (isTimestamp(candidate)) {
                break;
              }

              if (
                isSpeakerTimestampBoundary(
                  candidate,
                  followingLine,
                  knownSpeakers,
                )
              ) {
                break;
              }

              textLines.push(candidate);
              cursor += 1;
            }

            const text = normalizeText(textLines.join(' '));
            if (text && !isLifecycleMessage(text)) {
              results.push({
                speaker: currentSpeaker,
                timestamp: combinedHeader.timestamp,
                text,
              });
            }

            index = Math.max(cursor, index + 1);
            continue;
          }

          if (isSpeakerTimestampBoundary(line, nextLine, knownSpeakers)) {
            pushDebugHeader(
              'speaker-line-then-timestamp-line',
              line,
              nextLine,
              \`\${line} \${nextLine}\`,
            );
            currentSpeaker = line;
            knownSpeakers.add(currentSpeaker);
            index += 1;
            continue;
          }

          if (!isTimestamp(line)) {
            index += 1;
            continue;
          }

          const timestamp = line;
          const previousLine = lines[index - 1] || '';
          if (looksLikeSpeakerName(previousLine, knownSpeakers)) {
            currentSpeaker = previousLine;
            knownSpeakers.add(currentSpeaker);
          }

          const textLines = [];
          let cursor = index + 1;
          const immediateNextLine = lines[cursor] || '';

          if (
            looksLikeSpeakerNameAfterTimestamp(
              immediateNextLine,
              knownSpeakers,
            )
          ) {
            currentSpeaker = immediateNextLine;
            knownSpeakers.add(currentSpeaker);
            pushDebugHeader(
              'timestamp-line-then-speaker-line',
              currentSpeaker,
              timestamp,
              \`\${timestamp} \${currentSpeaker}\`,
            );
            cursor += 1;
          }

          while (cursor < lines.length) {
            const candidate = lines[cursor];
            const followingLine = lines[cursor + 1] || '';

            if (parseSpeakerTimestampLine(candidate, knownSpeakers)) {
              break;
            }

            if (isTimestamp(candidate)) {
              break;
            }

            if (
              isSpeakerTimestampBoundary(
                candidate,
                followingLine,
                knownSpeakers,
              )
            ) {
              break;
            }

            textLines.push(candidate);
            cursor += 1;
          }

          const text = normalizeText(textLines.join(' '));
          if (text && currentSpeaker && !isLifecycleMessage(text)) {
            results.push({
              speaker: currentSpeaker,
              timestamp,
              text,
            });
          }

          index = Math.max(cursor, index + 1);
        }

        return {
          strategy: 'text',
          selectedSelector: '',
          candidateCellCount: 0,
          lineCount: lines.length,
          entries: results,
        };
      }

      function extractVisibleEntries() {
        const cellResult = extractVisibleEntriesFromCells();
        if (cellResult.entries.length > 0) {
          return cellResult;
        }

        const textResult = extractVisibleEntriesFromText();
        return {
          ...textResult,
          fallbackFromSelector: cellResult.selectedSelector,
          fallbackCandidateCellCount: cellResult.candidateCellCount,
        };
      }

      const allEntries = new Map();
      const viewportChunk = Math.floor(
        scrollEl.clientHeight * settings.viewportChunkRatio,
      );

      let scrollCount = 0;
      let previousLastTimestamp = '';
      let stablePasses = 0;

      while (
        stablePasses < settings.stableScrollPasses &&
        scrollCount < settings.maxScrollIterations
      ) {
        const visibleResult = extractVisibleEntries();
        let newEntryCount = 0;

        for (const entry of visibleResult.entries) {
          const key = buildEntryKey(entry);
          if (!allEntries.has(key)) {
            allEntries.set(key, entry);
            newEntryCount += 1;
          }
        }

        scrollEl.scrollBy(0, viewportChunk);
        scrollCount += 1;
        await new Promise((resolvePromise) =>
          setTimeout(resolvePromise, settings.scrollSettleMs),
        );

        const timestamps = collectUniqueElements([
          ...scrollEl.querySelectorAll('[aria-hidden="true"]'),
          ...scrollEl.querySelectorAll('time'),
        ]);
        let newestVisibleTimestamp = '';

        for (const timestampEl of timestamps) {
          const maybeTimestamp = normalizeText(timestampEl.textContent);
          if (isTimestamp(maybeTimestamp)) {
            newestVisibleTimestamp = maybeTimestamp;
          }
        }

        if (newestVisibleTimestamp === previousLastTimestamp) {
          stablePasses += 1;
        } else {
          stablePasses = 0;
          previousLastTimestamp = newestVisibleTimestamp;
        }

        if (debugEnabled) {
          pushDebugSample(
            debug.passes,
            {
              pass: scrollCount,
              strategy: visibleResult.strategy,
              visibleEntryCount: visibleResult.entries.length,
              newEntryCount,
              selectedSelector: visibleResult.selectedSelector || '',
              candidateCellCount: visibleResult.candidateCellCount || 0,
              fallbackFromSelector: visibleResult.fallbackFromSelector || '',
              fallbackCandidateCellCount:
                visibleResult.fallbackCandidateCellCount || 0,
              lineCount: visibleResult.lineCount || 0,
              newestVisibleTimestamp,
              stablePasses,
            },
            250,
          );
        }
      }

      const finalVisibleResult = extractVisibleEntries();
      for (const entry of finalVisibleResult.entries) {
        const key = buildEntryKey(entry);
        if (!allEntries.has(key)) {
          allEntries.set(key, entry);
        }
      }

      const finalEntries = Array.from(allEntries.values());
      if (debugEnabled) {
        debug.final = {
          entryCount: finalEntries.length,
          uniqueSpeakerCount: new Set(
            finalEntries.map((entry) => entry.speaker).filter(Boolean),
          ).size,
          lastTimestamp: previousLastTimestamp,
          finalVisibleStrategy: finalVisibleResult.strategy,
          finalVisibleEntryCount: finalVisibleResult.entries.length,
        };
      }

      return {
        entries: finalEntries,
        scrollCount,
        lastTimestamp: previousLastTimestamp,
        debug,
      };
    })()
  `;
}
