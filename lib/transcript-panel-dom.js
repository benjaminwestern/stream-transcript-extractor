export function buildTranscriptPanelAutomationExpression(action = 'inspect') {
  const serializedAction = JSON.stringify(action);

  return `
    (() => {
      const action = ${serializedAction};

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
          rect.width > 1 &&
          rect.height > 1
        );
      }

      function buildLabel(element) {
        return normalizeText(
          [
            element.getAttribute('aria-label'),
            element.getAttribute('title'),
            element.getAttribute('data-tid'),
            element.getAttribute('name'),
            element.textContent,
          ]
            .filter(Boolean)
            .join(' '),
        );
      }

      function isTranscriptLabel(text) {
        return /\\b(transcript|captions?|subtitles?)\\b/i.test(text);
      }

      function looksLikeTranscriptContainer(element) {
        if (!(element instanceof HTMLElement) || !isVisible(element)) {
          return false;
        }

        const overflowY = window.getComputedStyle(element).overflowY;
        const sampleText = normalizeText(
          element.innerText || element.textContent || '',
        ).slice(0, 800);

        return (
          (overflowY === 'auto' || overflowY === 'scroll') &&
          element.scrollHeight > element.clientHeight + 20 &&
          element.clientHeight > 80 &&
          (/\\d{1,2}:\\d{2}/.test(sampleText) ||
            /\\b(?:started|stopped) transcription\\b/i.test(sampleText))
        );
      }

      function findTranscriptContainer() {
        for (const element of document.querySelectorAll('*')) {
          if (looksLikeTranscriptContainer(element)) {
            return element;
          }
        }

        return null;
      }

      function scoreControl(element) {
        const label = buildLabel(element);
        if (!label || !isTranscriptLabel(label)) {
          return -1;
        }

        const lowerLabel = label.toLowerCase();
        let score = 0;

        if (element.tagName === 'BUTTON') {
          score += 4;
        }

        const role = String(element.getAttribute('role') || '').toLowerCase();
        if (role === 'button') {
          score += 3;
        } else if (role === 'tab' || role === 'menuitem') {
          score += 2;
        }

        if (/^(transcript|captions?|subtitles?)$/.test(lowerLabel)) {
          score += 10;
        }

        if (/\\b(open|show|view)\\s+(the\\s+)?(transcript|captions?|subtitles?)\\b/.test(lowerLabel)) {
          score += 9;
        }

        if (/\\b(hide|close)\\s+(the\\s+)?(transcript|captions?|subtitles?)\\b/.test(lowerLabel)) {
          score += 8;
        }

        if (
          String(element.getAttribute('aria-controls') || '')
            .toLowerCase()
            .includes('transcript')
        ) {
          score += 4;
        }

        if (String(element.getAttribute('data-tid') || '').toLowerCase().includes('transcript')) {
          score += 4;
        }

        if (String(element.getAttribute('aria-expanded') || '') === 'false') {
          score += action === 'open' ? 2 : 0;
        }

        if (String(element.getAttribute('aria-expanded') || '') === 'true') {
          score += action === 'close' ? 2 : 0;
        }

        if (label.length <= 80) {
          score += 1;
        }

        return score;
      }

      function findBestControl() {
        const selectors = [
          'button',
          '[role="button"]',
          '[role="tab"]',
          '[role="menuitem"]',
          '[aria-label]',
          '[data-tid]',
        ];
        const elements = new Set();

        for (const selector of selectors) {
          document.querySelectorAll(selector).forEach((element) =>
            elements.add(element),
          );
        }

        let bestControl = null;
        let bestScore = -1;

        for (const element of elements) {
          if (!(element instanceof HTMLElement) || !isVisible(element)) {
            continue;
          }

          const score = scoreControl(element);
          if (score > bestScore) {
            bestScore = score;
            bestControl = element;
          }
        }

        return bestControl;
      }

      function clickElement(element) {
        if (!(element instanceof HTMLElement) || element.hasAttribute('disabled')) {
          return false;
        }

        try {
          element.scrollIntoView({
            block: 'center',
            inline: 'center',
          });
        } catch {
          // Ignore scroll failures and keep trying to click.
        }

        try {
          for (const eventName of [
            'pointerdown',
            'mousedown',
            'pointerup',
            'mouseup',
            'click',
          ]) {
            element.dispatchEvent(
              new MouseEvent(eventName, {
                bubbles: true,
                cancelable: true,
                view: window,
              }),
            );
          }
          element.click();
          return true;
        } catch {
          try {
            element.click();
            return true;
          } catch {
            return false;
          }
        }
      }

      const container = findTranscriptContainer();
      const control = findBestControl();
      const controlLabel = control ? buildLabel(control) : '';
      const controlExpanded = control
        ? String(control.getAttribute('aria-expanded') || '')
        : '';
      const result = {
        action,
        panelOpen: Boolean(container),
        panelLikelyOpen: Boolean(container) || controlExpanded === 'true',
        controlFound: Boolean(control),
        controlLabel,
        controlExpanded,
        scrollTop: container ? container.scrollTop : 0,
        clientHeight: container ? container.clientHeight : 0,
        scrollHeight: container ? container.scrollHeight : 0,
      };

      if (action === 'inspect') {
        return result;
      }

      if (action === 'open') {
        if (result.panelLikelyOpen) {
          return {
            ...result,
            ok: true,
            performed: 'already-open',
          };
        }

        if (!control) {
          return {
            ...result,
            ok: false,
            performed: 'missing-control',
          };
        }

        return {
          ...result,
          ok: clickElement(control),
          performed: 'clicked-open',
        };
      }

      if (action === 'close') {
        if (!result.panelLikelyOpen) {
          return {
            ...result,
            ok: true,
            performed: 'already-closed',
          };
        }

        if (!control) {
          return {
            ...result,
            ok: false,
            performed: 'missing-control',
          };
        }

        return {
          ...result,
          ok: clickElement(control),
          performed: 'clicked-close',
        };
      }

      if (action === 'nudge-scroll') {
        if (!container) {
          return {
            ...result,
            ok: false,
            performed: 'missing-container',
          };
        }

        const beforeScrollTop = container.scrollTop;
        const maxScrollTop = Math.max(
          0,
          container.scrollHeight - container.clientHeight,
        );
        const targetScrollTop =
          maxScrollTop > 0
            ? Math.min(
                maxScrollTop,
                Math.max(
                  beforeScrollTop + Math.floor(container.clientHeight * 1.5),
                  Math.floor(maxScrollTop * 0.6),
                ),
              )
            : 0;
        container.scrollTop =
          targetScrollTop === beforeScrollTop && maxScrollTop > 0
            ? maxScrollTop
            : targetScrollTop;

        return {
          ...result,
          ok: true,
          performed: 'nudged-scroll',
          beforeScrollTop,
          afterScrollTop: container.scrollTop,
        };
      }

      return {
        ...result,
        ok: false,
        performed: 'unsupported-action',
      };
    })()
  `;
}
