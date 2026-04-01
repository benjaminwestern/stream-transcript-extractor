export function createMeetingNavigationRuntime({
  defaultCrawlStartUrl,
  defaultPageNavigationTimeoutMs,
  defaultCrawlScrollSettleMs,
  navigatePage,
  evaluate,
  normalizeMeetingUrl,
  errorFactory,
}) {
  function isClearlyWrongMeetingLandingPage(pageSnapshot, expectedUrl = '') {
    const title = String(pageSnapshot?.title || '').trim().toLowerCase();
    const url = normalizeMeetingUrl(pageSnapshot?.url || '').toLowerCase();
    const bodyTextPreview = String(pageSnapshot?.bodyTextPreview || '')
      .trim()
      .toLowerCase();
    const expectedNormalizedUrl = normalizeMeetingUrl(expectedUrl).toLowerCase();

    if (!url) {
      return true;
    }

    if (/login\.microsoftonline\.com/.test(url) || title === 'sign in to your account') {
      return true;
    }

    if (
      url === defaultCrawlStartUrl.toLowerCase() ||
      /\/launch\/stream\/\?auth=2&home=1(?:&|$)/.test(url)
    ) {
      if (/m365 copilot|stream \| m365 copilot/.test(title)) {
        return true;
      }
    }

    if (
      expectedNormalizedUrl &&
      /meetingcatchupportal|meetingcatchup/.test(expectedNormalizedUrl) &&
      url !== expectedNormalizedUrl &&
      /m365\.cloud\.microsoft/.test(url)
    ) {
      return true;
    }

    if (
      /organizational policy requires you to sign in again|forgot my password/.test(
        bodyTextPreview,
      )
    ) {
      return true;
    }

    return false;
  }

  async function getCurrentPageSnapshot(cdp) {
    return evaluate(
      cdp,
      `
        (() => ({
          title: document.title.replace(' - Microsoft Stream', '').trim(),
          url: location.href,
          bodyTextPreview: String(document.body?.innerText || '')
            .replace(/\\s+/g, ' ')
            .trim()
            .slice(0, 500),
        }))()
      `,
    );
  }

  async function navigatePageAndWait(
    cdp,
    url,
    timeoutMs = defaultPageNavigationTimeoutMs,
  ) {
    await navigatePage(cdp, url, {
      timeoutMs,
      settleMs: defaultCrawlScrollSettleMs,
      errorFactory,
    });
    return getCurrentPageSnapshot(cdp).catch(() => ({
      title: '',
      url,
    }));
  }

  return {
    isClearlyWrongMeetingLandingPage,
    getCurrentPageSnapshot,
    navigatePageAndWait,
  };
}
