function renderPageChoice(page, maxUrlLength = 80) {
  return `${page.title} (${String(page.url || '').slice(0, maxUrlLength)})`;
}

export function isLikelyMeetingPage(page) {
  const haystack = [page?.title || '', page?.url || ''].join(' ').toLowerCase();

  return /stream|sharepoint|recordings|meeting|transcript|stream\.aspx/.test(
    haystack,
  );
}

export async function selectTranscriptPage({
  prompt,
  debugPort,
  findPageTargets,
  chooseFromList,
  emptyPageMessage = 'No browser pages were found. Open Microsoft Stream before continuing.',
  selectionTitle = 'Open pages',
  selectionQuestion = '\nWhich page contains the transcript? (number): ',
}) {
  const pages = await findPageTargets(debugPort);

  if (pages.length === 0) {
    throw new Error(emptyPageMessage);
  }

  if (pages.length === 1) {
    return pages[0];
  }

  const likelyMeetingPages = pages.filter((page) => isLikelyMeetingPage(page));
  if (likelyMeetingPages.length === 1) {
    return likelyMeetingPages[0];
  }

  return chooseFromList(
    prompt,
    selectionTitle,
    pages,
    renderPageChoice,
    selectionQuestion,
  );
}
