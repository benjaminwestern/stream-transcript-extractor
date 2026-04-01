import { buildTranscriptPanelAutomationExpression } from './transcript-panel-dom.js';

export function logAutomaticAction(
  captureFeedback,
  debugEnabled,
  step,
  details = null,
) {
  const entry = {
    at: new Date().toISOString(),
    step,
    details,
  };

  captureFeedback.automaticActions.push(entry);
  if (!debugEnabled) {
    return;
  }

  const detailSuffix = details ? ` ${JSON.stringify(details)}` : '';
  console.log(`[automatic debug] ${step}${detailSuffix}`);
}

export function createAutomaticTranscriptCaptureRuntime({
  evaluate,
  reloadPageAndWait,
  sleep,
  defaultAutomaticUiPollMs,
  defaultAutomaticPanelOpenTimeoutMs,
  defaultAutomaticClickSettleMs,
  defaultAutomaticSignalTimeoutMs,
  defaultAutomaticSignalRetryTimeoutMs,
  defaultAutomaticRequestSettleTimeoutMs,
  defaultCaptureSettleMs,
}) {
  function printAutomaticFallbackInstructions() {
    console.log('\nAutomatic mode needs a manual assist.');
    console.log('1. Open or reopen the Transcript panel in Microsoft Stream.');
    console.log(
      '2. Let the transcript load, and scroll once if the app lazily fetches chunks.',
    );
    console.log(
      '3. Watch this terminal for a transcript-traffic confirmation if one is seen.',
    );
    console.log('4. Return to this terminal and press Enter.');
  }

  async function waitForAsyncCondition(
    check,
    timeoutMs,
    pollMs = defaultAutomaticUiPollMs,
  ) {
    const deadline = Date.now() + timeoutMs;

    while (Date.now() < deadline) {
      const result = await check();
      if (result) {
        return result;
      }

      await sleep(pollMs);
    }

    return null;
  }

  function summarizeTranscriptPanelUiState(state) {
    if (!state) {
      return null;
    }

    return {
      panelOpen: Boolean(state.panelOpen),
      panelLikelyOpen: Boolean(state.panelLikelyOpen),
      controlFound: Boolean(state.controlFound),
      controlLabel: state.controlLabel || '',
      controlExpanded: state.controlExpanded || '',
      scrollTop:
        typeof state.scrollTop === 'number' ? state.scrollTop : null,
      clientHeight:
        typeof state.clientHeight === 'number' ? state.clientHeight : null,
      scrollHeight:
        typeof state.scrollHeight === 'number' ? state.scrollHeight : null,
      performed: state.performed || '',
      ok: typeof state.ok === 'boolean' ? state.ok : null,
    };
  }

  function countAutomaticActions(captureFeedback, prefix) {
    return captureFeedback.automaticActions.filter((entry) =>
      entry.step.startsWith(prefix),
    ).length;
  }

  function buildAutomaticFallbackSummary(
    reason,
    captureFeedback,
    panelResult = null,
  ) {
    const lines = [];
    const state = panelResult?.state || null;
    const controlLabel = panelResult?.controlLabel || state?.controlLabel || '';
    const openAttempts = countAutomaticActions(
      captureFeedback,
      'open transcript panel attempt',
    );
    const nudgeAttempts = countAutomaticActions(
      captureFeedback,
      'nudge transcript panel',
    );

    if (reason === 'panel-open-failed') {
      if (state?.controlFound) {
        lines.push(
          controlLabel
            ? `I found a likely Transcript control ("${controlLabel}") but the panel did not appear to open.`
            : 'I found a likely Transcript control, but the panel did not appear to open.',
        );
      } else {
        lines.push(
          'I could not find a visible Transcript control to click on this page.',
        );
      }
    } else if (reason === 'traffic-not-detected') {
      if (state?.panelLikelyOpen) {
        lines.push(
          'I was able to get what looked like an open Transcript panel.',
        );
      } else if (state?.controlFound) {
        lines.push(
          controlLabel
            ? `I found and tried the likely Transcript control ("${controlLabel}"), but I could not confirm that the panel stayed open.`
            : 'I found and tried a likely Transcript control, but I could not confirm that the panel stayed open.',
        );
      } else {
        lines.push(
          'I did not reliably confirm an open Transcript panel before capture timed out.',
        );
      }

      if (captureFeedback.transcriptHintCount > 0) {
        lines.push(
          `I saw ${captureFeedback.transcriptHintCount} transcript-like network response` +
            `${captureFeedback.transcriptHintCount === 1 ? '' : 's'}, but none produced a usable transcript payload yet.`,
        );
      } else if (captureFeedback.candidateResponseCount > 0) {
        lines.push(
          `I saw ${captureFeedback.candidateResponseCount} potentially relevant network response` +
            `${captureFeedback.candidateResponseCount === 1 ? '' : 's'}, but none looked transcript-specific.`,
        );
      } else {
        lines.push(
          'I did not see any transcript-like network responses after the automatic panel actions.',
        );
      }
    }

    const attemptedActions = [];
    if (panelResult?.refreshed) {
      attemptedActions.push('refreshed the panel state');
    }
    if (openAttempts > 0) {
      attemptedActions.push(
        `${openAttempts} open attempt${openAttempts === 1 ? '' : 's'}`,
      );
    }
    if (nudgeAttempts > 0) {
      attemptedActions.push(
        `${nudgeAttempts} scroll nudge${nudgeAttempts === 1 ? '' : 's'}`,
      );
    }

    if (attemptedActions.length > 0) {
      lines.push(`Automatic actions attempted: ${attemptedActions.join(', ')}.`);
    }

    return lines;
  }

  function printAutomaticFallbackSummary(
    reason,
    captureFeedback,
    panelResult = null,
  ) {
    const lines = buildAutomaticFallbackSummary(
      reason,
      captureFeedback,
      panelResult,
    );

    if (lines.length === 0) {
      return;
    }

    console.log('Automatic mode summary:');
    for (const line of lines) {
      console.log(`- ${line}`);
    }
  }

  async function handleAutomaticManualAssistFallback(
    prompt,
    captureFeedback,
    debugEnabled,
    reason,
    panelResult,
    allowManualAssist,
  ) {
    printAutomaticFallbackSummary(reason, captureFeedback, panelResult);

    if (!allowManualAssist) {
      logAutomaticAction(
        captureFeedback,
        debugEnabled,
        `${reason}, continuing without manual assist`,
      );
      console.log(
        'Automatic mode: continuing without manual assist so batch extraction can keep moving.',
      );
      return;
    }

    logAutomaticAction(
      captureFeedback,
      debugEnabled,
      `${reason}, falling back to manual assist`,
    );
    printAutomaticFallbackInstructions();
    await prompt.waitForEnter(
      '\nPress Enter after the transcript panel has loaded...\n',
    );
    await sleep(defaultCaptureSettleMs);
  }

  async function inspectTranscriptPanelUi(cdp) {
    return evaluate(cdp, buildTranscriptPanelAutomationExpression('inspect'));
  }

  async function performTranscriptPanelAction(cdp, action) {
    return evaluate(cdp, buildTranscriptPanelAutomationExpression(action));
  }

  async function waitForTranscriptPanelToOpen(
    cdp,
    timeoutMs = defaultAutomaticPanelOpenTimeoutMs,
  ) {
    return waitForAsyncCondition(
      async () => {
        const state = await inspectTranscriptPanelUi(cdp);
        return state.panelLikelyOpen ? state : null;
      },
      timeoutMs,
    );
  }

  async function ensureTranscriptPanelOpenAutomatically(
    cdp,
    captureFeedback,
    debugEnabled,
    { refreshIfOpen = false } = {},
  ) {
    const initialState = await inspectTranscriptPanelUi(cdp);
    logAutomaticAction(
      captureFeedback,
      debugEnabled,
      'inspect transcript panel',
      summarizeTranscriptPanelUiState(initialState),
    );
    let refreshed = false;

    if (initialState.panelLikelyOpen && !refreshIfOpen) {
      return {
        opened: true,
        refreshed,
        controlLabel: initialState.controlLabel || '',
        state: initialState,
      };
    }

    if (initialState.panelLikelyOpen && refreshIfOpen) {
      const closeResult = await performTranscriptPanelAction(cdp, 'close');
      logAutomaticAction(
        captureFeedback,
        debugEnabled,
        'close transcript panel',
        summarizeTranscriptPanelUiState(closeResult),
      );
      if (closeResult.ok) {
        refreshed = true;
        await sleep(defaultAutomaticClickSettleMs);
      }
    }

    for (let attempt = 1; attempt <= 2; attempt += 1) {
      const openResult = await performTranscriptPanelAction(cdp, 'open');
      logAutomaticAction(
        captureFeedback,
        debugEnabled,
        `open transcript panel attempt ${attempt}`,
        summarizeTranscriptPanelUiState(openResult),
      );
      if (!openResult.controlFound && !openResult.panelLikelyOpen) {
        return {
          opened: false,
          refreshed,
          controlLabel: '',
          state: openResult,
        };
      }

      await sleep(defaultAutomaticClickSettleMs);
      const openState = await waitForTranscriptPanelToOpen(cdp);
      logAutomaticAction(
        captureFeedback,
        debugEnabled,
        `wait for transcript panel attempt ${attempt}`,
        summarizeTranscriptPanelUiState(openState),
      );
      if (openState?.panelLikelyOpen) {
        return {
          opened: true,
          refreshed,
          controlLabel: openResult.controlLabel || openState.controlLabel || '',
          state: openState,
        };
      }
    }

    const finalState = await inspectTranscriptPanelUi(cdp);
    logAutomaticAction(
      captureFeedback,
      debugEnabled,
      'inspect transcript panel after retries',
      summarizeTranscriptPanelUiState(finalState),
    );
    return {
      opened: Boolean(finalState?.panelLikelyOpen),
      refreshed,
      controlLabel: finalState?.controlLabel || '',
      state: finalState,
    };
  }

  async function nudgeTranscriptPanelAutomatically(
    cdp,
    captureFeedback,
    debugEnabled,
  ) {
    const nudgeResult = await performTranscriptPanelAction(cdp, 'nudge-scroll');
    logAutomaticAction(
      captureFeedback,
      debugEnabled,
      'nudge transcript panel',
      summarizeTranscriptPanelUiState(nudgeResult),
    );
    return (
      nudgeResult.ok &&
      typeof nudgeResult.afterScrollTop === 'number' &&
      nudgeResult.afterScrollTop > nudgeResult.beforeScrollTop
    );
  }

  async function waitForTranscriptSignal(captureFeedback, timeoutMs) {
    return waitForAsyncCondition(
      async () =>
        captureFeedback.transcriptHintCount > 0
          ? {
              candidateResponseCount: captureFeedback.candidateResponseCount,
              transcriptHintCount: captureFeedback.transcriptHintCount,
            }
          : null,
      timeoutMs,
    );
  }

  async function waitForTranscriptSignalToSettle(captureFeedback, timeoutMs) {
    return waitForAsyncCondition(
      async () =>
        captureFeedback.transcriptHintSettledCount > 0 ||
        captureFeedback.transcriptHintFailedCount > 0
          ? {
              settledCount: captureFeedback.transcriptHintSettledCount,
              failedCount: captureFeedback.transcriptHintFailedCount,
            }
          : null,
      timeoutMs,
    );
  }

  async function runAutomaticTranscriptCaptureFlow(
    cdp,
    prompt,
    captureFeedback,
    debugEnabled,
    { allowManualAssist = true } = {},
  ) {
    console.log('\nAutomatic mode: reloading the page with capture armed...');
    logAutomaticAction(captureFeedback, debugEnabled, 'reload page');
    await reloadPageAndWait(cdp);

    console.log('Automatic mode: locating the Transcript control...');
    let panelResult = await ensureTranscriptPanelOpenAutomatically(
      cdp,
      captureFeedback,
      debugEnabled,
      {
        refreshIfOpen: true,
      },
    );
    logAutomaticAction(
      captureFeedback,
      debugEnabled,
      'transcript panel result',
      {
        opened: panelResult.opened,
        refreshed: panelResult.refreshed,
        controlLabel: panelResult.controlLabel || '',
        state: summarizeTranscriptPanelUiState(panelResult.state),
      },
    );

    if (!panelResult.opened) {
      console.log(
        'Automatic mode: could not open the Transcript panel automatically.',
      );
      await handleAutomaticManualAssistFallback(
        prompt,
        captureFeedback,
        debugEnabled,
        'panel-open-failed',
        panelResult,
        allowManualAssist,
      );
      return;
    }

    console.log(
      panelResult.refreshed
        ? 'Automatic mode: refreshed and reopened the Transcript panel.'
        : 'Automatic mode: opened the Transcript panel automatically.',
    );
    if (panelResult.controlLabel) {
      console.log(`Automatic mode: using control "${panelResult.controlLabel}".`);
    }

    if (
      await nudgeTranscriptPanelAutomatically(cdp, captureFeedback, debugEnabled)
    ) {
      console.log(
        'Automatic mode: nudged the Transcript panel to trigger lazy loading.',
      );
    }

    console.log('Automatic mode: waiting for transcript network traffic...');
    logAutomaticAction(
      captureFeedback,
      debugEnabled,
      'wait for transcript network traffic',
      {
        timeoutMs: defaultAutomaticSignalTimeoutMs,
      },
    );
    let transcriptSignal = await waitForTranscriptSignal(
      captureFeedback,
      defaultAutomaticSignalTimeoutMs,
    );

    if (!transcriptSignal) {
      console.log(
        'Automatic mode: no transcript traffic yet. Nudging the panel once more.',
      );
      logAutomaticAction(
        captureFeedback,
        debugEnabled,
        'transcript traffic not seen after first wait',
      );
      await nudgeTranscriptPanelAutomatically(cdp, captureFeedback, debugEnabled);
      transcriptSignal = await waitForTranscriptSignal(
        captureFeedback,
        defaultAutomaticSignalRetryTimeoutMs,
      );
    }

    if (!transcriptSignal) {
      console.log(
        'Automatic mode: no transcript traffic after opening the panel. Retrying one panel refresh.',
      );
      logAutomaticAction(
        captureFeedback,
        debugEnabled,
        'retry transcript panel refresh',
      );
      panelResult = await ensureTranscriptPanelOpenAutomatically(
        cdp,
        captureFeedback,
        debugEnabled,
        {
          refreshIfOpen: true,
        },
      );
      if (panelResult.opened) {
        await nudgeTranscriptPanelAutomatically(cdp, captureFeedback, debugEnabled);
        transcriptSignal = await waitForTranscriptSignal(
          captureFeedback,
          defaultAutomaticSignalRetryTimeoutMs,
        );
      }
    }

    if (!transcriptSignal) {
      console.log(
        'Automatic mode: transcript traffic was not detected automatically.',
      );
      await handleAutomaticManualAssistFallback(
        prompt,
        captureFeedback,
        debugEnabled,
        'traffic-not-detected',
        panelResult,
        allowManualAssist,
      );
      return;
    }

    console.log(
      `Automatic mode: observed ${transcriptSignal.transcriptHintCount} transcript-like network response` +
        `${transcriptSignal.transcriptHintCount === 1 ? '' : 's'}.`,
    );
    logAutomaticAction(
      captureFeedback,
      debugEnabled,
      'transcript traffic detected',
      transcriptSignal,
    );

    const settledSignal = await waitForTranscriptSignalToSettle(
      captureFeedback,
      defaultAutomaticRequestSettleTimeoutMs,
    );
    if (settledSignal) {
      console.log(
        'Automatic mode: transcript response activity settled. Continuing extraction.',
      );
      logAutomaticAction(
        captureFeedback,
        debugEnabled,
        'transcript activity settled',
        settledSignal,
      );
    } else {
      console.log(
        'Automatic mode: transcript traffic was seen, but the response did not fully settle before capture ended.',
      );
      logAutomaticAction(
        captureFeedback,
        debugEnabled,
        'transcript activity did not settle before timeout',
        {
          timeoutMs: defaultAutomaticRequestSettleTimeoutMs,
        },
      );
    }

    await sleep(defaultCaptureSettleMs);
  }

  return {
    runAutomaticTranscriptCaptureFlow,
  };
}
