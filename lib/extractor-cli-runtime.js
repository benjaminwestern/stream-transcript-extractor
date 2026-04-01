import {
  connectCdp as connectCoreCdp,
  evaluate as evaluateWithCdp,
  findPageTargets as findCorePageTargets,
  waitForBrowserDebugEndpoint as waitForCoreBrowserDebugEndpoint,
  waitForBrowserPageTarget as waitForCoreBrowserPageTarget,
} from './cdp.js';
import {
  chooseFromList as chooseCliFromList,
  chooseManyFromList as chooseCliManyFromList,
  createPrompt as createCliPrompt,
  parseSelectionSpec as parseCliSelectionSpec,
} from './cli.js';
import {
  ensureBrowserIsClosed as ensureCoreBrowserIsClosed,
  findAvailablePort as findCoreAvailablePort,
  launchBrowser as launchCoreBrowser,
  prepareTempProfile as prepareCoreTempProfile,
  selectBrowserAndProfile as selectCoreBrowserAndProfile,
  shutdownBrowser as shutdownCoreBrowser,
} from './browser-session.js';
import {
  APP_NAME,
  DEFAULT_BROWSER_READY_TIMEOUT_MS,
  DEFAULT_BROWSER_SHUTDOWN_TIMEOUT_MS,
  DEFAULT_CDP_HOST,
  DEFAULT_CDP_TIMEOUT_MS,
  IS_MACOS,
  IS_WINDOWS,
} from './extractor-config.js';

export class CliError extends Error {
  /**
   * @param {string} message
   * @param {number} [exitCode]
   */
  constructor(message, exitCode = 1) {
    super(message);
    this.name = 'CliError';
    this.exitCode = exitCode;
  }
}

export function normalizeBrowserKey(value) {
  return String(value).trim().toLowerCase();
}

export function createCliRuntime({
  appName = APP_NAME,
  defaultCdpHost = DEFAULT_CDP_HOST,
  defaultCdpTimeoutMs = DEFAULT_CDP_TIMEOUT_MS,
  defaultBrowserReadyTimeoutMs = DEFAULT_BROWSER_READY_TIMEOUT_MS,
  defaultBrowserShutdownTimeoutMs = DEFAULT_BROWSER_SHUTDOWN_TIMEOUT_MS,
} = {}) {
  const errorFactory = (message) => new CliError(message);

  function ensureSupportedPlatform() {
    if (!IS_MACOS && !IS_WINDOWS) {
      throw errorFactory(`${appName} currently supports macOS and Windows only.`);
    }
  }

  function createPrompt() {
    return createCliPrompt();
  }

  /**
   * @template T
   * @param {{ ask: (question: string) => Promise<string> }} prompt
   * @param {string} title
   * @param {T[]} items
   * @param {(item: T) => string} renderItem
   * @param {string} question
   * @returns {Promise<T>}
   */
  async function chooseFromList(prompt, title, items, renderItem, question) {
    return chooseCliFromList(prompt, title, items, renderItem, question);
  }

  /**
   * @template T
   * @param {{ ask: (question: string) => Promise<string> }} prompt
   * @param {string} title
   * @param {T[]} items
   * @param {(item: T) => string} renderItem
   * @param {string} question
   * @param {object} [config]
   * @returns {Promise<{ indexes: number[], items: T[] }>}
   */
  async function chooseManyFromList(
    prompt,
    title,
    items,
    renderItem,
    question,
    config = {},
  ) {
    return chooseCliManyFromList(
      prompt,
      title,
      items,
      renderItem,
      question,
      errorFactory,
      config,
    );
  }

  function parseSelectionSpec(value, itemCount) {
    return parseCliSelectionSpec(value, itemCount, errorFactory);
  }

  async function findAvailablePort(requestedPort) {
    return findCoreAvailablePort({
      requestedPort,
      host: defaultCdpHost,
      errorFactory,
    });
  }

  async function connectCdp(websocketUrl) {
    return connectCoreCdp(websocketUrl, {
      defaultTimeoutMs: defaultCdpTimeoutMs,
      errorFactory,
    });
  }

  async function findPageTargets(port) {
    return findCorePageTargets({
      host: defaultCdpHost,
      port,
      errorFactory,
    });
  }

  async function connectToPage(pageWebsocketUrl, { enableRuntime = true } = {}) {
    const cdp = await connectCdp(pageWebsocketUrl);
    if (enableRuntime) {
      await cdp.send('Runtime.enable');
    }
    return cdp;
  }

  async function evaluate(cdp, expression, timeoutMs = defaultCdpTimeoutMs) {
    return evaluateWithCdp(cdp, expression, {
      timeoutMs,
      errorFactory,
    });
  }

  async function waitForBrowserDebugEndpoint(port) {
    return waitForCoreBrowserDebugEndpoint({
      host: defaultCdpHost,
      port,
      timeoutMs: defaultBrowserReadyTimeoutMs,
      errorFactory,
    });
  }

  async function waitForBrowserPageTarget(
    port,
    timeoutMs = defaultBrowserReadyTimeoutMs,
  ) {
    return waitForCoreBrowserPageTarget({
      host: defaultCdpHost,
      port,
      timeoutMs,
      errorFactory,
    });
  }

  async function ensureBrowserIsClosed(prompt, browser) {
    return ensureCoreBrowserIsClosed(prompt, browser);
  }

  async function selectBrowserAndProfile(options, prompt, config = {}) {
    return selectCoreBrowserAndProfile(options, prompt, {
      selectFromList: config.selectFromList || chooseFromList,
      errorFactory,
    });
  }

  function prepareTempProfile(basePath, profile, tempDir) {
    return prepareCoreTempProfile(basePath, profile, tempDir);
  }

  function launchBrowser(browser, profile, tempDataDir, debugPort) {
    return launchCoreBrowser(browser, profile, tempDataDir, debugPort);
  }

  async function shutdownBrowser(browserProcess, debugPort) {
    return shutdownCoreBrowser({
      browserProcess,
      debugPort,
      host: defaultCdpHost,
      defaultTimeoutMs: defaultCdpTimeoutMs,
      errorFactory,
      shutdownTimeoutMs: defaultBrowserShutdownTimeoutMs,
    });
  }

  return {
    errorFactory,
    normalizeBrowserKey,
    ensureSupportedPlatform,
    createPrompt,
    chooseFromList,
    chooseManyFromList,
    parseSelectionSpec,
    findAvailablePort,
    connectCdp,
    findPageTargets,
    connectToPage,
    evaluate,
    waitForBrowserDebugEndpoint,
    waitForBrowserPageTarget,
    ensureBrowserIsClosed,
    selectBrowserAndProfile,
    prepareTempProfile,
    launchBrowser,
    shutdownBrowser,
  };
}
