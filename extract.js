#!/usr/bin/env bun

import { spawn, execSync } from 'node:child_process';
import { createDecipheriv } from 'node:crypto';
import {
  cpSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  rmSync,
  statSync,
  symlinkSync,
  writeFileSync,
} from 'node:fs';
import { createServer } from 'node:net';
import { homedir, platform, tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { createInterface } from 'node:readline';

const APP_NAME = 'Stream Transcript Extractor';
const APP_DESCRIPTION =
  'Extract Microsoft Teams recording transcripts from Microsoft Stream ' +
  'using your signed-in Chrome or Edge profile.';
const DEFAULT_EXTRACTOR_MODE = 'network';
const SUPPORTED_EXTRACTOR_MODES = ['network', 'automatic', 'dom'];
const SUPPORTED_BROWSER_KEYS = ['chrome', 'edge'];
const SUPPORTED_OUTPUT_FORMATS = ['json', 'md', 'both'];
const DEFAULT_OUTPUT_DIR = 'output';
const DEFAULT_CDP_HOST = '127.0.0.1';
const DEFAULT_BROWSER_READY_TIMEOUT_MS = 15_000;
const DEFAULT_CDP_TIMEOUT_MS = 30_000;
const EXTRACTION_TIMEOUT_MS = 10 * 60_000;
const DEFAULT_BROWSER_SHUTDOWN_TIMEOUT_MS = 8_000;

const EXTRACTION_SETTINGS = Object.freeze({
  dedupePrefixLength: 50,
  maxScrollIterations: 2_000,
  scrollSettleMs: 300,
  stableScrollPasses: 8,
  viewportChunkRatio: 0.8,
});

const BUILD_VERSION =
  typeof __BUILD_VERSION__ === 'string' ? __BUILD_VERSION__ : 'dev';
const BUILD_TIME =
  typeof __BUILD_TIME__ === 'string' ? __BUILD_TIME__ : '';

/**
 * Public source layout
 *
 * This repository intentionally keeps the extractor in one shareable source
 * file. The module is split into three logical sections:
 * 1. The shared CLI entrypoint and DOM fallback implementation
 * 2. The embedded network extractor, kept self-contained in a closure
 * 3. The final mode dispatcher that routes `--mode` before mode-specific
 *    argument parsing runs
 */

const CURRENT_PLATFORM = platform();
const IS_WINDOWS = CURRENT_PLATFORM === 'win32';
const IS_MACOS = CURRENT_PLATFORM === 'darwin';

/**
 * @typedef {Object} CliOptions
 * @property {string} browser
 * @property {string} profile
 * @property {string} outputName
 * @property {string} outputDir
 * @property {string} outputFormat
 * @property {number | null} debugPort
 * @property {boolean} debug
 * @property {boolean} keepBrowserOpen
 * @property {boolean} help
 * @property {boolean} version
 */

/**
 * @typedef {Object} BrowserProfile
 * @property {string} dirName
 * @property {string} displayName
 * @property {string} profileName
 * @property {string} gaiaName
 * @property {string} email
 * @property {string} path
 */

/**
 * @typedef {Object} BrowserConfig
 * @property {string} key
 * @property {string} name
 * @property {string} basePath
 * @property {string[]} binaryCandidates
 * @property {string} processName
 * @property {string} binary
 * @property {BrowserProfile[]} profiles
 */

/**
 * @typedef {Object} TranscriptEntry
 * @property {string} speaker
 * @property {string} timestamp
 * @property {string} text
 */

class CliError extends Error {
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

function sleep(ms) {
  return new Promise((resolvePromise) => setTimeout(resolvePromise, ms));
}

function normalizeExtractorMode(value) {
  return String(value || '').trim().toLowerCase();
}

/**
 * Strip the top-level `--mode` flag before handing control to the selected
 * extractor. Each mode keeps its own argument parser so the shared entrypoint
 * stays thin and easy to reason about.
 */
function resolveExtractorMode(argv) {
  let mode = DEFAULT_EXTRACTOR_MODE;
  const forwardedArgs = [];

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    const [flag, inlineValue] = arg.split(/=(.*)/s, 2);

    if (flag !== '--mode') {
      forwardedArgs.push(arg);
      continue;
    }

    const value = readOptionValue(flag, inlineValue, argv[index + 1]);
    mode = normalizeExtractorMode(value);

    if (!SUPPORTED_EXTRACTOR_MODES.includes(mode)) {
      throw new CliError(
        `Unsupported mode "${value}". Use one of: ` +
          `${SUPPORTED_EXTRACTOR_MODES.join(', ')}.`,
      );
    }

    if (inlineValue == null) {
      index += 1;
    }
  }

  return {
    mode,
    forwardedArgs,
  };
}

function readOptionValue(flag, inlineValue, nextArg) {
  const value = inlineValue ?? nextArg;

  if (
    value == null ||
    value === '' ||
    (inlineValue == null && String(value).startsWith('--'))
  ) {
    throw new CliError(`Missing value for ${flag}.`);
  }

  return value;
}

/**
 * @param {string[]} argv
 * @returns {CliOptions}
 */
function parseDomModeArgs(argv) {
  /** @type {CliOptions} */
  const options = {
    browser: '',
    profile: '',
    outputName: '',
    outputDir: DEFAULT_OUTPUT_DIR,
    outputFormat: 'json',
    debugPort: null,
    debug: false,
    keepBrowserOpen: false,
    help: false,
    version: false,
  };

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];

    if (arg === '--help' || arg === '-h') {
      options.help = true;
      continue;
    }

    if (arg === '--version' || arg === '-v') {
      options.version = true;
      continue;
    }

    if (arg === '--keep-browser-open') {
      options.keepBrowserOpen = true;
      continue;
    }

    if (arg === '--debug') {
      options.debug = true;
      continue;
    }

    if (!arg.startsWith('--')) {
      throw new CliError(`Unexpected positional argument "${arg}".`);
    }

    const [flag, inlineValue] = arg.split(/=(.*)/s, 2);
    const value = readOptionValue(flag, inlineValue, argv[index + 1]);

    if (flag === '--browser') {
      options.browser = normalizeBrowserKey(value);
    } else if (flag === '--profile') {
      options.profile = value;
    } else if (flag === '--output') {
      options.outputName = value;
    } else if (flag === '--output-dir') {
      options.outputDir = value;
    } else if (flag === '--format') {
      options.outputFormat = String(value).trim().toLowerCase();
    } else if (flag === '--debug-port') {
      const parsedPort = Number.parseInt(value, 10);
      if (!Number.isInteger(parsedPort) || parsedPort < 1 || parsedPort > 65_535) {
        throw new CliError(`Invalid port "${value}".`);
      }
      options.debugPort = parsedPort;
    } else {
      throw new CliError(`Unknown option "${flag}".`);
    }

    if (inlineValue == null) {
      index += 1;
    }
  }

  if (
    options.browser &&
    !SUPPORTED_BROWSER_KEYS.includes(options.browser)
  ) {
    throw new CliError(
      `Unsupported browser "${options.browser}". Use one of: ` +
        `${SUPPORTED_BROWSER_KEYS.join(', ')}.`,
    );
  }

  if (!SUPPORTED_OUTPUT_FORMATS.includes(options.outputFormat)) {
    throw new CliError(
      `Unsupported output format "${options.outputFormat}". Use one of: ` +
        `${SUPPORTED_OUTPUT_FORMATS.join(', ')}.`,
    );
  }

  return options;
}

function printEntrypointHelp() {
  console.log(`${APP_NAME}`);
  console.log(`${APP_DESCRIPTION}\n`);
  console.log('Usage:');
  console.log('  bun extract.js [options]');
  console.log('  ./<compiled-binary> [options]\n');
  console.log('Options:');
  console.log(
    `  --mode <network|automatic|dom>  Choose the extractor mode (default: ${DEFAULT_EXTRACTOR_MODE}).`,
  );
  console.log('  --browser <chrome|edge>  Use a specific browser.');
  console.log('  --profile <query>        Match a profile by name, email, or directory.');
  console.log('  --output <name>          Override the output filename prefix.');
  console.log('  --output-dir <path>      Write output files to a custom directory.');
  console.log('  --format <json|md|both>  Choose JSON, Markdown, or both outputs.');
  console.log('  --debug-port <port>      Force a specific remote-debugging port.');
  console.log(
    '  --debug                  Write extra diagnostics. In automatic mode this also',
  );
  console.log(
    '                           prints each UI action, retry, and fallback reason.',
  );
  console.log('  --keep-browser-open      Leave the launched browser open after extraction.');
  console.log('  --version, -v            Print the build version.');
  console.log('  --help, -h               Show this help text.\n');
  console.log('Mode flow:');
  console.log(
    '  network    Open the Stream page first with the Transcript panel closed.',
  );
  console.log(
    '             After capture is armed, open the Transcript panel and let it load.',
  );
  console.log(
    '  automatic Reload with capture armed, try to open the Transcript panel,',
  );
  console.log(
    '             nudge/retry automatically, then explain the fallback if manual',
  );
  console.log(
    '             help is still needed.',
  );
  console.log(
    '  dom        Open the Stream page and Transcript panel before extraction starts.\n',
  );
  console.log('Examples:');
  console.log('  bun extract.js');
  console.log('  bun extract.js --mode network --debug');
  console.log('  bun extract.js --mode automatic');
  console.log('  bun extract.js --mode automatic --debug');
  console.log('  bun extract.js --mode dom --format md');
}

function printEntrypointVersion() {
  const buildSuffix = BUILD_TIME ? ` (${BUILD_TIME})` : '';
  console.log(`${APP_NAME} ${BUILD_VERSION}${buildSuffix}`);
}

function normalizeBrowserKey(value) {
  return String(value).trim().toLowerCase();
}

function ensureSupportedPlatform() {
  if (!IS_MACOS && !IS_WINDOWS) {
    throw new CliError(
      `${APP_NAME} currently supports macOS and Windows only.`,
    );
  }
}

function createPrompt() {
  const readline = createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  return {
    ask(question) {
      return new Promise((resolveAnswer) =>
        readline.question(question, resolveAnswer),
      );
    },
    waitForEnter(message = 'Press Enter to continue...') {
      return new Promise((resolveAnswer) =>
        readline.question(message, () => resolveAnswer()),
      );
    },
    close() {
      readline.close();
    },
  };
}

/**
 * @param {string} title
 * @param {string[]} items
 */
function printMenu(title, items) {
  console.log(`\n${title}`);
  items.forEach((item, index) => {
    console.log(`  ${index + 1}. ${item}`);
  });
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
  printMenu(title, items.map(renderItem));

  while (true) {
    const answer = (await prompt.ask(question)).trim();
    const selectedIndex = Number.parseInt(answer, 10) - 1;
    const selectedItem = items[selectedIndex];

    if (selectedItem) {
      return selectedItem;
    }

    console.log('Enter one of the listed numbers.');
  }
}

async function findAvailablePort(requestedPort) {
  return new Promise((resolvePort, rejectPort) => {
    const server = createServer();

    server.unref();
    server.on('error', (error) => {
      const reason =
        requestedPort == null
          ? 'Unable to allocate a remote-debugging port.'
          : `Port ${requestedPort} is not available.`;
      rejectPort(new CliError(`${reason} ${error.message}`));
    });

    server.listen(requestedPort ?? 0, DEFAULT_CDP_HOST, () => {
      const address = server.address();
      if (address == null || typeof address === 'string') {
        server.close(() =>
          rejectPort(new CliError('Unable to determine the debug port.')),
        );
        return;
      }

      const { port } = address;
      server.close((closeError) => {
        if (closeError) {
          rejectPort(new CliError(closeError.message));
          return;
        }

        resolvePort(port);
      });
    });
  });
}

/**
 * A small CDP client is enough here: we only need `Runtime.evaluate`.
 *
 * @param {string} websocketUrl
 */
async function connectCdp(websocketUrl) {
  const websocket = new WebSocket(websocketUrl);
  const pendingRequests = new Map();
  let nextMessageId = 0;

  await new Promise((resolveConnection, rejectConnection) => {
    const timeout = setTimeout(() => {
      rejectConnection(new CliError('WebSocket connection timeout.'));
    }, 10_000);

    websocket.onopen = () => {
      clearTimeout(timeout);
      resolveConnection();
    };

    websocket.onerror = (event) => {
      clearTimeout(timeout);
      rejectConnection(
        new CliError(
          `WebSocket connection failed: ${event.message || 'unknown error'}.`,
        ),
      );
    };
  });

  websocket.onmessage = (event) => {
    const payload = JSON.parse(event.data);
    if (!payload.id || !pendingRequests.has(payload.id)) {
      return;
    }

    const { resolveRequest, rejectRequest, timeout } =
      pendingRequests.get(payload.id);
    clearTimeout(timeout);
    pendingRequests.delete(payload.id);

    if (payload.error) {
      rejectRequest(
        new CliError(`CDP request failed: ${payload.error.message}`),
      );
      return;
    }

    resolveRequest(payload.result);
  };

  websocket.onclose = () => {
    for (const [requestId, request] of pendingRequests.entries()) {
      clearTimeout(request.timeout);
      request.rejectRequest(
        new CliError(`CDP connection closed before request ${requestId} completed.`),
      );
      pendingRequests.delete(requestId);
    }
  };

  return {
    send(method, params = {}, timeoutMs = DEFAULT_CDP_TIMEOUT_MS) {
      const messageId = nextMessageId + 1;
      nextMessageId = messageId;

      return new Promise((resolveRequest, rejectRequest) => {
        const timeout = setTimeout(() => {
          pendingRequests.delete(messageId);
          rejectRequest(new CliError(`CDP ${method} timed out.`));
        }, timeoutMs);

        pendingRequests.set(messageId, {
          resolveRequest,
          rejectRequest,
          timeout,
        });

        websocket.send(
          JSON.stringify({
            id: messageId,
            method,
            params,
          }),
        );
      });
    },
    close() {
      for (const request of pendingRequests.values()) {
        clearTimeout(request.timeout);
      }
      pendingRequests.clear();
      websocket.close();
    },
  };
}

async function findPageTargets(port) {
  const response = await fetch(
    `http://${DEFAULT_CDP_HOST}:${port}/json/list`,
  );

  if (!response.ok) {
    throw new CliError(`CDP target discovery failed with ${response.status}.`);
  }

  const targets = await response.json();
  return targets.filter(
    (target) => target.type === 'page' && !target.url.startsWith('chrome'),
  );
}

async function connectToPage(pageWebsocketUrl) {
  const cdp = await connectCdp(pageWebsocketUrl);
  await cdp.send('Runtime.enable');
  return cdp;
}

async function evaluate(cdp, expression, timeoutMs = DEFAULT_CDP_TIMEOUT_MS) {
  const result = await cdp.send(
    'Runtime.evaluate',
    {
      expression,
      returnByValue: true,
      awaitPromise: true,
    },
    timeoutMs,
  );

  if (result.exceptionDetails) {
    throw new CliError(`Evaluation failed: ${result.exceptionDetails.text}`);
  }

  return result.result.value;
}

async function waitForBrowserDebugEndpoint(port) {
  const deadline = Date.now() + DEFAULT_BROWSER_READY_TIMEOUT_MS;

  while (Date.now() < deadline) {
    try {
      const response = await fetch(
        `http://${DEFAULT_CDP_HOST}:${port}/json/version`,
      );
      if (response.ok) {
        return;
      }
    } catch {
      // Keep polling until the deadline.
    }

    await sleep(500);
  }

  throw new CliError('Browser failed to start in debug mode.');
}

async function waitForBrowserPageTarget(
  port,
  timeoutMs = DEFAULT_BROWSER_READY_TIMEOUT_MS,
) {
  const deadline = Date.now() + timeoutMs;

  while (Date.now() < deadline) {
    try {
      const pages = await findPageTargets(port);
      if (pages.length > 0) {
        return pages;
      }
    } catch {
      // Keep polling until the deadline.
    }

    await sleep(500);
  }

  throw new CliError(
    'Browser launched in debug mode but no page window became available.',
  );
}

async function isBrowserDebugEndpointAvailable(port) {
  try {
    const response = await fetch(
      `http://${DEFAULT_CDP_HOST}:${port}/json/version`,
    );
    return response.ok;
  } catch {
    return false;
  }
}

async function waitForBrowserDebugEndpointToClose(
  port,
  timeoutMs = DEFAULT_BROWSER_SHUTDOWN_TIMEOUT_MS,
) {
  const deadline = Date.now() + timeoutMs;

  while (Date.now() < deadline) {
    const isAvailable = await isBrowserDebugEndpointAvailable(port);
    if (!isAvailable) {
      return true;
    }

    await sleep(250);
  }

  return !(await isBrowserDebugEndpointAvailable(port));
}

async function getBrowserDebuggerWebSocketUrl(port) {
  try {
    const response = await fetch(
      `http://${DEFAULT_CDP_HOST}:${port}/json/version`,
    );
    if (!response.ok) {
      return '';
    }

    const payload = await response.json();
    return String(payload.webSocketDebuggerUrl || '');
  } catch {
    return '';
  }
}

function getBrowserConfigs() {
  if (IS_WINDOWS) {
    const localAppData =
      process.env.LOCALAPPDATA || join(homedir(), 'AppData', 'Local');
    const programFiles =
      process.env.ProgramFiles || 'C:\\Program Files';
    const programFilesX86 =
      process.env['ProgramFiles(x86)'] || 'C:\\Program Files (x86)';

    return {
      chrome: {
        key: 'chrome',
        name: 'Google Chrome',
        basePath: join(localAppData, 'Google', 'Chrome', 'User Data'),
        binaryCandidates: [
          join(programFiles, 'Google', 'Chrome', 'Application', 'chrome.exe'),
          join(
            programFilesX86,
            'Google',
            'Chrome',
            'Application',
            'chrome.exe',
          ),
          join(localAppData, 'Google', 'Chrome', 'Application', 'chrome.exe'),
        ],
        processName: 'chrome.exe',
      },
      edge: {
        key: 'edge',
        name: 'Microsoft Edge',
        basePath: join(localAppData, 'Microsoft', 'Edge', 'User Data'),
        binaryCandidates: [
          join(programFiles, 'Microsoft', 'Edge', 'Application', 'msedge.exe'),
          join(
            programFilesX86,
            'Microsoft',
            'Edge',
            'Application',
            'msedge.exe',
          ),
        ],
        processName: 'msedge.exe',
      },
    };
  }

  return {
    chrome: {
      key: 'chrome',
      name: 'Google Chrome',
      basePath: join(
        homedir(),
        'Library',
        'Application Support',
        'Google',
        'Chrome',
      ),
      binaryCandidates: [
        '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
      ],
      processName: 'Google Chrome',
    },
    edge: {
      key: 'edge',
      name: 'Microsoft Edge',
      basePath: join(
        homedir(),
        'Library',
        'Application Support',
        'Microsoft Edge',
      ),
      binaryCandidates: [
        '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
      ],
      processName: 'Microsoft Edge',
    },
  };
}

function findBrowserBinary(paths) {
  for (const filePath of paths) {
    if (existsSync(filePath)) {
      return filePath;
    }
  }

  return null;
}

function readProfileInfoCache(basePath) {
  try {
    const localState = JSON.parse(
      readFileSync(join(basePath, 'Local State'), 'utf-8'),
    );
    return localState?.profile?.info_cache || {};
  } catch {
    return {};
  }
}

/**
 * @param {string} basePath
 * @returns {BrowserProfile[]}
 */
function discoverProfiles(basePath) {
  const infoCache = readProfileInfoCache(basePath);

  /** @type {BrowserProfile[]} */
  const profiles = [];

  let entries;
  try {
    entries = readdirSync(basePath);
  } catch {
    return profiles;
  }

  for (const entry of entries) {
    if (entry !== 'Default' && !entry.startsWith('Profile ')) {
      continue;
    }

    const entryPath = join(basePath, entry);

    try {
      if (!statSync(entryPath).isDirectory()) {
        continue;
      }
    } catch {
      continue;
    }

    if (
      !existsSync(join(entryPath, 'Preferences')) &&
      !existsSync(join(entryPath, 'Cookies'))
    ) {
      continue;
    }

    const info = infoCache[entry] || {};
    const profileName = info.name || '';
    const email = info.user_name || '';
    const displayName =
      profileName && email
        ? `${profileName} (${email})`
        : profileName || email || entry;

    profiles.push({
      dirName: entry,
      displayName,
      profileName,
      gaiaName: info.gaia_name || '',
      email,
      path: entryPath,
    });
  }

  return profiles;
}

/**
 * @returns {BrowserConfig[]}
 */
function discoverBrowsers() {
  const configs = getBrowserConfigs();
  const browsers = [];

  for (const config of Object.values(configs)) {
    if (!existsSync(config.basePath)) {
      continue;
    }

    const binary = findBrowserBinary(config.binaryCandidates);
    if (!binary) {
      continue;
    }

    const profiles = discoverProfiles(config.basePath);
    if (profiles.length === 0) {
      continue;
    }

    browsers.push({
      ...config,
      binary,
      profiles,
    });
  }

  return browsers;
}

function isBrowserRunning(processName) {
  try {
    if (IS_WINDOWS) {
      const output = execSync(
        `tasklist /FI "IMAGENAME eq ${processName}" /NH`,
        { stdio: 'pipe' },
      ).toString();
      return output.includes(processName);
    }

    execSync(`pgrep -x "${processName}"`, { stdio: 'pipe' });
    return true;
  } catch {
    return false;
  }
}

async function ensureBrowserIsClosed(prompt, browser) {
  while (isBrowserRunning(browser.processName)) {
    console.log(`\n${browser.name} is still running.`);
    console.log(
      'Close it so the extractor can relaunch the selected profile in debug mode.',
    );
    await prompt.waitForEnter('\nPress Enter after the browser is closed...\n');
  }
}

/**
 * @param {BrowserProfile[]} profiles
 * @param {string} rawQuery
 * @returns {BrowserProfile | undefined}
 */
function findProfileMatch(profiles, rawQuery) {
  const query = rawQuery.trim().toLowerCase();
  return profiles.find((profile) =>
    [
      profile.dirName,
      profile.displayName,
      profile.profileName,
      profile.gaiaName,
      profile.email,
    ]
      .filter(Boolean)
      .some((value) => value.toLowerCase().includes(query)),
  );
}

async function selectBrowserAndProfile(options, prompt) {
  const browsers = discoverBrowsers();

  if (browsers.length === 0) {
    throw new CliError(
      'No supported Chrome or Edge profiles were found on this machine.',
    );
  }

  let browser = null;

  if (options.browser) {
    browser = browsers.find(
      (candidate) => candidate.key === options.browser,
    );
    if (!browser) {
      throw new CliError(
        `Browser "${options.browser}" was not found. Available browsers: ` +
          `${browsers.map((candidate) => candidate.key).join(', ')}.`,
      );
    }
  } else if (browsers.length === 1) {
    browser = browsers[0];
    console.log(`Using ${browser.name}.`);
  } else {
    browser = await chooseFromList(
      prompt,
      'Available browsers',
      browsers,
      (candidate) => candidate.name,
      '\nSelect browser (number): ',
    );
  }

  let profile = null;

  if (options.profile) {
    profile = findProfileMatch(browser.profiles, options.profile);
    if (!profile) {
      throw new CliError(
        `Profile "${options.profile}" was not found for ${browser.name}.`,
      );
    }
  } else if (browser.profiles.length === 1) {
    profile = browser.profiles[0];
    console.log(`Using profile: ${profile.displayName}.`);
  } else {
    profile = await chooseFromList(
      prompt,
      `${browser.name} profiles`,
      browser.profiles,
      (candidate) => candidate.displayName,
      '\nSelect profile (number): ',
    );
  }

  return { browser, profile };
}

function prepareTempProfile(basePath, profile, tempDir) {
  const localStatePath = join(basePath, 'Local State');
  if (existsSync(localStatePath)) {
    cpSync(localStatePath, join(tempDir, 'Local State'));
  }

  if (IS_WINDOWS) {
    cpSync(profile.path, join(tempDir, profile.dirName), {
      recursive: true,
    });
    return;
  }

  symlinkSync(profile.path, join(tempDir, profile.dirName));
}

function launchBrowser(browser, profile, tempDataDir, debugPort) {
  const browserArgs = [
    `--user-data-dir=${tempDataDir}`,
    `--profile-directory=${profile.dirName}`,
    `--remote-debugging-port=${debugPort}`,
    '--no-first-run',
    '--no-default-browser-check',
    '--disable-extensions',
    '--disable-sync',
    '--new-window',
    'about:blank',
  ];

  const browserProcess = spawn(browser.binary, browserArgs, {
    stdio: 'ignore',
    detached: !IS_WINDOWS,
    windowsHide: true,
  });

  if (!IS_WINDOWS) {
    browserProcess.unref();
  }

  return browserProcess;
}

function stopProcess(pid) {
  if (!pid) {
    return;
  }

  try {
    if (IS_WINDOWS) {
      execSync(`taskkill /PID ${pid} /F /T`, { stdio: 'pipe' });
      return;
    }

    process.kill(-pid, 'SIGTERM');
  } catch {
    try {
      process.kill(pid, 'SIGTERM');
    } catch {
      // Ignore cleanup failures.
    }
  }
}

function isProcessRunning(pid) {
  if (!pid) {
    return false;
  }

  try {
    if (IS_WINDOWS) {
      const output = execSync(`tasklist /FI "PID eq ${pid}" /NH`, {
        stdio: 'pipe',
      }).toString();
      return output.trim() !== '' && !output.includes('No tasks are running');
    }

    process.kill(pid, 0);
    return true;
  } catch {
    return false;
  }
}

async function waitForProcessExit(
  pid,
  timeoutMs = DEFAULT_BROWSER_SHUTDOWN_TIMEOUT_MS,
) {
  const deadline = Date.now() + timeoutMs;

  while (Date.now() < deadline) {
    if (!isProcessRunning(pid)) {
      return true;
    }

    await sleep(250);
  }

  return !isProcessRunning(pid);
}

async function closeBrowserViaCdp(debugPort) {
  const browserWebSocketUrl = await getBrowserDebuggerWebSocketUrl(debugPort);
  if (!browserWebSocketUrl) {
    return false;
  }

  let browserCdp = null;

  try {
    browserCdp = await connectCdp(browserWebSocketUrl);
    await browserCdp.send('Browser.close', {}, 5_000);
  } catch {
    return false;
  } finally {
    if (browserCdp) {
      browserCdp.close();
    }
  }

  return waitForBrowserDebugEndpointToClose(debugPort);
}

async function shutdownBrowser(browserProcess, debugPort) {
  let closed = false;

  if (debugPort != null) {
    closed = await closeBrowserViaCdp(debugPort);
  }

  if (!closed && browserProcess?.pid) {
    stopProcess(browserProcess.pid);
  }

  if (browserProcess?.pid) {
    await waitForProcessExit(browserProcess.pid);
  }

  if (!closed && debugPort != null) {
    await waitForBrowserDebugEndpointToClose(debugPort);
  }
}

function isLikelyMeetingPage(page) {
  const haystack = [
    page?.title || '',
    page?.url || '',
  ]
    .join(' ')
    .toLowerCase();

  return /stream|sharepoint|recordings|meeting|transcript|stream\.aspx/.test(
    haystack,
  );
}

function buildTranscriptExtractionExpression(debugEnabled = false) {
  const settings = JSON.stringify(EXTRACTION_SETTINGS);
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

async function extractMeetingMetadata(cdp) {
  return evaluate(
    cdp,
    `
      (() => {
        const title =
          document.querySelector('h1')?.textContent?.trim() ||
          document.title.replace(' - Microsoft Stream', '').trim();
        const allText = document.body.innerText;
        const dateMatch = allText.match(
          /(\\d{4}-\\d{2}-\\d{2}[\\s\\d:]*(?:UTC|GMT)?)/,
        );
        const recordedByMatch = allText.match(/Recorded by\\s*\\n?\\s*(.+)/);

        return {
          title,
          date: dateMatch?.[1]?.trim() || '',
          recordedBy: recordedByMatch?.[1]?.trim() || '',
        };
      })()
    `,
  );
}

function buildOutputPayload(metadata, entries) {
  const speakers = [...new Set(entries.map((entry) => entry.speaker).filter(Boolean))];

  return {
    meeting: metadata,
    extractedAt: new Date().toISOString(),
    entryCount: entries.length,
    speakers,
    entries,
  };
}

function buildSpeakerDebugSummary(entries) {
  const speakerCounts = new Map();

  for (const entry of entries) {
    const speaker = String(entry.speaker || '').trim();
    if (!speaker) {
      continue;
    }

    speakerCounts.set(speaker, (speakerCounts.get(speaker) || 0) + 1);
  }

  const rankedSpeakers = Array.from(speakerCounts.entries())
    .sort((left, right) => right[1] - left[1])
    .map(([speaker, count]) => ({
      speaker,
      count,
    }));

  const suspiciousSpeakers = rankedSpeakers.filter(({ speaker }) => {
    const wordCount = speaker.split(/\s+/).filter(Boolean).length;
    return (
      /[?!,:;]/.test(speaker) ||
      speaker.length > 80 ||
      wordCount > 6 ||
      (/[.]$/.test(speaker) && !/^(?:[A-Z]\.?|[A-Z][a-z]+\.)$/.test(speaker))
    );
  });

  return {
    uniqueSpeakerCount: speakerCounts.size,
    mostCommonSpeakers: rankedSpeakers.slice(0, 20),
    suspiciousSpeakers: suspiciousSpeakers.slice(0, 20),
  };
}

function buildDebugPayload({
  options,
  browser,
  profile,
  debugPort,
  targetPage,
  extractionResult,
  metadata,
  entries,
  error,
}) {
  return {
    app: {
      name: APP_NAME,
      version: BUILD_VERSION,
      buildTime: BUILD_TIME,
      platform: CURRENT_PLATFORM,
    },
    run: {
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
      targetPage: targetPage
        ? {
            title: targetPage.title,
            url: targetPage.url,
          }
        : null,
      outputFormat: options.outputFormat,
      debugPort,
      extractedAt: new Date().toISOString(),
    },
    error: error || '',
    metadata: metadata || null,
    extraction: extractionResult
      ? {
          scrollCount: extractionResult.scrollCount,
          lastTimestamp: extractionResult.lastTimestamp,
          error: extractionResult.error || '',
          debug: extractionResult.debug || null,
        }
      : null,
    summary: {
      entryCount: entries.length,
      speakerAnalysis: buildSpeakerDebugSummary(entries),
      firstEntries: entries.slice(0, 5),
      lastEntries: entries.slice(-5),
    },
  };
}

function sanitizeFilename(value) {
  const sanitized = value
    .replace(/[^a-zA-Z0-9-_ ]/g, '')
    .replace(/\s+/g, '_')
    .slice(0, 80);

  return sanitized || 'meeting';
}

function resolveOutputDirectory(outputDir) {
  return resolve(process.cwd(), outputDir || DEFAULT_OUTPUT_DIR);
}

function buildOutputBasePath(defaultName, outputName, outputDir) {
  const absoluteOutputDir = resolveOutputDirectory(outputDir);
  mkdirSync(absoluteOutputDir, { recursive: true });

  const timestamp = new Date()
    .toISOString()
    .replace(/[:.]/g, '-')
    .slice(0, 19);

  const baseName = sanitizeFilename(outputName || defaultName || 'meeting');
  return join(absoluteOutputDir, `${baseName}_${timestamp}`);
}

function buildMarkdownOutput(payload) {
  const metadataLines = [
    `Title: ${payload.meeting.title || ''}`,
    `Date: ${payload.meeting.date || ''}`,
    `Recorded by: ${payload.meeting.recordedBy || ''}`,
    `Extracted at: ${payload.extractedAt}`,
    `Entry count: ${payload.entryCount}`,
  ];

  const entryBlocks = payload.entries.map((entry) => {
    const speaker = entry.speaker || 'Unknown speaker';
    const timestamp = entry.timestamp || '';
    const heading = `${speaker} - ${timestamp}:`;
    const text = String(entry.text || '')
      .replace(/\r\n/g, '\n')
      .replace(/\n{2,}/g, '\n')
      .trim();

    return `${heading}\n${text}`;
  });

  return [
    'Transcript information',
    '',
    ...metadataLines,
    '',
    '---',
    '',
    entryBlocks.join('\n\n'),
    '',
  ].join('\n');
}

function saveOutputs(payload, outputBasePath, outputFormat) {
  const transcriptPaths = [];

  if (outputFormat === 'json' || outputFormat === 'both') {
    const jsonPath = `${outputBasePath}.json`;
    writeFileSync(jsonPath, JSON.stringify(payload, null, 2));
    transcriptPaths.push(jsonPath);
  }

  if (outputFormat === 'md' || outputFormat === 'both') {
    const markdownPath = `${outputBasePath}.md`;
    writeFileSync(markdownPath, buildMarkdownOutput(payload));
    transcriptPaths.push(markdownPath);
  }

  return transcriptPaths;
}

function saveDebugOutput(debugPayload, outputBasePath) {
  const debugPath = `${outputBasePath}.debug.json`;
  writeFileSync(debugPath, JSON.stringify(debugPayload, null, 2));
  return debugPath;
}

async function selectTranscriptPage(prompt, debugPort) {
  const pages = await findPageTargets(debugPort);

  if (pages.length === 0) {
    throw new CliError(
      'No browser pages were found. Open Microsoft Stream before continuing.',
    );
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
    'Open pages',
    pages,
    (page) => `${page.title} (${page.url.slice(0, 80)})`,
    '\nWhich page contains the transcript? (number): ',
  );
}

function printRunInstructions() {
  console.log('\nBrowser is ready.');
  console.log('1. Open the meeting in Microsoft Stream.');
  console.log('2. Open the Transcript panel.');
  console.log('3. Return to this terminal and continue.');
}

function handleEntrypointError(error) {
  if (error instanceof CliError) {
    console.error(`\nError: ${error.message}`);
    return error.exitCode;
  }

  console.error('\nUnexpected error:');
  console.error(error);
  return 1;
}

async function runDomMode() {
  const options = parseDomModeArgs(process.argv.slice(2));

  if (options.help) {
    printEntrypointHelp();
    return 0;
  }

  if (options.version) {
    printEntrypointVersion();
    return 0;
  }

  ensureSupportedPlatform();

  const prompt = createPrompt();
  let browserProcess = null;
  let cdp = null;
  let tempDataDir = null;
  let browser = null;
  let profile = null;
  let debugPort = null;
  let targetPage = null;
  let extractionResult = null;

  try {
    ({ browser, profile } = await selectBrowserAndProfile(options, prompt));
    debugPort = await findAvailablePort(options.debugPort);

    console.log(`\nUsing ${browser.name} / "${profile.displayName}".`);
    await ensureBrowserIsClosed(prompt, browser);

    tempDataDir = join(tmpdir(), `stream-transcript-extractor-${Date.now()}`);
    mkdirSync(tempDataDir, { recursive: true });

    console.log('Preparing temporary browser profile...');
    prepareTempProfile(browser.basePath, profile, tempDataDir);

    console.log(`Launching ${browser.name} on debug port ${debugPort}...`);
    browserProcess = launchBrowser(browser, profile, tempDataDir, debugPort);

    console.log('Waiting for the browser debug endpoint...');
    await waitForBrowserDebugEndpoint(debugPort);
    await waitForBrowserPageTarget(debugPort);

    printRunInstructions();
    await prompt.waitForEnter('\nPress Enter to start extracting...\n');

    targetPage = await selectTranscriptPage(prompt, debugPort);
    console.log(`\nConnecting to: ${targetPage.title}`);

    cdp = await connectToPage(targetPage.webSocketDebuggerUrl);

    console.log('Scrolling the transcript and extracting entries...');
    extractionResult = await evaluate(
      cdp,
      buildTranscriptExtractionExpression(options.debug),
      EXTRACTION_TIMEOUT_MS,
    );

    if (extractionResult.error) {
      if (options.debug) {
        const failureOutputBasePath = buildOutputBasePath(
          targetPage?.title || 'meeting',
          options.outputName,
          options.outputDir,
        );
        const debugPayload = buildDebugPayload({
          options,
          browser,
          profile,
          debugPort,
          targetPage,
          extractionResult,
          metadata: null,
          entries: extractionResult.entries || [],
          error: extractionResult.error,
        });
        const debugPath = saveDebugOutput(debugPayload, failureOutputBasePath);
        console.log(`Saved debug output to: ${debugPath}`);
      }

      throw new CliError(extractionResult.error);
    }

    const entries = extractionResult.entries;
    console.log(
      `Scroll passes: ${extractionResult.scrollCount}, ` +
        `last timestamp: ${extractionResult.lastTimestamp || 'n/a'}.`,
    );
    if (options.debug && extractionResult.debug) {
      console.log(
        `Debug strategy counts: cells=${extractionResult.debug.strategyCounts.cells}, ` +
          `text=${extractionResult.debug.strategyCounts.text}.`,
      );
    }

    if (!entries.length) {
      if (options.debug) {
        const failureOutputBasePath = buildOutputBasePath(
          targetPage?.title || 'meeting',
          options.outputName,
          options.outputDir,
        );
        const debugPayload = buildDebugPayload({
          options,
          browser,
          profile,
          debugPort,
          targetPage,
          extractionResult,
          metadata: null,
          entries,
          error:
            'No transcript entries were found. Confirm the transcript panel is open and visible.',
        });
        const debugPath = saveDebugOutput(debugPayload, failureOutputBasePath);
        console.log(`Saved debug output to: ${debugPath}`);
      }

      throw new CliError(
        'No transcript entries were found. Confirm the transcript panel is open and visible.',
      );
    }

    const metadata = await extractMeetingMetadata(cdp);
    const outputPayload = buildOutputPayload(metadata, entries);
    const outputBasePath = buildOutputBasePath(
      outputPayload.meeting.title,
      options.outputName,
      options.outputDir,
    );
    const outputPaths = saveOutputs(outputPayload, outputBasePath, options.outputFormat);
    let debugPath = '';

    if (options.debug) {
      const debugPayload = buildDebugPayload({
        options,
        browser,
        profile,
        debugPort,
        targetPage,
        extractionResult,
        metadata,
        entries,
        error: '',
      });
      debugPath = saveDebugOutput(debugPayload, outputBasePath);
    }

    console.log(`Extracted ${entries.length} entries.`);
    console.log(`Speakers: ${outputPayload.speakers.join(', ')}`);
    for (const outputPath of outputPaths) {
      console.log(`Saved transcript to: ${outputPath}`);
    }
    if (debugPath) {
      console.log(`Saved debug output to: ${debugPath}`);
    }

    return 0;
  } catch (error) {
    return handleError(error);
  } finally {
    prompt.close();

    if (cdp) {
      cdp.close();
    }

    if (!options.keepBrowserOpen && browserProcess?.pid) {
      await shutdownBrowser(browserProcess, debugPort);
    }

    if (options.keepBrowserOpen && tempDataDir) {
      console.log(
        `Temporary profile preserved because --keep-browser-open is set: ${tempDataDir}`,
      );
    } else if (tempDataDir) {
      try {
        rmSync(tempDataDir, { recursive: true, force: true });
      } catch {
        // Ignore cleanup failures.
      }
    }
  }
}

/**
 * Network mode stays wrapped in a closure so it can remain a complete,
 * shareable extractor implementation without leaking its internal helpers into
 * the DOM mode namespace.
 */
const runEmbeddedNetworkMode = (() => {
  const APP_NAME = 'Stream Transcript Extractor (Network Mode)';
  const APP_DESCRIPTION =
    'Extract Microsoft Teams recording transcripts from Microsoft Stream ' +
    'using your signed-in Chrome or Edge profile.';
  const SUPPORTED_BROWSER_KEYS = ['chrome', 'edge'];
  const SUPPORTED_OUTPUT_FORMATS = ['json', 'md', 'both'];
  const DEFAULT_OUTPUT_DIR = 'output';
  const DEFAULT_CDP_HOST = '127.0.0.1';
  const DEFAULT_BROWSER_READY_TIMEOUT_MS = 15_000;
  const DEFAULT_CDP_TIMEOUT_MS = 30_000;
  const DEFAULT_CAPTURE_SETTLE_MS = 1_500;
  const DEFAULT_BROWSER_SHUTDOWN_TIMEOUT_MS = 8_000;
  const DEFAULT_AUTOMATIC_UI_POLL_MS = 250;
  const DEFAULT_AUTOMATIC_PANEL_OPEN_TIMEOUT_MS = 7_500;
  const DEFAULT_AUTOMATIC_SIGNAL_TIMEOUT_MS = 8_000;
  const DEFAULT_AUTOMATIC_SIGNAL_RETRY_TIMEOUT_MS = 6_000;
  const DEFAULT_AUTOMATIC_REQUEST_SETTLE_TIMEOUT_MS = 8_000;
  const DEFAULT_AUTOMATIC_CLICK_SETTLE_MS = 750;
  const MAX_TRACKED_RESPONSES = 150;
  const MAX_CANDIDATE_PREVIEW_LENGTH = 4_000;
  const MAX_SAVED_CANDIDATES = 40;
  const LIVE_TRANSCRIPT_SIGNAL_SCORE = 6;
  const MAX_DEBUG_REQUESTS = 1_000;
  const MAX_DEBUG_RESPONSES = 1_000;
  const MAX_DEBUG_BODIES = 150;
  const MAX_DEBUG_WEBSOCKETS = 100;
  const MAX_DEBUG_WEBSOCKET_FRAMES = 400;
  const MAX_DEBUG_FRAME_PREVIEW_LENGTH = 4_000;
  const MAX_DEBUG_POST_DATA_PREVIEW_LENGTH = 4_000;
  
  const BUILD_VERSION =
    typeof __BUILD_VERSION__ === 'string' ? __BUILD_VERSION__ : 'dev';
  const BUILD_TIME =
    typeof __BUILD_TIME__ === 'string' ? __BUILD_TIME__ : '';
  
  const CURRENT_PLATFORM = platform();
  const IS_WINDOWS = CURRENT_PLATFORM === 'win32';
  const IS_MACOS = CURRENT_PLATFORM === 'darwin';
  
  /**
   * @typedef {Object} CliOptions
   * @property {string} browser
   * @property {string} profile
   * @property {string} outputName
   * @property {string} outputDir
   * @property {string} outputFormat
   * @property {number | null} debugPort
   * @property {boolean} debug
   * @property {boolean} keepBrowserOpen
   * @property {boolean} help
   * @property {boolean} version
   */
  
  /**
   * @typedef {Object} BrowserProfile
   * @property {string} dirName
   * @property {string} displayName
   * @property {string} profileName
   * @property {string} gaiaName
   * @property {string} email
   * @property {string} path
   */
  
  /**
   * @typedef {Object} BrowserConfig
   * @property {string} key
   * @property {string} name
   * @property {string} basePath
   * @property {string[]} binaryCandidates
   * @property {string} processName
   * @property {string} binary
   * @property {BrowserProfile[]} profiles
   */
  
  /**
   * @typedef {Object} TranscriptEntry
   * @property {string} speaker
   * @property {string} timestamp
   * @property {string} text
   */
  
  class CliError extends Error {
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
  
  function sleep(ms) {
    return new Promise((resolvePromise) => setTimeout(resolvePromise, ms));
  }
  
  function readOptionValue(flag, inlineValue, nextArg) {
    const value = inlineValue ?? nextArg;
  
    if (
      value == null ||
      value === '' ||
      (inlineValue == null && String(value).startsWith('--'))
    ) {
      throw new CliError(`Missing value for ${flag}.`);
    }
  
    return value;
  }
  
  /**
   * @param {string[]} argv
   * @returns {CliOptions}
   */
  function parseArgs(argv) {
    /** @type {CliOptions} */
    const options = {
      browser: '',
      profile: '',
      outputName: '',
      outputDir: DEFAULT_OUTPUT_DIR,
      outputFormat: 'json',
      debugPort: null,
      debug: false,
      keepBrowserOpen: false,
      help: false,
      version: false,
    };
  
    for (let index = 0; index < argv.length; index += 1) {
      const arg = argv[index];
  
      if (arg === '--help' || arg === '-h') {
        options.help = true;
        continue;
      }
  
      if (arg === '--version' || arg === '-v') {
        options.version = true;
        continue;
      }
  
      if (arg === '--keep-browser-open') {
        options.keepBrowserOpen = true;
        continue;
      }
  
      if (arg === '--debug') {
        options.debug = true;
        continue;
      }
  
      if (!arg.startsWith('--')) {
        throw new CliError(`Unexpected positional argument "${arg}".`);
      }
  
      const [flag, inlineValue] = arg.split(/=(.*)/s, 2);
      const value = readOptionValue(flag, inlineValue, argv[index + 1]);
  
      if (flag === '--browser') {
        options.browser = normalizeBrowserKey(value);
      } else if (flag === '--profile') {
        options.profile = value;
      } else if (flag === '--output') {
        options.outputName = value;
      } else if (flag === '--output-dir') {
        options.outputDir = value;
      } else if (flag === '--format') {
        options.outputFormat = String(value).trim().toLowerCase();
      } else if (flag === '--debug-port') {
        const parsedPort = Number.parseInt(value, 10);
        if (!Number.isInteger(parsedPort) || parsedPort < 1 || parsedPort > 65_535) {
          throw new CliError(`Invalid port "${value}".`);
        }
        options.debugPort = parsedPort;
      } else {
        throw new CliError(`Unknown option "${flag}".`);
      }
  
      if (inlineValue == null) {
        index += 1;
      }
    }
  
    if (
      options.browser &&
      !SUPPORTED_BROWSER_KEYS.includes(options.browser)
    ) {
      throw new CliError(
        `Unsupported browser "${options.browser}". Use one of: ` +
          `${SUPPORTED_BROWSER_KEYS.join(', ')}.`,
      );
    }
  
    if (!SUPPORTED_OUTPUT_FORMATS.includes(options.outputFormat)) {
      throw new CliError(
        `Unsupported output format "${options.outputFormat}". Use one of: ` +
          `${SUPPORTED_OUTPUT_FORMATS.join(', ')}.`,
      );
    }
  
    return options;
  }
  
  function printHelp() {
    console.log(`${APP_NAME}`);
    console.log(`${APP_DESCRIPTION}\n`);
    console.log('Usage:');
    console.log('  bun extract.js --mode <network|automatic> [options]');
    console.log('  ./<compiled-binary> [options]\n');
    console.log('Options:');
    console.log('  --browser <chrome|edge>  Use a specific browser.');
    console.log('  --profile <query>        Match a profile by name, email, or directory.');
    console.log('  --output <name>          Override the output filename prefix.');
    console.log('  --output-dir <path>      Write files to a custom directory.');
    console.log('  --format <json|md|both>  Choose JSON, Markdown, or both outputs.');
    console.log('  --debug-port <port>      Force a specific remote-debugging port.');
    console.log(
      '  --debug                  Save extended network diagnostics, including ' +
        'request/response lifecycle data, candidate bodies, and WebSocket frames.',
    );
    console.log(
      '                           In automatic mode, --debug also prints each UI',
    );
    console.log(
      '                           action, retry, and fallback reason.',
    );
    console.log('  --keep-browser-open      Leave the launched browser open after extraction.');
    console.log('  --version, -v            Print the build version.');
    console.log('  --help, -h               Show this help text.\n');
    console.log('Capture flow:');
    console.log(
      '  network: open the Stream page first with the Transcript panel closed.',
    );
    console.log(
      '  automatic: let the extractor reload the page, try the Transcript panel',
    );
    console.log(
      '             automatically, retry the panel actions, then explain any',
    );
    console.log(
      '             manual fallback before it asks for help.',
    );
  }
  
  function printVersion() {
    const buildSuffix = BUILD_TIME ? ` (${BUILD_TIME})` : '';
    console.log(`${APP_NAME} ${BUILD_VERSION}${buildSuffix}`);
  }
  
  function normalizeBrowserKey(value) {
    return String(value).trim().toLowerCase();
  }
  
  function ensureSupportedPlatform() {
    if (!IS_MACOS && !IS_WINDOWS) {
      throw new CliError(
        `${APP_NAME} currently supports macOS and Windows only.`,
      );
    }
  }
  
  function createPrompt() {
    const readline = createInterface({
      input: process.stdin,
      output: process.stdout,
    });
  
    return {
      ask(question) {
        return new Promise((resolveAnswer) =>
          readline.question(question, resolveAnswer),
        );
      },
      waitForEnter(message = 'Press Enter to continue...') {
        return new Promise((resolveAnswer) =>
          readline.question(message, () => resolveAnswer()),
        );
      },
      close() {
        readline.close();
      },
    };
  }
  
  /**
   * @param {string} title
   * @param {string[]} items
   */
  function printMenu(title, items) {
    console.log(`\n${title}`);
    items.forEach((item, index) => {
      console.log(`  ${index + 1}. ${item}`);
    });
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
    printMenu(title, items.map(renderItem));
  
    while (true) {
      const answer = (await prompt.ask(question)).trim();
      const selectedIndex = Number.parseInt(answer, 10) - 1;
      const selectedItem = items[selectedIndex];
  
      if (selectedItem) {
        return selectedItem;
      }
  
      console.log('Enter one of the listed numbers.');
    }
  }
  
  async function findAvailablePort(requestedPort) {
    return new Promise((resolvePort, rejectPort) => {
      const server = createServer();
  
      server.unref();
      server.on('error', (error) => {
        const reason =
          requestedPort == null
            ? 'Unable to allocate a remote-debugging port.'
            : `Port ${requestedPort} is not available.`;
        rejectPort(new CliError(`${reason} ${error.message}`));
      });
  
      server.listen(requestedPort ?? 0, DEFAULT_CDP_HOST, () => {
        const address = server.address();
        if (address == null || typeof address === 'string') {
          server.close(() =>
            rejectPort(new CliError('Unable to determine the debug port.')),
          );
          return;
        }
  
        const { port } = address;
        server.close((closeError) => {
          if (closeError) {
            rejectPort(new CliError(closeError.message));
            return;
          }
  
          resolvePort(port);
        });
      });
    });
  }
  
  /**
   * @param {string} websocketUrl
   */
  async function connectCdp(websocketUrl) {
    const websocket = new WebSocket(websocketUrl);
    const pendingRequests = new Map();
    const eventHandlers = new Map();
    let nextMessageId = 0;
  
    await new Promise((resolveConnection, rejectConnection) => {
      const timeout = setTimeout(() => {
        rejectConnection(new CliError('WebSocket connection timeout.'));
      }, 10_000);
  
      websocket.onopen = () => {
        clearTimeout(timeout);
        resolveConnection();
      };
  
      websocket.onerror = (event) => {
        clearTimeout(timeout);
        rejectConnection(
          new CliError(
            `WebSocket connection failed: ${event.message || 'unknown error'}.`,
          ),
        );
      };
    });
  
    websocket.onmessage = (event) => {
      const payload = JSON.parse(event.data);
  
      if (payload.id && pendingRequests.has(payload.id)) {
        const { resolveRequest, rejectRequest, timeout } =
          pendingRequests.get(payload.id);
        clearTimeout(timeout);
        pendingRequests.delete(payload.id);
  
        if (payload.error) {
          rejectRequest(
            new CliError(`CDP request failed: ${payload.error.message}`),
          );
          return;
        }
  
        resolveRequest(payload.result);
        return;
      }
  
      if (!payload.method) {
        return;
      }
  
      const handlers = eventHandlers.get(payload.method);
      if (!handlers || handlers.size === 0) {
        return;
      }
  
      for (const handler of handlers) {
        try {
          handler(payload.params || {});
        } catch {
          // Ignore individual event handler failures so capture can continue.
        }
      }
    };
  
    websocket.onclose = () => {
      for (const [requestId, request] of pendingRequests.entries()) {
        clearTimeout(request.timeout);
        request.rejectRequest(
          new CliError(`CDP connection closed before request ${requestId} completed.`),
        );
        pendingRequests.delete(requestId);
      }
    };
  
    return {
      send(method, params = {}, timeoutMs = DEFAULT_CDP_TIMEOUT_MS) {
        const messageId = nextMessageId + 1;
        nextMessageId = messageId;
  
        return new Promise((resolveRequest, rejectRequest) => {
          const timeout = setTimeout(() => {
            pendingRequests.delete(messageId);
            rejectRequest(new CliError(`CDP ${method} timed out.`));
          }, timeoutMs);
  
          pendingRequests.set(messageId, {
            resolveRequest,
            rejectRequest,
            timeout,
          });
  
          websocket.send(
            JSON.stringify({
              id: messageId,
              method,
              params,
            }),
          );
        });
      },
      on(method, handler) {
        if (!eventHandlers.has(method)) {
          eventHandlers.set(method, new Set());
        }
  
        const handlers = eventHandlers.get(method);
        handlers.add(handler);
  
        return () => {
          handlers.delete(handler);
          if (handlers.size === 0) {
            eventHandlers.delete(method);
          }
        };
      },
      close() {
        for (const request of pendingRequests.values()) {
          clearTimeout(request.timeout);
        }
        pendingRequests.clear();
        eventHandlers.clear();
        websocket.close();
      },
    };
  }
  
  async function findPageTargets(port) {
    const response = await fetch(
      `http://${DEFAULT_CDP_HOST}:${port}/json/list`,
    );
  
    if (!response.ok) {
      throw new CliError(`CDP target discovery failed with ${response.status}.`);
    }
  
    const targets = await response.json();
    return targets.filter(
      (target) => target.type === 'page' && !target.url.startsWith('chrome'),
    );
  }
  
  async function connectToPage(pageWebsocketUrl) {
    return connectCdp(pageWebsocketUrl);
  }
  
  async function evaluate(cdp, expression, timeoutMs = DEFAULT_CDP_TIMEOUT_MS) {
    const result = await cdp.send(
      'Runtime.evaluate',
      {
        expression,
        returnByValue: true,
        awaitPromise: true,
      },
      timeoutMs,
    );
  
    if (result.exceptionDetails) {
      throw new CliError(`Evaluation failed: ${result.exceptionDetails.text}`);
    }
  
    return result.result.value;
  }
  
  async function waitForBrowserDebugEndpoint(port) {
    const deadline = Date.now() + DEFAULT_BROWSER_READY_TIMEOUT_MS;
  
    while (Date.now() < deadline) {
      try {
        const response = await fetch(
          `http://${DEFAULT_CDP_HOST}:${port}/json/version`,
        );
        if (response.ok) {
          return;
        }
      } catch {
        // Keep polling until the deadline.
      }
  
      await sleep(500);
    }
  
    throw new CliError('Browser failed to start in debug mode.');
  }
  
  async function waitForBrowserPageTarget(
    port,
    timeoutMs = DEFAULT_BROWSER_READY_TIMEOUT_MS,
  ) {
    const deadline = Date.now() + timeoutMs;
  
    while (Date.now() < deadline) {
      try {
        const pages = await findPageTargets(port);
        if (pages.length > 0) {
          return pages;
        }
      } catch {
        // Keep polling until the deadline.
      }
  
      await sleep(500);
    }
  
    throw new CliError(
      'Browser launched in debug mode but no page window became available.',
    );
  }
  
  async function isBrowserDebugEndpointAvailable(port) {
    try {
      const response = await fetch(
        `http://${DEFAULT_CDP_HOST}:${port}/json/version`,
      );
      return response.ok;
    } catch {
      return false;
    }
  }
  
  async function waitForBrowserDebugEndpointToClose(
    port,
    timeoutMs = DEFAULT_BROWSER_SHUTDOWN_TIMEOUT_MS,
  ) {
    const deadline = Date.now() + timeoutMs;
  
    while (Date.now() < deadline) {
      const isAvailable = await isBrowserDebugEndpointAvailable(port);
      if (!isAvailable) {
        return true;
      }
  
      await sleep(250);
    }
  
    return !(await isBrowserDebugEndpointAvailable(port));
  }
  
  async function getBrowserDebuggerWebSocketUrl(port) {
    try {
      const response = await fetch(
        `http://${DEFAULT_CDP_HOST}:${port}/json/version`,
      );
      if (!response.ok) {
        return '';
      }
  
      const payload = await response.json();
      return String(payload.webSocketDebuggerUrl || '');
    } catch {
      return '';
    }
  }
  
  function getBrowserConfigs() {
    if (IS_WINDOWS) {
      const localAppData =
        process.env.LOCALAPPDATA || join(homedir(), 'AppData', 'Local');
      const programFiles =
        process.env.ProgramFiles || 'C:\\Program Files';
      const programFilesX86 =
        process.env['ProgramFiles(x86)'] || 'C:\\Program Files (x86)';
  
      return {
        chrome: {
          key: 'chrome',
          name: 'Google Chrome',
          basePath: join(localAppData, 'Google', 'Chrome', 'User Data'),
          binaryCandidates: [
            join(programFiles, 'Google', 'Chrome', 'Application', 'chrome.exe'),
            join(
              programFilesX86,
              'Google',
              'Chrome',
              'Application',
              'chrome.exe',
            ),
            join(localAppData, 'Google', 'Chrome', 'Application', 'chrome.exe'),
          ],
          processName: 'chrome.exe',
        },
        edge: {
          key: 'edge',
          name: 'Microsoft Edge',
          basePath: join(localAppData, 'Microsoft', 'Edge', 'User Data'),
          binaryCandidates: [
            join(programFiles, 'Microsoft', 'Edge', 'Application', 'msedge.exe'),
            join(
              programFilesX86,
              'Microsoft',
              'Edge',
              'Application',
              'msedge.exe',
            ),
          ],
          processName: 'msedge.exe',
        },
      };
    }
  
    return {
      chrome: {
        key: 'chrome',
        name: 'Google Chrome',
        basePath: join(
          homedir(),
          'Library',
          'Application Support',
          'Google',
          'Chrome',
        ),
        binaryCandidates: [
          '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
        ],
        processName: 'Google Chrome',
      },
      edge: {
        key: 'edge',
        name: 'Microsoft Edge',
        basePath: join(
          homedir(),
          'Library',
          'Application Support',
          'Microsoft Edge',
        ),
        binaryCandidates: [
          '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
        ],
        processName: 'Microsoft Edge',
      },
    };
  }
  
  function findBrowserBinary(paths) {
    for (const filePath of paths) {
      if (existsSync(filePath)) {
        return filePath;
      }
    }
  
    return null;
  }
  
  function readProfileInfoCache(basePath) {
    try {
      const localState = JSON.parse(
        readFileSync(join(basePath, 'Local State'), 'utf-8'),
      );
      return localState?.profile?.info_cache || {};
    } catch {
      return {};
    }
  }
  
  /**
   * @param {string} basePath
   * @returns {BrowserProfile[]}
   */
  function discoverProfiles(basePath) {
    const infoCache = readProfileInfoCache(basePath);
  
    /** @type {BrowserProfile[]} */
    const profiles = [];
  
    let entries;
    try {
      entries = readdirSync(basePath);
    } catch {
      return profiles;
    }
  
    for (const entry of entries) {
      if (entry !== 'Default' && !entry.startsWith('Profile ')) {
        continue;
      }
  
      const entryPath = join(basePath, entry);
  
      try {
        if (!statSync(entryPath).isDirectory()) {
          continue;
        }
      } catch {
        continue;
      }
  
      if (
        !existsSync(join(entryPath, 'Preferences')) &&
        !existsSync(join(entryPath, 'Cookies'))
      ) {
        continue;
      }
  
      const info = infoCache[entry] || {};
      const profileName = info.name || '';
      const email = info.user_name || '';
      const displayName =
        profileName && email
          ? `${profileName} (${email})`
          : profileName || email || entry;
  
      profiles.push({
        dirName: entry,
        displayName,
        profileName,
        gaiaName: info.gaia_name || '',
        email,
        path: entryPath,
      });
    }
  
    return profiles;
  }
  
  /**
   * @returns {BrowserConfig[]}
   */
  function discoverBrowsers() {
    const configs = getBrowserConfigs();
    const browsers = [];
  
    for (const config of Object.values(configs)) {
      if (!existsSync(config.basePath)) {
        continue;
      }
  
      const binary = findBrowserBinary(config.binaryCandidates);
      if (!binary) {
        continue;
      }
  
      const profiles = discoverProfiles(config.basePath);
      if (profiles.length === 0) {
        continue;
      }
  
      browsers.push({
        ...config,
        binary,
        profiles,
      });
    }
  
    return browsers;
  }
  
  function isBrowserRunning(processName) {
    try {
      if (IS_WINDOWS) {
        const output = execSync(
          `tasklist /FI "IMAGENAME eq ${processName}" /NH`,
          { stdio: 'pipe' },
        ).toString();
        return output.includes(processName);
      }
  
      execSync(`pgrep -x "${processName}"`, { stdio: 'pipe' });
      return true;
    } catch {
      return false;
    }
  }
  
  async function ensureBrowserIsClosed(prompt, browser) {
    while (isBrowserRunning(browser.processName)) {
      console.log(`\n${browser.name} is still running.`);
      console.log(
        'Close it so the extractor can relaunch the selected profile in debug mode.',
      );
      await prompt.waitForEnter('\nPress Enter after the browser is closed...\n');
    }
  }
  
  /**
   * @param {BrowserProfile[]} profiles
   * @param {string} rawQuery
   * @returns {BrowserProfile | undefined}
   */
  function findProfileMatch(profiles, rawQuery) {
    const query = rawQuery.trim().toLowerCase();
    return profiles.find((profile) =>
      [
        profile.dirName,
        profile.displayName,
        profile.profileName,
        profile.gaiaName,
        profile.email,
      ]
        .filter(Boolean)
        .some((value) => value.toLowerCase().includes(query)),
    );
  }
  
  async function selectBrowserAndProfile(options, prompt) {
    const browsers = discoverBrowsers();
  
    if (browsers.length === 0) {
      throw new CliError(
        'No supported Chrome or Edge profiles were found on this machine.',
      );
    }
  
    let browser = null;
  
    if (options.browser) {
      browser = browsers.find(
        (candidate) => candidate.key === options.browser,
      );
      if (!browser) {
        throw new CliError(
          `Browser "${options.browser}" was not found. Available browsers: ` +
            `${browsers.map((candidate) => candidate.key).join(', ')}.`,
        );
      }
    } else if (browsers.length === 1) {
      browser = browsers[0];
      console.log(`Using ${browser.name}.`);
    } else {
      browser = await chooseFromList(
        prompt,
        'Available browsers',
        browsers,
        (candidate) => candidate.name,
        '\nSelect browser (number): ',
      );
    }
  
    let profile = null;
  
    if (options.profile) {
      profile = findProfileMatch(browser.profiles, options.profile);
      if (!profile) {
        throw new CliError(
          `Profile "${options.profile}" was not found for ${browser.name}.`,
        );
      }
    } else if (browser.profiles.length === 1) {
      profile = browser.profiles[0];
      console.log(`Using profile: ${profile.displayName}.`);
    } else {
      profile = await chooseFromList(
        prompt,
        `${browser.name} profiles`,
        browser.profiles,
        (candidate) => candidate.displayName,
        '\nSelect profile (number): ',
      );
    }
  
    return { browser, profile };
  }
  
  function prepareTempProfile(basePath, profile, tempDir) {
    const localStatePath = join(basePath, 'Local State');
    if (existsSync(localStatePath)) {
      cpSync(localStatePath, join(tempDir, 'Local State'));
    }
  
    if (IS_WINDOWS) {
      cpSync(profile.path, join(tempDir, profile.dirName), {
        recursive: true,
      });
      return;
    }
  
    symlinkSync(profile.path, join(tempDir, profile.dirName));
  }
  
  function launchBrowser(browser, profile, tempDataDir, debugPort) {
    const browserArgs = [
      `--user-data-dir=${tempDataDir}`,
      `--profile-directory=${profile.dirName}`,
      `--remote-debugging-port=${debugPort}`,
      '--no-first-run',
      '--no-default-browser-check',
      '--disable-extensions',
      '--disable-sync',
      '--new-window',
      'about:blank',
    ];
  
    const browserProcess = spawn(browser.binary, browserArgs, {
      stdio: 'ignore',
      detached: !IS_WINDOWS,
      windowsHide: true,
    });
  
    if (!IS_WINDOWS) {
      browserProcess.unref();
    }
  
    return browserProcess;
  }
  
  function stopProcess(pid) {
    if (!pid) {
      return;
    }
  
    try {
      if (IS_WINDOWS) {
        execSync(`taskkill /PID ${pid} /F /T`, { stdio: 'pipe' });
        return;
      }
  
      process.kill(-pid, 'SIGTERM');
    } catch {
      try {
        process.kill(pid, 'SIGTERM');
      } catch {
        // Ignore cleanup failures.
      }
    }
  }
  
  function isProcessRunning(pid) {
    if (!pid) {
      return false;
    }
  
    try {
      if (IS_WINDOWS) {
        const output = execSync(`tasklist /FI "PID eq ${pid}" /NH`, {
          stdio: 'pipe',
        }).toString();
        return output.trim() !== '' && !output.includes('No tasks are running');
      }
  
      process.kill(pid, 0);
      return true;
    } catch {
      return false;
    }
  }
  
  async function waitForProcessExit(
    pid,
    timeoutMs = DEFAULT_BROWSER_SHUTDOWN_TIMEOUT_MS,
  ) {
    const deadline = Date.now() + timeoutMs;
  
    while (Date.now() < deadline) {
      if (!isProcessRunning(pid)) {
        return true;
      }
  
      await sleep(250);
    }
  
    return !isProcessRunning(pid);
  }
  
  async function closeBrowserViaCdp(debugPort) {
    const browserWebSocketUrl = await getBrowserDebuggerWebSocketUrl(debugPort);
    if (!browserWebSocketUrl) {
      return false;
    }
  
    let browserCdp = null;
  
    try {
      browserCdp = await connectCdp(browserWebSocketUrl);
      await browserCdp.send('Browser.close', {}, 5_000);
    } catch {
      return false;
    } finally {
      if (browserCdp) {
        browserCdp.close();
      }
    }
  
    return waitForBrowserDebugEndpointToClose(debugPort);
  }
  
  async function shutdownBrowser(browserProcess, debugPort) {
    let closed = false;
  
    if (debugPort != null) {
      closed = await closeBrowserViaCdp(debugPort);
    }
  
    if (!closed && browserProcess?.pid) {
      stopProcess(browserProcess.pid);
    }
  
    if (browserProcess?.pid) {
      await waitForProcessExit(browserProcess.pid);
    }
  
    if (!closed && debugPort != null) {
      await waitForBrowserDebugEndpointToClose(debugPort);
    }
  }
  
  function isLikelyMeetingPage(page) {
    const haystack = [
      page?.title || '',
      page?.url || '',
    ]
      .join(' ')
      .toLowerCase();
  
    return /stream|sharepoint|recordings|meeting|transcript|stream\.aspx/.test(
      haystack,
    );
  }
  
  function normalizeText(value) {
    return String(value || '')
      .replace(/\u200b/g, '')
      .replace(/\r\n/g, '\n')
      .replace(/[ \t]+/g, ' ')
      .replace(/\n{3,}/g, '\n\n')
      .trim();
  }
  
  function normalizeInlineText(value) {
    return normalizeText(value).replace(/\s+/g, ' ').trim();
  }
  
  function stripUtf8Bom(value) {
    return String(value || '').replace(/^\uFEFF/, '');
  }
  
  function trimForPreview(value, maxLength = MAX_CANDIDATE_PREVIEW_LENGTH) {
    const text = String(value || '');
    if (text.length <= maxLength) {
      return text;
    }
  
    return `${text.slice(0, maxLength - 3)}...`;
  }
  
  function sanitizeHeaders(headers) {
    return Object.fromEntries(
      Object.entries(headers || {}).map(([key, value]) => [
        key,
        Array.isArray(value) ? value.join(', ') : String(value),
      ]),
    );
  }
  
  function lowerCaseHeaderMap(headers) {
    return Object.fromEntries(
      Object.entries(sanitizeHeaders(headers)).map(([key, value]) => [
        key.toLowerCase(),
        value,
      ]),
    );
  }
  
  function isTextLikeContent(value) {
    return /json|text|xml|javascript|plain|html|vtt|caption|subtitle/.test(
      String(value || '').toLowerCase(),
    );
  }
  
  function isStreamRelatedUrl(url) {
    return /stream|microsoftstream|office\.com|office365\.com|office\.net|sharepoint\.com|teams\.microsoft\.com/.test(
      String(url || '').toLowerCase(),
    );
  }

  function isDataUrl(url) {
    return String(url || '').trim().toLowerCase().startsWith('data:');
  }

  function isStaticAssetResourceType(resourceType) {
    return [
      'Font',
      'Image',
      'Media',
      'Stylesheet',
      'Manifest',
    ].includes(String(resourceType || ''));
  }

  function buildNetworkSignalHaystack({
    url,
    mimeType = '',
    contentType = '',
    contentDisposition = '',
    resourceType = '',
  }) {
    const rawUrl = String(url || '');
    let normalizedUrl = rawUrl.toLowerCase();

    if (isDataUrl(rawUrl)) {
      normalizedUrl = 'data:';
    } else {
      try {
        const parsed = new URL(rawUrl);
        normalizedUrl =
          `${parsed.hostname}${parsed.pathname}${parsed.search}`.toLowerCase();
      } catch {
        normalizedUrl = rawUrl.toLowerCase();
      }
    }

    return [
      normalizedUrl,
      String(mimeType || '').toLowerCase(),
      String(contentType || '').toLowerCase(),
      String(contentDisposition || '').toLowerCase(),
      String(resourceType || '').toLowerCase(),
    ].join(' ');
  }
  
  function containsTranscriptSignal(value) {
    return /transcript|caption|subtitle|subtitles|utterance|speaker|speech|vtt|closedcaption/.test(
      String(value || '').toLowerCase(),
    );
  }
  
  function shouldTrackDebugTraffic({
    url,
    resourceType,
    method = '',
    headers = {},
    postData = '',
    captureAll = false,
  }) {
    if (captureAll) {
      return true;
    }
  
    const lowerHeaders = lowerCaseHeaderMap(headers);
    const haystack = [
      url,
      resourceType,
      method,
      lowerHeaders['content-type'] || '',
      lowerHeaders.accept || '',
      trimForPreview(postData, 400),
    ]
      .join(' ')
      .toLowerCase();
  
    if (containsTranscriptSignal(haystack)) {
      return true;
    }
  
    if (!isStreamRelatedUrl(url)) {
      return false;
    }
  
    return [
      'Document',
      'Fetch',
      'XHR',
      'WebSocket',
      'EventSource',
      'Other',
      'Preflight',
    ].includes(String(resourceType || ''));
  }
  
  function shouldFetchDebugResponseBody(responseRecord) {
    if (!responseRecord || !responseRecord.finished || responseRecord.loadingFailure) {
      return false;
    }
  
    const haystack = [
      responseRecord.url,
      responseRecord.mimeType,
      responseRecord.contentType,
    ].join(' ');
  
    if (containsTranscriptSignal(haystack)) {
      return true;
    }
  
    if (!['Document', 'Fetch', 'XHR', 'Other'].includes(responseRecord.resourceType)) {
      return false;
    }
  
    return isTextLikeContent(haystack);
  }
  
  function summarizeWebSocketPayload(payloadData) {
    const payloadText = String(payloadData || '');
    const jsonValue = tryParseJson(payloadText);
  
    return {
      payloadLength: payloadText.length,
      payloadPreview: trimForPreview(
        payloadText,
        MAX_DEBUG_FRAME_PREVIEW_LENGTH,
      ),
      jsonPreview:
        jsonValue == null
          ? ''
          : trimForPreview(
              JSON.stringify(jsonValue),
              MAX_DEBUG_FRAME_PREVIEW_LENGTH,
            ),
      transcriptSignalDetected: containsTranscriptSignal(payloadText),
    };
  }
  
  function formatTimestampValue(value) {
    if (typeof value === 'number' && Number.isFinite(value)) {
      if (value >= 86_400) {
        return String(value);
      }
  
      const wholeSeconds = Math.max(0, Math.floor(value));
      const hours = Math.floor(wholeSeconds / 3600);
      const minutes = Math.floor((wholeSeconds % 3600) / 60);
      const seconds = wholeSeconds % 60;
  
      if (hours > 0) {
        return `${hours}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
      }
  
      return `${minutes}:${String(seconds).padStart(2, '0')}`;
    }
  
    const text = normalizeInlineText(value);
    if (!text) {
      return '';
    }
  
    return text;
  }
  
  function stripHtmlTags(value) {
    return String(value || '')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/gi, ' ')
      .replace(/&amp;/gi, '&')
      .replace(/&lt;/gi, '<')
      .replace(/&gt;/gi, '>');
  }
  
  function shouldTrackNetworkResponse(response, resourceType) {
    const url = String(response?.url || '').toLowerCase();
    const mimeType = String(response?.mimeType || '').toLowerCase();
    const headers = lowerCaseHeaderMap(response?.headers);
    const contentType = String(headers['content-type'] || '').toLowerCase();
    const contentDisposition = String(headers['content-disposition'] || '').toLowerCase();

    if (isDataUrl(url) || isStaticAssetResourceType(resourceType)) {
      return false;
    }

    const score = scoreNetworkResponse({
      url,
      mimeType,
      contentType,
      contentDisposition,
      resourceType,
    });
  
    if (score >= 3) {
      return true;
    }
  
    if (!['Fetch', 'XHR'].includes(resourceType)) {
      return false;
    }
  
    return /json|text|vtt|plain|javascript/.test(`${mimeType} ${contentType}`);
  }
  
  function scoreNetworkResponse({
    url,
    mimeType,
    contentType,
    contentDisposition,
    resourceType,
  }) {
    let score = 0;
    const haystack = buildNetworkSignalHaystack({
      url,
      mimeType,
      contentType,
      contentDisposition,
      resourceType,
    });

    if (isDataUrl(url) || isStaticAssetResourceType(resourceType)) {
      return 0;
    }
  
    if (['Fetch', 'XHR'].includes(resourceType)) {
      score += 2;
    }
  
    if (/transcript|caption|subtitle|subtitles|vtt|utterance|speaker|speech/.test(haystack)) {
      score += 4;
    }
  
    if (/json|text|vtt|plain|javascript/.test(haystack)) {
      score += 2;
    }
  
    if (/stream|microsoftstream|api/.test(url)) {
      score += 1;
    }
  
    return score;
  }

  function isLikelyTranscriptResponseSignal({
    url,
    mimeType,
    contentType,
    resourceType,
    score,
  }) {
    if (isDataUrl(url) || isStaticAssetResourceType(resourceType)) {
      return false;
    }

    const haystack = buildNetworkSignalHaystack({
      url,
      mimeType,
      contentType,
      resourceType,
    });

    return (
      isEncryptedTranscriptUrl(url) ||
      containsTranscriptSignal(haystack) ||
      score >= LIVE_TRANSCRIPT_SIGNAL_SCORE
    );
  }

  function compareTrackedResponsePriority(left, right) {
    const leftScore = Number(left?.score || 0);
    const rightScore = Number(right?.score || 0);
    if (leftScore !== rightScore) {
      return leftScore - rightScore;
    }

    const leftFetchLike = ['Fetch', 'XHR'].includes(String(left?.resourceType || ''))
      ? 1
      : 0;
    const rightFetchLike = ['Fetch', 'XHR'].includes(String(right?.resourceType || ''))
      ? 1
      : 0;
    if (leftFetchLike !== rightFetchLike) {
      return leftFetchLike - rightFetchLike;
    }

    return String(left?.url || '').length - String(right?.url || '').length;
  }

  function retainTrackedResponse(trackedResponses, responseRecord, maxTrackedResponses) {
    if (trackedResponses.has(responseRecord.requestId)) {
      trackedResponses.set(responseRecord.requestId, responseRecord);
      return true;
    }

    if (trackedResponses.size < maxTrackedResponses) {
      trackedResponses.set(responseRecord.requestId, responseRecord);
      return true;
    }

    let lowestPriorityEntry = null;

    for (const [requestId, candidate] of trackedResponses.entries()) {
      if (
        !lowestPriorityEntry ||
        compareTrackedResponsePriority(candidate, lowestPriorityEntry.record) < 0
      ) {
        lowestPriorityEntry = {
          requestId,
          record: candidate,
        };
      }
    }

    if (
      !lowestPriorityEntry ||
      compareTrackedResponsePriority(responseRecord, lowestPriorityEntry.record) <= 0
    ) {
      return false;
    }

    trackedResponses.delete(lowestPriorityEntry.requestId);
    trackedResponses.set(responseRecord.requestId, responseRecord);
    return true;
  }

  function formatNetworkResponseForTerminal({
    url,
    resourceType,
    mimeType,
    contentType,
  }) {
    let displayUrl = String(url || '');

    try {
      const parsed = new URL(displayUrl);
      displayUrl = `${parsed.pathname}${parsed.search}`;
    } catch {
      // Fall back to the raw URL when parsing fails.
    }

    const typeLabel = [resourceType, mimeType || contentType]
      .filter(Boolean)
      .join(', ');
    const preview = trimForPreview(displayUrl, 140);
    return typeLabel ? `${preview} [${typeLabel}]` : preview;
  }
  
  function tryParseJson(value) {
    try {
      return JSON.parse(value);
    } catch {
      return null;
    }
  }
  
  function extractLeadingJsonText(value) {
    const text = stripUtf8Bom(String(value || '')).trimStart();
    if (!text || !['{', '['].includes(text[0])) {
      return '';
    }
  
    const stack = [text[0]];
    let inString = false;
    let isEscaped = false;
  
    for (let index = 1; index < text.length; index += 1) {
      const char = text[index];
  
      if (inString) {
        if (isEscaped) {
          isEscaped = false;
          continue;
        }
  
        if (char === '\\') {
          isEscaped = true;
          continue;
        }
  
        if (char === '"') {
          inString = false;
        }
        continue;
      }
  
      if (char === '"') {
        inString = true;
        continue;
      }
  
      if (char === '{' || char === '[') {
        stack.push(char);
        continue;
      }
  
      if (char === '}' || char === ']') {
        const expected = char === '}' ? '{' : '[';
        if (stack[stack.length - 1] !== expected) {
          return '';
        }
  
        stack.pop();
        if (stack.length === 0) {
          return text.slice(0, index + 1);
        }
      }
    }
  
    return '';
  }
  
  function tryParseJsonLenient(value) {
    const direct = tryParseJson(value);
    if (direct != null) {
      return direct;
    }
  
    const leadingJson = extractLeadingJsonText(value);
    if (leadingJson) {
      const extracted = tryParseJson(leadingJson);
      if (extracted != null) {
        return extracted;
      }
    }
  
    const text = String(value || '');
    if (!text.includes('\\\\')) {
      return null;
    }
  
    return tryParseJson(text.replace(/\\\\/g, '\\'));
  }
  
  function extractDisplayDate(value) {
    const text = normalizeText(value);
    if (!text) {
      return '';
    }
  
    const monthDateMatch = text.match(
      /\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}\b/i,
    );
    if (monthDateMatch) {
      return normalizeInlineText(monthDateMatch[0]);
    }
  
    const isoDateMatch = text.match(/\b\d{4}-\d{2}-\d{2}\b/);
    if (isoDateMatch) {
      return normalizeInlineText(isoDateMatch[0]);
    }
  
    return '';
  }
  
  function formatIsoDateForDisplay(value) {
    const text = normalizeInlineText(value);
    if (!text) {
      return '';
    }
  
    const date = new Date(text);
    if (Number.isNaN(date.getTime())) {
      return '';
    }
  
    return new Intl.DateTimeFormat('en-US', {
      month: 'long',
      day: 'numeric',
      year: 'numeric',
      timeZone: 'UTC',
    }).format(date);
  }
  
  function formatIsoDateTimeForDisplay(value) {
    const text = normalizeInlineText(value);
    if (!text) {
      return '';
    }
  
    const date = new Date(text);
    if (Number.isNaN(date.getTime())) {
      return text;
    }
  
    return new Intl.DateTimeFormat('en-US', {
      month: 'long',
      day: 'numeric',
      year: 'numeric',
      hour: 'numeric',
      minute: '2-digit',
      second: '2-digit',
      hour12: true,
      timeZone: 'UTC',
      timeZoneName: 'short',
    }).format(date);
  }
  
  function parseSharePointLookupValue(value) {
    const fields = {};
  
    for (const line of String(value || '').split(/\r?\n/)) {
      const separatorIndex = line.indexOf('|');
      const typeIndex = line.indexOf(':');
      if (separatorIndex <= 0 || typeIndex <= 0 || typeIndex > separatorIndex) {
        continue;
      }
  
      const key = line.slice(0, typeIndex).trim();
      const fieldValue = line.slice(separatorIndex + 1);
      if (!key || !fieldValue) {
        continue;
      }
  
      fields[key] = fieldValue;
    }
  
    return fields;
  }
  
  function extractUsersFromActivity(value) {
    const activity = tryParseJsonLenient(value);
    const users = activity?.FileActivityUsersOnPage;
    if (!Array.isArray(users)) {
      return [];
    }
  
    return users
      .map((user) => ({
        displayName: normalizeInlineText(user?.DisplayName || user?.displayName || ''),
        id: normalizeInlineText(user?.Id || user?.id || ''),
      }))
      .filter((user) => user.displayName || user.id);
  }
  
  function findDisplayNameForEmail(users, email) {
    const normalizedEmail = normalizeInlineText(email).toLowerCase();
    if (!normalizedEmail) {
      return '';
    }
  
    const matchedUser = users.find(
      (user) => normalizeInlineText(user.id).toLowerCase() === normalizedEmail,
    );
  
    return matchedUser?.displayName || '';
  }
  
  function extractStreamMeetingMetadataFromBody(bodyText) {
    const jsonValue = tryParseJsonLenient(bodyText);
    const rows = jsonValue?.ListData?.Row;
    if (!Array.isArray(rows) || rows.length === 0) {
      return null;
    }
  
    for (const row of rows) {
      const metaInfoValues = Array.isArray(row?.MetaInfo)
        ? row.MetaInfo
            .map((item) =>
              typeof item === 'string' ? item : item?.lookupValue || '',
            )
            .filter(Boolean)
        : typeof row?.MetaInfo === 'string'
          ? [row.MetaInfo]
          : [];
  
      const lookupValue = metaInfoValues.find(
        (value) =>
          String(value).includes('vti_stream_tmr_organizerupn:') ||
          String(value).includes('vti_stream_mediaitemmetadata:') ||
          String(value).includes('vti_title:'),
      );
      if (!lookupValue) {
        continue;
      }
  
      const fields = parseSharePointLookupValue(lookupValue);
      const mediaItemMetadata = tryParseJsonLenient(
        fields.vti_stream_mediaitemmetadata || '',
      );
      const organizerEmail = normalizeInlineText(fields.vti_stream_tmr_organizerupn || '');
      const activityUsers = extractUsersFromActivity(row?._activity || row?.Activity || '');
  
      return {
        title: normalizeInlineText(fields.vti_title || row?.FileLeafRef || ''),
        date: formatIsoDateForDisplay(mediaItemMetadata?.recordingStartDateTime || ''),
        createdBy: findDisplayNameForEmail(activityUsers, organizerEmail),
        createdByEmail: organizerEmail,
        createdByTenantId: normalizeInlineText(
          fields.vti_stream_tmr_organizertenantid || '',
        ),
        recordedBy: '',
        sourceUrl: '',
        sourceApplication: normalizeInlineText(fields.vti_stream_sourceapplication || ''),
        recordingStartDateTime: normalizeInlineText(
          mediaItemMetadata?.recordingStartDateTime || '',
        ),
        recordingEndDateTime: normalizeInlineText(
          mediaItemMetadata?.recordingEndDateTime || '',
        ),
        sharePointFilePath: normalizeInlineText(row?.FileRef || ''),
        sharePointItemUrl: normalizeInlineText(row?.['.spItemUrl'] || ''),
      };
    }
  
    return null;
  }
  
  function extractMeetingMetadataFromCapture(captureResult) {
    const bodySources = [];
  
    for (const candidate of captureResult?.candidates || []) {
      bodySources.push({
        url: candidate.url,
        body: candidate.body,
      });
    }
  
    for (const responseBody of captureResult?.debug?.responseBodies || []) {
      bodySources.push({
        url: responseBody.url,
        body: responseBody.body,
      });
    }
  
    for (const bodySource of bodySources) {
      if (!bodySource?.body || !isStreamRelatedUrl(bodySource.url)) {
        continue;
      }
  
      const metadata = extractStreamMeetingMetadataFromBody(bodySource.body);
      if (metadata) {
        return metadata;
      }
    }
  
    return {
      title: '',
      date: '',
      createdBy: '',
      createdByEmail: '',
      createdByTenantId: '',
      recordedBy: '',
      sourceUrl: '',
      sourceApplication: '',
      recordingStartDateTime: '',
      recordingEndDateTime: '',
      sharePointFilePath: '',
      sharePointItemUrl: '',
    };
  }
  
  function mergeMeetingMetadata(pageMetadata, captureMetadata, targetPage) {
    return {
      title:
        pageMetadata.title ||
        captureMetadata.title ||
        targetPage?.title ||
        '',
      date:
        extractDisplayDate(pageMetadata.date) ||
        extractDisplayDate(captureMetadata.date) ||
        captureMetadata.date ||
        formatIsoDateForDisplay(captureMetadata.recordingStartDateTime) ||
        '',
      recordedBy:
        pageMetadata.recordedBy ||
        captureMetadata.recordedBy ||
        '',
      createdBy:
        captureMetadata.createdBy ||
        pageMetadata.createdBy ||
        '',
      createdByEmail:
        captureMetadata.createdByEmail ||
        pageMetadata.createdByEmail ||
        '',
      createdByTenantId:
        captureMetadata.createdByTenantId ||
        pageMetadata.createdByTenantId ||
        '',
      sourceUrl:
        pageMetadata.sourceUrl ||
        targetPage?.url ||
        captureMetadata.sourceUrl ||
        '',
      sourceApplication:
        captureMetadata.sourceApplication ||
        pageMetadata.sourceApplication ||
        '',
      recordingStartDateTime:
        captureMetadata.recordingStartDateTime ||
        pageMetadata.recordingStartDateTime ||
        '',
      recordingEndDateTime:
        captureMetadata.recordingEndDateTime ||
        pageMetadata.recordingEndDateTime ||
        '',
      sharePointFilePath:
        captureMetadata.sharePointFilePath ||
        pageMetadata.sharePointFilePath ||
        '',
      sharePointItemUrl:
        captureMetadata.sharePointItemUrl ||
        pageMetadata.sharePointItemUrl ||
        '',
    };
  }
  
  function vttTimestampToDisplay(value) {
    const match = String(value || '').trim().match(
      /^(?:(\d{2,}):)?(\d{2}):(\d{2})(?:\.(\d{1,3}))?$/,
    );
  
    if (!match) {
      return normalizeInlineText(value);
    }
  
    const hours = Number(match[1] || 0);
    const minutes = Number(match[2] || 0);
    const seconds = Number(match[3] || 0);
  
    if (hours > 0) {
      return `${hours}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
    }
  
    return `${minutes}:${String(seconds).padStart(2, '0')}`;
  }
  
  /**
   * @param {string} bodyText
   * @returns {TranscriptEntry[]}
   */
  function parseWebVttTranscript(bodyText) {
    const trimmed = String(bodyText || '').trim();
    if (!trimmed.startsWith('WEBVTT')) {
      return [];
    }
  
    const blocks = trimmed.split(/\n{2,}/);
    const entries = [];
  
    for (const rawBlock of blocks) {
      const lines = rawBlock
        .split(/\r?\n/)
        .map((line) => line.trim())
        .filter(Boolean);
  
      if (lines.length === 0 || lines[0] === 'WEBVTT' || lines[0].startsWith('NOTE')) {
        continue;
      }
  
      let cursor = 0;
      if (!lines[cursor].includes('-->') && lines[cursor + 1]?.includes('-->')) {
        cursor += 1;
      }
  
      const timingLine = lines[cursor];
      if (!timingLine || !timingLine.includes('-->')) {
        continue;
      }
  
      const [startTime] = timingLine.split(/\s+-->\s+/, 2);
      const textLines = lines.slice(cursor + 1);
      if (textLines.length === 0) {
        continue;
      }
  
      let speaker = '';
      const normalizedLines = textLines.map((line) => {
        const speakerTagMatch = line.match(/<v(?:\.[^>]*)?\s+([^>]+)>(.*)$/i);
        if (speakerTagMatch) {
          speaker = normalizeInlineText(stripHtmlTags(speakerTagMatch[1]));
          return speakerTagMatch[2];
        }
  
        return line;
      });
  
      let text = normalizeInlineText(stripHtmlTags(normalizedLines.join(' ')));
      if (!speaker) {
        const speakerPrefixMatch = text.match(/^([^:]{2,80}):\s+(.*)$/);
        if (speakerPrefixMatch) {
          speaker = normalizeInlineText(speakerPrefixMatch[1]);
          text = normalizeInlineText(speakerPrefixMatch[2]);
        }
      }
  
      if (!text) {
        continue;
      }
  
      entries.push({
        speaker,
        timestamp: vttTimestampToDisplay(startTime),
        text,
      });
    }
  
    return entries;
  }
  
  function getFirstStringValue(target, keyPaths) {
    for (const path of keyPaths) {
      const parts = path.split('.');
      let current = target;
  
      for (const part of parts) {
        if (current == null || typeof current !== 'object' || !(part in current)) {
          current = undefined;
          break;
        }
  
        current = current[part];
      }
  
      if (typeof current === 'string' && normalizeInlineText(current)) {
        return current;
      }
    }
  
    return '';
  }
  
  function getFirstNumberValue(target, keyPaths) {
    for (const path of keyPaths) {
      const parts = path.split('.');
      let current = target;
  
      for (const part of parts) {
        if (current == null || typeof current !== 'object' || !(part in current)) {
          current = undefined;
          break;
        }
  
        current = current[part];
      }
  
      if (typeof current === 'number' && Number.isFinite(current)) {
        return current;
      }
    }
  
    return null;
  }
  
  function extractEntryFromObject(value) {
    if (!value || typeof value !== 'object' || Array.isArray(value)) {
      return null;
    }
  
    const speaker = getFirstStringValue(value, [
      'speaker',
      'speakerName',
      'speakerDisplayName',
      'displayName',
      'participantName',
      'userDisplayName',
      'speaker.displayName',
      'speaker.name',
      'user.displayName',
      'identity.displayName',
      'from.user.displayName',
    ]);
  
    const text = getFirstStringValue(value, [
      'text',
      'displayText',
      'content',
      'caption',
      'utterance',
      'transcript',
      'message',
      'body',
      'cueText',
    ]);
  
    if (!normalizeInlineText(text)) {
      return null;
    }
  
    const numericTimestamp = getFirstNumberValue(value, [
      'start',
      'startTime',
      'offset',
      'begin',
      'startOffset',
    ]);
    const stringTimestamp = getFirstStringValue(value, [
      'timestamp',
      'startDateTime',
      'createdDateTime',
      'time',
      'begin',
      'startTime',
      'startOffset',
      'endOffset',
    ]);
  
    return {
      speaker: normalizeInlineText(speaker),
      timestamp: formatTimestampValue(
        numericTimestamp != null ? numericTimestamp : stringTimestamp,
      ),
      text: normalizeInlineText(stripHtmlTags(text)),
    };
  }
  
  function extractEntriesFromJson(value) {
    if (
      value &&
      typeof value === 'object' &&
      !Array.isArray(value) &&
      Array.isArray(value.entries)
    ) {
      const topLevelEntries = value.entries
        .map((item) => extractEntryFromObject(item))
        .filter((entry) => entry && entry.text);
  
      if (topLevelEntries.length > 0) {
        return {
          path: '$.entries',
          entries: topLevelEntries,
          score:
            topLevelEntries.length * 4 +
            topLevelEntries.filter((entry) => entry.speaker).length * 2 +
            topLevelEntries.filter((entry) => entry.timestamp).length +
            12,
        };
      }
    }
  
    const matches = [];
  
    function walk(current, path, depth = 0) {
      if (depth > 10 || current == null) {
        return;
      }
  
      if (Array.isArray(current)) {
        const entries = current
          .map((item) => extractEntryFromObject(item))
          .filter((entry) => entry && entry.text);
  
        if (entries.length > 0) {
          const pathText = path.toLowerCase();
          let score = entries.length * 3;
          score += entries.filter((entry) => entry.speaker).length * 2;
          score += entries.filter((entry) => entry.timestamp).length;
  
          if (/transcript|caption|subtitle|utterance|cue/.test(pathText)) {
            score += 10;
          }
  
          matches.push({
            path,
            entries,
            score,
          });
        }
  
        current.forEach((item, index) => {
          walk(item, `${path}[${index}]`, depth + 1);
        });
        return;
      }
  
      if (typeof current !== 'object') {
        return;
      }
  
      for (const [key, child] of Object.entries(current)) {
        walk(child, `${path}.${key}`, depth + 1);
      }
    }
  
    walk(value, '$');
    matches.sort((left, right) => right.score - left.score);
    return matches[0] || {
      path: '',
      entries: [],
      score: 0,
    };
  }
  
  function extractProtectionKeyFromLookupValue(value, sourceUrl = '') {
    const protectionLine = String(value || '')
      .split(/\r?\n/)
      .find((line) => line.startsWith('vti_mediaserviceprotectionkey:SW|'));
    if (!protectionLine) {
      return null;
    }
  
    const jsonStart = protectionLine.indexOf('|{');
    if (jsonStart < 0) {
      return null;
    }
  
    const protectionEnvelope = tryParseJsonLenient(
      protectionLine.slice(jsonStart + 1),
    );
    if (!protectionEnvelope || typeof protectionEnvelope !== 'object') {
      return null;
    }
  
    const protectionData = tryParseJsonLenient(
      String(protectionEnvelope.ProtectionKeyData || ''),
    );
    if (!protectionData || typeof protectionData !== 'object') {
      return null;
    }
  
    const keyBase64 = String(protectionData.Key || '');
    const ivBase64 = String(protectionData.IV || '');
    let keyBytes;
    let ivBytes;
  
    try {
      keyBytes = Buffer.from(keyBase64, 'base64');
      ivBytes = Buffer.from(ivBase64, 'base64');
    } catch {
      return null;
    }
  
    if (keyBytes.length !== 16 || ivBytes.length !== 16) {
      return null;
    }
  
    const kid = String(protectionData.Kid || protectionEnvelope.keyId || '').trim();
    if (!kid) {
      return null;
    }
  
    return {
      kid,
      sourceUrl,
      encryptionAlgorithm:
        typeof protectionData.EncryptionAlgorithm === 'number'
          ? protectionData.EncryptionAlgorithm
          : null,
      encryptionMode:
        typeof protectionData.EncryptionMode === 'number'
          ? protectionData.EncryptionMode
          : null,
      keySize:
        typeof protectionData.KeySize === 'number'
          ? protectionData.KeySize
          : keyBytes.length * 8,
      padding: String(protectionData.Padding || ''),
      keyBase64,
      ivBase64,
    };
  }
  
  function collectProtectionKeysFromBody(bodyText, sourceUrl = '') {
    const jsonValue = tryParseJson(bodyText);
    if (jsonValue == null) {
      return [];
    }
  
    const matches = [];
  
    function walk(current, depth = 0) {
      if (current == null || depth > 12) {
        return;
      }
  
      if (typeof current === 'string') {
        if (current.includes('vti_mediaserviceprotectionkey:SW|')) {
          const protectionKey = extractProtectionKeyFromLookupValue(
            current,
            sourceUrl,
          );
          if (protectionKey) {
            matches.push(protectionKey);
          }
        }
        return;
      }
  
      if (Array.isArray(current)) {
        current.forEach((item) => walk(item, depth + 1));
        return;
      }
  
      if (typeof current !== 'object') {
        return;
      }
  
      Object.values(current).forEach((child) => walk(child, depth + 1));
    }
  
    walk(jsonValue);
    return matches;
  }
  
  function buildProtectionKeyMap(bodySources) {
    const keysByKid = new Map();
  
    for (const source of bodySources) {
      if (!source?.body) {
        continue;
      }
  
      const matches = collectProtectionKeysFromBody(source.body, source.url || '');
      for (const match of matches) {
        if (!keysByKid.has(match.kid)) {
          keysByKid.set(match.kid, match);
        }
      }
    }
  
    return keysByKid;
  }
  
  function getUrlSearchParam(url, paramName) {
    try {
      return new URL(url).searchParams.get(paramName) || '';
    } catch {
      return '';
    }
  }
  
  function isEncryptedTranscriptUrl(url) {
    return /\/cdnmedia\/transcripts/i.test(String(url || ''));
  }
  
  function decryptTranscriptBody(bodyRecord, protectionKey) {
    if (!bodyRecord?.rawBodyBase64) {
      return {
        body: bodyRecord?.body || '',
        decrypted: false,
        decryptionKeyId: '',
        decryptionError: '',
      };
    }
  
    const encryptionAlgorithm = Number(protectionKey?.encryptionAlgorithm);
    const encryptionMode = Number(protectionKey?.encryptionMode);
    if (encryptionAlgorithm !== 0 || encryptionMode !== 1) {
      return {
        body: bodyRecord.body || '',
        decrypted: false,
        decryptionKeyId: '',
        decryptionError:
          `Unsupported transcript encryption algorithm/mode: ` +
          `${encryptionAlgorithm}/${encryptionMode}`,
      };
    }
  
    try {
      const ciphertext = Buffer.from(bodyRecord.rawBodyBase64, 'base64');
      const decipher = createDecipheriv(
        'aes-128-cbc',
        Buffer.from(protectionKey.keyBase64, 'base64'),
        Buffer.from(protectionKey.ivBase64, 'base64'),
      );
  
      if (String(protectionKey.padding || '').toLowerCase() === 'none') {
        decipher.setAutoPadding(false);
      }
  
      const plaintext = Buffer.concat([
        decipher.update(ciphertext),
        decipher.final(),
      ]);
      const decoded = stripUtf8Bom(
        plaintext.toString('utf8').replace(/\u0000+$/g, ''),
      );
  
      return {
        body: decoded,
        decrypted: true,
        decryptionKeyId: protectionKey.kid,
        decryptionError: '',
      };
    } catch (error) {
      return {
        body: bodyRecord.body || '',
        decrypted: false,
        decryptionKeyId: '',
        decryptionError:
          error instanceof Error
            ? error.message
            : 'Transcript decryption failed.',
      };
    }
  }
  
  function finalizeBodyRecordForResponse(url, bodyRecord, protectionKeys) {
    const normalizedBody = stripUtf8Bom(bodyRecord?.body || '');
    let body = normalizedBody;
    let decrypted = false;
    let decryptionKeyId = '';
    let decryptionError = '';
  
    if (isEncryptedTranscriptUrl(url)) {
      const keyId = getUrlSearchParam(url, 'kid');
      const protectionKey =
        (keyId && protectionKeys.get(keyId)) ||
        (protectionKeys.size === 1
          ? Array.from(protectionKeys.values())[0]
          : null);
  
      if (protectionKey) {
        const decryptedBody = decryptTranscriptBody(
          {
            ...bodyRecord,
            body: normalizedBody,
          },
          protectionKey,
        );
        body = decryptedBody.body;
        decrypted = decryptedBody.decrypted;
        decryptionKeyId = decryptedBody.decryptionKeyId;
        decryptionError = decryptedBody.decryptionError;
      }
    }
  
    return {
      ...bodyRecord,
      body,
      bodyLength: body.length,
      bodyPreview: trimForPreview(body),
      decrypted,
      decryptionKeyId,
      decryptionError,
    };
  }
  
  function summarizeCapturedBody(bodyText) {
    const trimmed = stripUtf8Bom(String(bodyText || '')).trim();
    if (!trimmed) {
      return {
        format: '',
        path: '',
        entryCount: 0,
        entries: [],
      };
    }
  
    const vttEntries = parseWebVttTranscript(trimmed);
    if (vttEntries.length > 0) {
      return {
        format: 'vtt',
        path: '$',
        entryCount: vttEntries.length,
        entries: vttEntries,
      };
    }
  
    const jsonValue = tryParseJsonLenient(trimmed);
    if (jsonValue != null) {
      const jsonMatch = extractEntriesFromJson(jsonValue);
      return {
        format: 'json',
        path: jsonMatch.path,
        entryCount: jsonMatch.entries.length,
        entries: jsonMatch.entries,
      };
    }
  
    return {
      format: '',
      path: '',
      entryCount: 0,
      entries: [],
    };
  }
  
  async function extractMeetingMetadata(cdp) {
    return evaluate(
      cdp,
      `
        (() => {
          const title =
            document.querySelector('h1')?.textContent?.trim() ||
            document.title.replace(' - Microsoft Stream', '').trim();
          const allText = document.body.innerText;
          const dateMatch =
            allText.match(
              /\\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\\s+\\d{1,2},\\s+\\d{4}\\b/i,
            ) ||
            allText.match(/(\\d{4}-\\d{2}-\\d{2}[\\s\\d:]*(?:UTC|GMT)?)/);
          const recordedByMatch = allText.match(/Recorded by\\s*\\n?\\s*(.+)/);
  
          return {
            title,
            date: dateMatch?.[0]?.trim() || '',
            recordedBy: recordedByMatch?.[1]?.trim() || '',
            sourceUrl: location.href,
            createdBy: '',
            createdByEmail: '',
            createdByTenantId: '',
            sourceApplication: '',
            recordingStartDateTime: '',
            recordingEndDateTime: '',
            sharePointFilePath: '',
            sharePointItemUrl: '',
          };
        })()
      `,
    );
  }
  
  function buildOutputPayload(metadata, entries) {
    const speakers = [...new Set(entries.map((entry) => entry.speaker).filter(Boolean))];
  
    return {
      meeting: metadata,
      extractedAt: new Date().toISOString(),
      entryCount: entries.length,
      speakers,
      entries,
    };
  }
  
  function sanitizeFilename(value) {
    const sanitized = value
      .replace(/[^a-zA-Z0-9-_ ]/g, '')
      .replace(/\s+/g, '_')
      .slice(0, 80);
  
    return sanitized || 'meeting';
  }
  
  function resolveOutputDirectory(outputDir) {
    return resolve(process.cwd(), outputDir || DEFAULT_OUTPUT_DIR);
  }
  
  function buildOutputBasePath(defaultName, outputName, outputDir) {
    const absoluteOutputDir = resolveOutputDirectory(outputDir);
    mkdirSync(absoluteOutputDir, { recursive: true });
  
    const timestamp = new Date()
      .toISOString()
      .replace(/[:.]/g, '-')
      .slice(0, 19);
  
    const baseName = sanitizeFilename(outputName || defaultName || 'meeting');
    return join(absoluteOutputDir, `${baseName}_${timestamp}`);
  }
  
  function buildMarkdownOutput(payload) {
    const createdByLabel = payload.meeting.createdByEmail
      ? payload.meeting.createdBy &&
        payload.meeting.createdBy.toLowerCase() !==
          payload.meeting.createdByEmail.toLowerCase()
        ? `${payload.meeting.createdBy} <${payload.meeting.createdByEmail}>`
        : payload.meeting.createdByEmail
      : payload.meeting.createdBy || '';
    const speakerList = Array.isArray(payload.speakers)
      ? payload.speakers.filter(Boolean).join(', ')
      : '';
    const startDateTime = formatIsoDateTimeForDisplay(
      payload.meeting.recordingStartDateTime,
    );
    const endDateTime = formatIsoDateTimeForDisplay(
      payload.meeting.recordingEndDateTime,
    );
    const metadataLines = [
      payload.meeting.title ? `Title: ${payload.meeting.title}` : '',
      payload.meeting.date ? `Date: ${payload.meeting.date}` : '',
      startDateTime ? `Start date/time: ${startDateTime}` : '',
      endDateTime ? `End date/time: ${endDateTime}` : '',
      createdByLabel ? `Created by: ${createdByLabel}` : '',
      payload.meeting.recordedBy ? `Recorded by: ${payload.meeting.recordedBy}` : '',
      speakerList ? `Speakers: ${speakerList}` : '',
      payload.meeting.sourceUrl ? `Source URL: ${payload.meeting.sourceUrl}` : '',
      `Extracted at: ${payload.extractedAt}`,
      `Entry count: ${payload.entryCount}`,
    ].filter(Boolean);
  
    const entryBlocks = payload.entries.map((entry) => {
      const speaker = entry.speaker || 'Unknown speaker';
      const timestamp = entry.timestamp || '';
      const heading = `${speaker} - ${timestamp}:`;
      const text = String(entry.text || '')
        .replace(/\r\n/g, '\n')
        .replace(/\n{2,}/g, '\n')
        .trim();
  
      return `${heading}\n${text}`;
    });
  
    return [
      'Transcript information',
      '',
      ...metadataLines,
      '',
      '---',
      '',
      entryBlocks.join('\n\n'),
      '',
    ].join('\n');
  }
  
  function saveOutputs(payload, outputBasePath, outputFormat) {
    const transcriptPaths = [];
  
    if (outputFormat === 'json' || outputFormat === 'both') {
      const jsonPath = `${outputBasePath}.json`;
      writeFileSync(jsonPath, JSON.stringify(payload, null, 2));
      transcriptPaths.push(jsonPath);
    }
  
    if (outputFormat === 'md' || outputFormat === 'both') {
      const markdownPath = `${outputBasePath}.md`;
      writeFileSync(markdownPath, buildMarkdownOutput(payload));
      transcriptPaths.push(markdownPath);
    }
  
    return transcriptPaths;
  }
  
  function saveNetworkCaptureOutput(payload, outputBasePath) {
    const outputPath = `${outputBasePath}.network.json`;
    writeFileSync(outputPath, JSON.stringify(payload, null, 2));
    return outputPath;
  }
  
  async function selectTranscriptPage(prompt, debugPort) {
    const pages = await findPageTargets(debugPort);
  
    if (pages.length === 0) {
      throw new CliError(
        'No browser pages were found. Open Microsoft Stream before continuing.',
      );
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
      'Open pages',
      pages,
      (page) => `${page.title} (${page.url.slice(0, 80)})`,
      '\nWhich page contains the Microsoft Stream meeting? (number): ',
    );
  }
  
  function printInitialRunInstructions(captureControl = 'manual') {
    console.log('\nBrowser is ready.');
    console.log('1. Open the meeting in Microsoft Stream.');
    if (captureControl === 'automatic') {
      console.log(
        '2. Return to this terminal and continue. The extractor will reload the page and manage the Transcript panel automatically.',
      );
      return;
    }

    console.log('2. Leave the Transcript panel closed for now.');
    console.log('3. Return to this terminal and continue so capture can be armed.');
  }
  
  function printCaptureInstructions(debugEnabled = false) {
    console.log('\nNetwork capture is armed.');
    if (debugEnabled) {
      console.log('1. Wait for the automatic page reload to finish.');
      console.log('2. Open the Transcript panel in Microsoft Stream.');
      console.log('3. If the panel is already open, close and reopen it if possible.');
      console.log('4. Let the transcript load, and scroll once if the app lazily fetches chunks.');
      console.log(
        '5. Watch this terminal for a transcript-traffic confirmation if one is seen.',
      );
      console.log('6. Return to this terminal and press Enter.');
      return;
    }
  
    console.log('1. Open the Transcript panel in Microsoft Stream.');
    console.log('2. If the panel is already open, close and reopen it if possible.');
    console.log('3. Let the transcript load, and scroll once if the app lazily fetches chunks.');
    console.log(
      '4. Watch this terminal for a transcript-traffic confirmation if one is seen.',
    );
    console.log('5. Return to this terminal and press Enter.');
  }

  function printAutomaticFallbackInstructions() {
    console.log('\nAutomatic mode needs a manual assist.');
    console.log('1. Open or reopen the Transcript panel in Microsoft Stream.');
    console.log('2. Let the transcript load, and scroll once if the app lazily fetches chunks.');
    console.log(
      '3. Watch this terminal for a transcript-traffic confirmation if one is seen.',
    );
    console.log('4. Return to this terminal and press Enter.');
  }

  function buildTranscriptPanelAutomationExpression(action = 'inspect') {
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

  async function waitForAsyncCondition(
    check,
    timeoutMs,
    pollMs = DEFAULT_AUTOMATIC_UI_POLL_MS,
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

  function logAutomaticAction(
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

  async function inspectTranscriptPanelUi(cdp) {
    return evaluate(cdp, buildTranscriptPanelAutomationExpression('inspect'));
  }

  async function performTranscriptPanelAction(cdp, action) {
    return evaluate(cdp, buildTranscriptPanelAutomationExpression(action));
  }

  async function waitForTranscriptPanelToOpen(
    cdp,
    timeoutMs = DEFAULT_AUTOMATIC_PANEL_OPEN_TIMEOUT_MS,
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
        await sleep(DEFAULT_AUTOMATIC_CLICK_SETTLE_MS);
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

      await sleep(DEFAULT_AUTOMATIC_CLICK_SETTLE_MS);
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
    logAutomaticAction(captureFeedback, debugEnabled, 'transcript panel result', {
      opened: panelResult.opened,
      refreshed: panelResult.refreshed,
      controlLabel: panelResult.controlLabel || '',
      state: summarizeTranscriptPanelUiState(panelResult.state),
    });

    if (!panelResult.opened) {
      console.log(
        'Automatic mode: could not open the Transcript panel automatically.',
      );
      printAutomaticFallbackSummary(
        'panel-open-failed',
        captureFeedback,
        panelResult,
      );
      logAutomaticAction(
        captureFeedback,
        debugEnabled,
        'automatic open failed, falling back to manual assist',
      );
      printAutomaticFallbackInstructions();
      await prompt.waitForEnter(
        '\nPress Enter after the transcript panel has loaded...\n',
      );
      await sleep(DEFAULT_CAPTURE_SETTLE_MS);
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
        timeoutMs: DEFAULT_AUTOMATIC_SIGNAL_TIMEOUT_MS,
      },
    );
    let transcriptSignal = await waitForTranscriptSignal(
      captureFeedback,
      DEFAULT_AUTOMATIC_SIGNAL_TIMEOUT_MS,
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
        DEFAULT_AUTOMATIC_SIGNAL_RETRY_TIMEOUT_MS,
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
          DEFAULT_AUTOMATIC_SIGNAL_RETRY_TIMEOUT_MS,
        );
      }
    }

    if (!transcriptSignal) {
      console.log(
        'Automatic mode: transcript traffic was not detected automatically.',
      );
      printAutomaticFallbackSummary(
        'traffic-not-detected',
        captureFeedback,
        panelResult,
      );
      logAutomaticAction(
        captureFeedback,
        debugEnabled,
        'automatic signal wait failed, falling back to manual assist',
      );
      printAutomaticFallbackInstructions();
      await prompt.waitForEnter(
        '\nPress Enter after the transcript panel has loaded...\n',
      );
      await sleep(DEFAULT_CAPTURE_SETTLE_MS);
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
      DEFAULT_AUTOMATIC_REQUEST_SETTLE_TIMEOUT_MS,
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
          timeoutMs: DEFAULT_AUTOMATIC_REQUEST_SETTLE_TIMEOUT_MS,
        },
      );
    }

    await sleep(DEFAULT_CAPTURE_SETTLE_MS);
  }
  
  async function loadResponseBody(cdp, requestId, bodyCache) {
    if (bodyCache.has(requestId)) {
      return bodyCache.get(requestId);
    }
  
    let body = '';
    let bodyError = '';
    let base64Encoded = false;
    let rawBodyBase64 = '';
    let rawBodyByteLength = 0;
    let rawBodyPreviewHex = '';
  
    try {
      const result = await cdp.send('Network.getResponseBody', { requestId });
      base64Encoded = Boolean(result.base64Encoded);
      rawBodyBase64 = base64Encoded ? result.body : '';
  
      const rawBytes = base64Encoded
        ? Buffer.from(result.body, 'base64')
        : Buffer.from(result.body, 'utf8');
      rawBodyByteLength = rawBytes.length;
      rawBodyPreviewHex = rawBytes.subarray(0, 32).toString('hex');
      body = base64Encoded ? rawBytes.toString('utf8') : result.body;
    } catch (error) {
      bodyError =
        error instanceof Error ? error.message : 'Unable to capture response body.';
    }
  
    const record = {
      body,
      bodyLength: body.length,
      bodyPreview: trimForPreview(body),
      bodyError,
      base64Encoded,
      rawBodyBase64,
      rawBodyByteLength,
      rawBodyPreviewHex,
      refetchedExternally: false,
      refetchError: '',
    };
    bodyCache.set(requestId, record);
    return record;
  }
  
  async function refetchBinaryBody(url) {
    const response = await fetch(url, {
      redirect: 'follow',
    });
  
    if (!response.ok) {
      throw new Error(`Refetch failed with status ${response.status}.`);
    }
  
    const rawBytes = Buffer.from(await response.arrayBuffer());
    return {
      body: rawBytes.toString('utf8'),
      bodyLength: rawBytes.length,
      bodyPreview: trimForPreview(rawBytes.toString('utf8')),
      bodyError: '',
      base64Encoded: true,
      rawBodyBase64: rawBytes.toString('base64'),
      rawBodyByteLength: rawBytes.length,
      rawBodyPreviewHex: rawBytes.subarray(0, 32).toString('hex'),
      refetchedExternally: true,
      refetchError: '',
    };
  }
  
  async function reloadPageAndWait(cdp, timeoutMs = DEFAULT_CDP_TIMEOUT_MS) {
    await cdp.send('Page.enable');
  
    await new Promise((resolveReload, rejectReload) => {
      let settled = false;
      const timeout = setTimeout(() => {
        if (settled) {
          return;
        }
  
        settled = true;
        stopListening();
        resolveReload();
      }, timeoutMs);
  
      const stopListening = cdp.on('Page.loadEventFired', () => {
        if (settled) {
          return;
        }
  
        settled = true;
        clearTimeout(timeout);
        stopListening();
        resolveReload();
      });
  
      cdp.send('Page.reload', { ignoreCache: true }).catch((error) => {
        if (settled) {
          return;
        }
  
        settled = true;
        clearTimeout(timeout);
        stopListening();
        rejectReload(error);
      });
    });
  
    await sleep(1_000);
  }
  
  async function captureStreamNetwork(
    cdp,
    prompt,
    debugEnabled = false,
    captureControl = 'manual',
  ) {
    await cdp.send('Network.enable', {
      maxTotalBufferSize: 100_000_000,
      maxResourceBufferSize: 10_000_000,
    });
  
    const trackedResponses = new Map();
    const finishedRequests = new Set();
    const failedRequests = new Map();
    const bodyCache = new Map();
    const captureFeedback = {
      candidateResponseCount: 0,
      transcriptHintCount: 0,
      transcriptHintSettledCount: 0,
      transcriptHintFailedCount: 0,
      firstTranscriptHint: null,
      printedTranscriptHint: false,
      automaticActions: [],
    };
    const transcriptHintRequestIds = new Set();
    const debugState = debugEnabled
      ? {
          requests: new Map(),
          responses: new Map(),
          failures: [],
          webSockets: new Map(),
          webSocketFrames: [],
          eventSourceMessages: [],
        }
      : null;
  
    function ensureTrackedWebSocket(requestId, url = '') {
      if (!debugState) {
        return null;
      }
  
      const normalizedRequestId = String(requestId || '');
      if (!normalizedRequestId) {
        return null;
      }
  
      if (debugState.webSockets.has(normalizedRequestId)) {
        const socket = debugState.webSockets.get(normalizedRequestId);
        if (url && !socket.url) {
          socket.url = url;
        }
        return socket;
      }
  
      if (
        debugState.webSockets.size >= MAX_DEBUG_WEBSOCKETS ||
        !shouldTrackDebugTraffic({
          url,
          resourceType: 'WebSocket',
          captureAll: true,
        })
      ) {
        return null;
      }
  
      const socket = {
        requestId: normalizedRequestId,
        url,
        created: false,
        createdAt: null,
        handshakeRequest: null,
        handshakeResponse: null,
        framesSent: 0,
        framesReceived: 0,
        transcriptSignalDetected: false,
        closedAt: null,
      };
      debugState.webSockets.set(normalizedRequestId, socket);
      return socket;
    }
  
    const stopResponseListener = cdp.on('Network.responseReceived', (params) => {
      if (shouldTrackNetworkResponse(params.response, params.type)) {
        const headers = sanitizeHeaders(params.response.headers);
        const lowerHeaders = lowerCaseHeaderMap(params.response.headers);
        const requestId = String(params.requestId);
        const url = String(params.response.url || '');
        const mimeType = String(params.response.mimeType || '');
        const resourceType = String(params.type || '');
        const contentType = String(lowerHeaders['content-type'] || '');
        const score = scoreNetworkResponse({
          url: url.toLowerCase(),
          mimeType: mimeType.toLowerCase(),
          contentType: contentType.toLowerCase(),
          contentDisposition: String(lowerHeaders['content-disposition'] || '').toLowerCase(),
          resourceType,
        });
  
        retainTrackedResponse(trackedResponses, {
          requestId,
          url,
          status: params.response.status,
          mimeType,
          resourceType,
          headers,
          score,
        }, MAX_TRACKED_RESPONSES);
        captureFeedback.candidateResponseCount = trackedResponses.size;

        if (
          isLikelyTranscriptResponseSignal({
            url,
            mimeType,
            contentType,
            resourceType,
            score,
          })
        ) {
          transcriptHintRequestIds.add(requestId);
          captureFeedback.transcriptHintCount += 1;
          if (captureControl === 'automatic') {
            logAutomaticAction(
              captureFeedback,
              debugEnabled,
              'transcript hint response received',
              {
                requestId,
                url,
                resourceType,
                mimeType,
                contentType,
                score,
              },
            );
          }

          if (!captureFeedback.firstTranscriptHint) {
            captureFeedback.firstTranscriptHint = {
              requestId,
              url,
              resourceType,
              mimeType,
              contentType,
              score,
            };
          }

          if (!captureFeedback.printedTranscriptHint) {
            captureFeedback.printedTranscriptHint = true;
            console.log(
              `\nDetected likely transcript network response: ` +
                `${formatNetworkResponseForTerminal({
                  url,
                  resourceType,
                  mimeType,
                  contentType,
                })}`,
            );
            console.log(
              captureControl === 'automatic'
                ? 'Automatic mode will keep waiting for the response to settle.'
                : 'Let the transcript finish loading, then press Enter to continue.',
            );
          }
        }
      }
  
      if (!debugState) {
        return;
      }
  
      const requestId = String(params.requestId);
      const url = String(params.response.url || '');
      const resourceType = String(params.type || '');
      if (
        !debugState.responses.has(requestId) &&
        debugState.responses.size >= MAX_DEBUG_RESPONSES
      ) {
        return;
      }
  
      if (!shouldTrackDebugTraffic({
        url,
        resourceType,
        headers: params.response.headers,
        captureAll: true,
      })) {
        return;
      }
  
      const headers = sanitizeHeaders(params.response.headers);
      const lowerHeaders = lowerCaseHeaderMap(params.response.headers);
      const existingResponse = debugState.responses.get(requestId);
      debugState.responses.set(requestId, {
        requestId,
        url,
        status: params.response.status,
        statusText: String(params.response.statusText || ''),
        mimeType: String(params.response.mimeType || ''),
        contentType: String(lowerHeaders['content-type'] || ''),
        resourceType,
        headers,
        fromDiskCache: Boolean(params.response.fromDiskCache),
        fromServiceWorker: Boolean(params.response.fromServiceWorker),
        fromPrefetchCache: Boolean(params.response.fromPrefetchCache),
        encodedDataLength:
          existingResponse?.encodedDataLength ?? null,
        finished: existingResponse?.finished ?? false,
        loadingFailure: existingResponse?.loadingFailure || '',
        remoteIPAddress: String(params.response.remoteIPAddress || ''),
        protocol: String(params.response.protocol || ''),
      });
    });
  
    const stopRequestListener = cdp.on('Network.requestWillBeSent', (params) => {
      if (!debugState) {
        return;
      }
  
      const requestId = String(params.requestId);
      const request = params.request || {};
      const url = String(request.url || '');
      const resourceType = String(params.type || '');
      if (
        !debugState.requests.has(requestId) &&
        debugState.requests.size >= MAX_DEBUG_REQUESTS
      ) {
        return;
      }
  
      if (
        !shouldTrackDebugTraffic({
          url,
          resourceType,
          method: String(request.method || ''),
          headers: request.headers,
          postData: String(request.postData || ''),
          captureAll: true,
        })
      ) {
        return;
      }
  
      debugState.requests.set(requestId, {
        requestId,
        url,
        method: String(request.method || ''),
        resourceType,
        headers: sanitizeHeaders(request.headers),
        hasPostData: typeof request.postData === 'string' && request.postData.length > 0,
        postDataPreview:
          typeof request.postData === 'string'
            ? trimForPreview(
                request.postData,
                MAX_DEBUG_POST_DATA_PREVIEW_LENGTH,
              )
            : '',
        initiatorType: String(params.initiator?.type || ''),
        documentURL: String(params.documentURL || ''),
        timestamp:
          typeof params.timestamp === 'number' ? params.timestamp : null,
        wallTime:
          typeof params.wallTime === 'number' ? params.wallTime : null,
        redirectResponseStatus:
          typeof params.redirectResponse?.status === 'number'
            ? params.redirectResponse.status
            : null,
        redirectResponseUrl: String(params.redirectResponse?.url || ''),
      });
    });
  
    const stopFinishedListener = cdp.on('Network.loadingFinished', (params) => {
      const requestId = String(params.requestId);
      finishedRequests.add(requestId);
      if (transcriptHintRequestIds.has(requestId)) {
        captureFeedback.transcriptHintSettledCount += 1;
        if (captureControl === 'automatic') {
          logAutomaticAction(
            captureFeedback,
            debugEnabled,
            'transcript hint response finished',
            {
              requestId,
            },
          );
        }
      }
  
      if (!debugState) {
        return;
      }
  
      const responseRecord = debugState.responses.get(requestId);
      if (responseRecord) {
        responseRecord.finished = true;
        responseRecord.encodedDataLength =
          typeof params.encodedDataLength === 'number'
            ? params.encodedDataLength
            : null;
      }
    });
  
    const stopFailedListener = cdp.on('Network.loadingFailed', (params) => {
      const requestId = String(params.requestId);
      const errorText = params.errorText || 'Request failed';
      failedRequests.set(requestId, errorText);
      if (transcriptHintRequestIds.has(requestId)) {
        captureFeedback.transcriptHintFailedCount += 1;
        if (captureControl === 'automatic') {
          logAutomaticAction(
            captureFeedback,
            debugEnabled,
            'transcript hint response failed',
            {
              requestId,
              errorText,
            },
          );
        }
      }
  
      if (!debugState) {
        return;
      }
  
      const responseRecord = debugState.responses.get(requestId);
      if (responseRecord) {
        responseRecord.loadingFailure = errorText;
      }
      debugState.failures.push({
        requestId,
        errorText,
        canceled: Boolean(params.canceled),
        blockedReason: String(params.blockedReason || ''),
        resourceType: String(params.type || ''),
      });
    });
  
    const stopWebSocketCreatedListener = cdp.on(
      'Network.webSocketCreated',
      (params) => {
        const socket = ensureTrackedWebSocket(
          params.requestId,
          String(params.url || ''),
        );
        if (!socket) {
          return;
        }
  
        socket.created = true;
        socket.createdAt =
          typeof params.timestamp === 'number' ? params.timestamp : null;
      },
    );
  
    const stopWebSocketHandshakeRequestListener = cdp.on(
      'Network.webSocketWillSendHandshakeRequest',
      (params) => {
        const socket = ensureTrackedWebSocket(params.requestId);
        if (!socket) {
          return;
        }
  
        socket.handshakeRequest = {
          timestamp:
            typeof params.timestamp === 'number' ? params.timestamp : null,
          wallTime:
            typeof params.wallTime === 'number' ? params.wallTime : null,
          headers: sanitizeHeaders(params.request?.headers),
        };
      },
    );
  
    const stopWebSocketHandshakeResponseListener = cdp.on(
      'Network.webSocketHandshakeResponseReceived',
      (params) => {
        const socket = ensureTrackedWebSocket(params.requestId);
        if (!socket) {
          return;
        }
  
        socket.handshakeResponse = {
          timestamp:
            typeof params.timestamp === 'number' ? params.timestamp : null,
          status:
            typeof params.response?.status === 'number'
              ? params.response.status
              : null,
          statusText: String(params.response?.statusText || ''),
          headers: sanitizeHeaders(params.response?.headers),
        };
      },
    );
  
    const stopWebSocketFrameReceivedListener = cdp.on(
      'Network.webSocketFrameReceived',
      (params) => {
        if (!debugState) {
          return;
        }
  
        const socket = ensureTrackedWebSocket(params.requestId);
        if (!socket) {
          return;
        }
  
        socket.framesReceived += 1;
        const payload = summarizeWebSocketPayload(
          params.response?.payloadData || '',
        );
        socket.transcriptSignalDetected =
          socket.transcriptSignalDetected || payload.transcriptSignalDetected;
  
        if (debugState.webSocketFrames.length >= MAX_DEBUG_WEBSOCKET_FRAMES) {
          return;
        }
  
        debugState.webSocketFrames.push({
          requestId: String(params.requestId),
          direction: 'received',
          timestamp:
            typeof params.timestamp === 'number' ? params.timestamp : null,
          opcode:
            typeof params.response?.opcode === 'number'
              ? params.response.opcode
              : null,
          mask: Boolean(params.response?.mask),
          ...payload,
        });
      },
    );
  
    const stopWebSocketFrameSentListener = cdp.on(
      'Network.webSocketFrameSent',
      (params) => {
        if (!debugState) {
          return;
        }
  
        const socket = ensureTrackedWebSocket(params.requestId);
        if (!socket) {
          return;
        }
  
        socket.framesSent += 1;
        const payload = summarizeWebSocketPayload(
          params.response?.payloadData || '',
        );
        socket.transcriptSignalDetected =
          socket.transcriptSignalDetected || payload.transcriptSignalDetected;
  
        if (debugState.webSocketFrames.length >= MAX_DEBUG_WEBSOCKET_FRAMES) {
          return;
        }
  
        debugState.webSocketFrames.push({
          requestId: String(params.requestId),
          direction: 'sent',
          timestamp:
            typeof params.timestamp === 'number' ? params.timestamp : null,
          opcode:
            typeof params.response?.opcode === 'number'
              ? params.response.opcode
              : null,
          mask: Boolean(params.response?.mask),
          ...payload,
        });
      },
    );
  
    const stopWebSocketClosedListener = cdp.on(
      'Network.webSocketClosed',
      (params) => {
        const socket = ensureTrackedWebSocket(params.requestId);
        if (!socket) {
          return;
        }
  
        socket.closedAt =
          typeof params.timestamp === 'number' ? params.timestamp : null;
      },
    );
  
    const stopEventSourceMessageListener = cdp.on(
      'Network.eventSourceMessageReceived',
      (params) => {
        if (!debugState) {
          return;
        }
  
        const requestId = String(params.requestId);
        const requestRecord = debugState.requests.get(requestId);
        const responseRecord = debugState.responses.get(requestId);
        const url = responseRecord?.url || requestRecord?.url || '';
        if (
          !url ||
          !shouldTrackDebugTraffic({
            url,
            resourceType: 'EventSource',
            captureAll: true,
          })
        ) {
          return;
        }
  
        debugState.eventSourceMessages.push({
          requestId,
          timestamp:
            typeof params.timestamp === 'number' ? params.timestamp : null,
          eventName: String(params.eventName || ''),
          eventId: String(params.eventId || ''),
          dataPreview: trimForPreview(
            String(params.data || ''),
            MAX_DEBUG_FRAME_PREVIEW_LENGTH,
          ),
          transcriptSignalDetected: containsTranscriptSignal(params.data || ''),
        });
      },
    );
  
    if (captureControl === 'automatic') {
      await runAutomaticTranscriptCaptureFlow(
        cdp,
        prompt,
        captureFeedback,
        debugEnabled,
      );
    } else if (debugEnabled) {
      console.log('Reloading the page with capture armed...');
      await reloadPageAndWait(cdp);
      printCaptureInstructions(debugEnabled);
      await prompt.waitForEnter('\nPress Enter after the transcript panel has loaded...\n');
      await sleep(DEFAULT_CAPTURE_SETTLE_MS);
    } else {
      printCaptureInstructions(debugEnabled);
      await prompt.waitForEnter('\nPress Enter after the transcript panel has loaded...\n');
      await sleep(DEFAULT_CAPTURE_SETTLE_MS);
    }
  
    stopRequestListener();
    stopResponseListener();
    stopFinishedListener();
    stopFailedListener();
    stopWebSocketCreatedListener();
    stopWebSocketHandshakeRequestListener();
    stopWebSocketHandshakeResponseListener();
    stopWebSocketFrameReceivedListener();
    stopWebSocketFrameSentListener();
    stopWebSocketClosedListener();
    stopEventSourceMessageListener();
  
    const debugBodyTargets = debugState
      ? Array.from(debugState.responses.values())
          .filter(shouldFetchDebugResponseBody)
          .slice(0, MAX_DEBUG_BODIES)
      : [];
    const bodyTargets = new Map();
  
    for (const candidate of trackedResponses.values()) {
      bodyTargets.set(candidate.requestId, {
        requestId: candidate.requestId,
        url: candidate.url,
        status: candidate.status,
        mimeType: candidate.mimeType,
        resourceType: candidate.resourceType,
        contentType: String(
          lowerCaseHeaderMap(candidate.headers)['content-type'] || '',
        ),
        finished: finishedRequests.has(candidate.requestId),
        loadingFailure: failedRequests.get(candidate.requestId) || '',
        wasCandidate: true,
      });
    }
  
    for (const responseRecord of debugBodyTargets) {
      if (bodyTargets.has(responseRecord.requestId)) {
        continue;
      }
  
      bodyTargets.set(responseRecord.requestId, {
        requestId: responseRecord.requestId,
        url: responseRecord.url,
        status: responseRecord.status,
        mimeType: responseRecord.mimeType,
        resourceType: responseRecord.resourceType,
        contentType: responseRecord.contentType,
        finished: responseRecord.finished,
        loadingFailure: responseRecord.loadingFailure || '',
        wasCandidate: false,
      });
    }
  
    const bodyRecordsByRequestId = new Map();
  
    for (const target of bodyTargets.values()) {
      const requestId = target.requestId;
      const failed = target.loadingFailure || failedRequests.get(requestId) || '';
      const finished = target.finished || finishedRequests.has(requestId);
      let bodyRecord;
  
      if (finished && !failed) {
        bodyRecord = await loadResponseBody(cdp, requestId, bodyCache);
      } else if (failed) {
        bodyRecord = {
          body: '',
          bodyLength: 0,
          bodyPreview: '',
          bodyError: failed,
          base64Encoded: false,
          rawBodyBase64: '',
          rawBodyByteLength: 0,
          rawBodyPreviewHex: '',
          refetchedExternally: false,
          refetchError: '',
        };
      } else {
        bodyRecord = {
          body: '',
          bodyLength: 0,
          bodyPreview: '',
          bodyError: 'Request did not finish before capture ended.',
          base64Encoded: false,
          rawBodyBase64: '',
          rawBodyByteLength: 0,
          rawBodyPreviewHex: '',
          refetchedExternally: false,
          refetchError: '',
        };
      }
  
      bodyRecordsByRequestId.set(requestId, bodyRecord);
    }
  
    const protectionKeys = buildProtectionKeyMap(
      Array.from(bodyTargets.values()).map((target) => ({
        url: target.url,
        body: bodyRecordsByRequestId.get(target.requestId)?.body || '',
      })),
    );
  
    for (const target of bodyTargets.values()) {
      if (!isEncryptedTranscriptUrl(target.url)) {
        continue;
      }
  
      const existingBodyRecord = bodyRecordsByRequestId.get(target.requestId);
      if (!existingBodyRecord || existingBodyRecord.bodyError) {
        continue;
      }
  
      const keyId = getUrlSearchParam(target.url, 'kid');
      const hasProtectionKey =
        (keyId && protectionKeys.has(keyId)) || protectionKeys.size === 1;
      if (!hasProtectionKey) {
        continue;
      }
  
      if (existingBodyRecord.base64Encoded && existingBodyRecord.rawBodyBase64) {
        continue;
      }
  
      try {
        const refetchedBodyRecord = await refetchBinaryBody(target.url);
        bodyRecordsByRequestId.set(target.requestId, {
          ...existingBodyRecord,
          ...refetchedBodyRecord,
        });
      } catch (error) {
        bodyRecordsByRequestId.set(target.requestId, {
          ...existingBodyRecord,
          refetchedExternally: false,
          refetchError:
            error instanceof Error ? error.message : 'Transcript refetch failed.',
        });
      }
    }
  
    const candidates = [];
  
    for (const candidate of trackedResponses.values()) {
      const requestId = candidate.requestId;
      const bodyRecord = finalizeBodyRecordForResponse(
        candidate.url,
        bodyRecordsByRequestId.get(requestId) || {
          body: '',
          bodyLength: 0,
          bodyPreview: '',
          bodyError: 'Response body was not captured.',
          base64Encoded: false,
          rawBodyBase64: '',
          rawBodyByteLength: 0,
          rawBodyPreviewHex: '',
          refetchedExternally: false,
          refetchError: '',
        },
        protectionKeys,
      );
      const parsed = summarizeCapturedBody(bodyRecord.body);
      candidates.push({
        ...candidate,
        body: bodyRecord.body,
        bodyLength: bodyRecord.bodyLength,
        bodyPreview: bodyRecord.bodyPreview,
        bodyError: bodyRecord.bodyError,
        base64Encoded: bodyRecord.base64Encoded,
        rawBodyBase64: bodyRecord.rawBodyBase64,
        rawBodyByteLength: bodyRecord.rawBodyByteLength,
        rawBodyPreviewHex: bodyRecord.rawBodyPreviewHex,
        refetchedExternally: bodyRecord.refetchedExternally,
        refetchError: bodyRecord.refetchError,
        decrypted: bodyRecord.decrypted,
        decryptionKeyId: bodyRecord.decryptionKeyId,
        decryptionError: bodyRecord.decryptionError,
        parsedFormat: parsed.format,
        parsedPath: parsed.path,
        parsedEntryCount: parsed.entryCount,
        parsedEntries: parsed.entries,
      });
    }
  
    let debug = null;
  
    if (debugState) {
      const debugBodies = [];
  
      for (const responseRecord of debugBodyTargets) {
        const bodyRecord = finalizeBodyRecordForResponse(
          responseRecord.url,
          bodyRecordsByRequestId.get(responseRecord.requestId) || {
            body: '',
            bodyLength: 0,
            bodyPreview: '',
            bodyError: 'Response body was not captured.',
            base64Encoded: false,
            rawBodyBase64: '',
            rawBodyByteLength: 0,
            rawBodyPreviewHex: '',
            refetchedExternally: false,
            refetchError: '',
          },
          protectionKeys,
        );
        const parsed = summarizeCapturedBody(bodyRecord.body);
        debugBodies.push({
          requestId: responseRecord.requestId,
          url: responseRecord.url,
          status: responseRecord.status,
          mimeType: responseRecord.mimeType,
          contentType: responseRecord.contentType,
          resourceType: responseRecord.resourceType,
          bodyLength: bodyRecord.bodyLength,
          bodyPreview: bodyRecord.bodyPreview,
          body: bodyRecord.body,
          bodyError: bodyRecord.bodyError,
          base64Encoded: bodyRecord.base64Encoded,
          rawBodyBase64: bodyRecord.rawBodyBase64,
          rawBodyByteLength: bodyRecord.rawBodyByteLength,
          rawBodyPreviewHex: bodyRecord.rawBodyPreviewHex,
          refetchedExternally: bodyRecord.refetchedExternally,
          refetchError: bodyRecord.refetchError,
          decrypted: bodyRecord.decrypted,
          decryptionKeyId: bodyRecord.decryptionKeyId,
          decryptionError: bodyRecord.decryptionError,
          parsedFormat: parsed.format,
          parsedPath: parsed.path,
          parsedEntryCount: parsed.entryCount,
          parsedEntrySample: parsed.entries.slice(0, 5),
          wasCandidate: trackedResponses.has(responseRecord.requestId),
        });
      }
  
      debug = {
        requestCount: debugState.requests.size,
        responseCount: debugState.responses.size,
        failureCount: debugState.failures.length,
        webSocketCount: debugState.webSockets.size,
        webSocketFrameCount: debugState.webSocketFrames.length,
        eventSourceMessageCount: debugState.eventSourceMessages.length,
        protectionKeys: Array.from(protectionKeys.values()).map((key) => ({
          kid: key.kid,
          sourceUrl: key.sourceUrl,
          encryptionAlgorithm: key.encryptionAlgorithm,
          encryptionMode: key.encryptionMode,
          keySize: key.keySize,
          padding: key.padding,
        })),
        requests: Array.from(debugState.requests.values()),
        responses: Array.from(debugState.responses.values()),
        failures: debugState.failures,
        responseBodies: debugBodies,
        webSockets: Array.from(debugState.webSockets.values()),
        webSocketFrames: debugState.webSocketFrames,
        eventSourceMessages: debugState.eventSourceMessages,
      };
    }
  
    await cdp.send('Network.disable');
  
    candidates.sort((left, right) => {
      if (right.parsedEntryCount !== left.parsedEntryCount) {
        return right.parsedEntryCount - left.parsedEntryCount;
      }
  
      return right.score - left.score;
    });

    const transcriptMatches = candidates.filter(
      (candidate) => candidate.parsedEntryCount > 0,
    );
  
    const matchedCandidate =
      transcriptMatches[0] || null;
  
    return {
      candidates,
      matchedCandidate,
      transcriptMatchCount: transcriptMatches.length,
      captureFeedback: {
        candidateResponseCount: captureFeedback.candidateResponseCount,
        transcriptHintCount: captureFeedback.transcriptHintCount,
        transcriptHintSettledCount: captureFeedback.transcriptHintSettledCount,
        transcriptHintFailedCount: captureFeedback.transcriptHintFailedCount,
        firstTranscriptHint: captureFeedback.firstTranscriptHint,
        automaticActions:
          captureFeedback.automaticActions.slice(0, 200),
      },
      debug,
    };
  }
  
  function buildNetworkCapturePayload({
    options,
    browser,
    profile,
    debugPort,
    targetPage,
    captureResult,
  }) {
    const visibleCandidates = captureResult.candidates.slice(0, MAX_SAVED_CANDIDATES);
  
    return {
      app: {
        name: APP_NAME,
        version: BUILD_VERSION,
        buildTime: BUILD_TIME,
        platform: CURRENT_PLATFORM,
      },
      run: {
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
        targetPage: targetPage
          ? {
              title: targetPage.title,
              url: targetPage.url,
            }
          : null,
        outputFormat: options.outputFormat,
        debugPort,
        capturedAt: new Date().toISOString(),
        debugCaptureEnabled: options.debug,
      },
      summary: {
        candidateCount: captureResult.candidates.length,
        captureFeedback: captureResult.captureFeedback || null,
        transcriptMatches: captureResult.transcriptMatchCount || 0,
        matchedCandidate: captureResult.matchedCandidate
          ? {
              url: captureResult.matchedCandidate.url,
              mimeType: captureResult.matchedCandidate.mimeType,
              resourceType: captureResult.matchedCandidate.resourceType,
              decrypted: captureResult.matchedCandidate.decrypted,
              decryptionKeyId: captureResult.matchedCandidate.decryptionKeyId,
              parsedFormat: captureResult.matchedCandidate.parsedFormat,
              parsedPath: captureResult.matchedCandidate.parsedPath,
              parsedEntryCount: captureResult.matchedCandidate.parsedEntryCount,
            }
          : null,
        debug:
          options.debug && captureResult.debug
            ? {
                requestCount: captureResult.debug.requestCount,
                responseCount: captureResult.debug.responseCount,
                failureCount: captureResult.debug.failureCount,
                webSocketCount: captureResult.debug.webSocketCount,
                webSocketFrameCount: captureResult.debug.webSocketFrameCount,
                eventSourceMessageCount:
                  captureResult.debug.eventSourceMessageCount,
                responseBodyCount: captureResult.debug.responseBodies.length,
                protectionKeyCount: captureResult.debug.protectionKeys.length,
              }
            : null,
      },
      candidates: visibleCandidates.map((candidate) => ({
        requestId: candidate.requestId,
        url: candidate.url,
        status: candidate.status,
        mimeType: candidate.mimeType,
        resourceType: candidate.resourceType,
        score: candidate.score,
        bodyLength: candidate.bodyLength,
        bodyError: candidate.bodyError,
        base64Encoded: candidate.base64Encoded,
        rawBodyByteLength: candidate.rawBodyByteLength,
        rawBodyPreviewHex: candidate.rawBodyPreviewHex,
        refetchedExternally: candidate.refetchedExternally,
        refetchError: candidate.refetchError,
        decrypted: candidate.decrypted,
        decryptionKeyId: candidate.decryptionKeyId,
        decryptionError: candidate.decryptionError,
        parsedFormat: candidate.parsedFormat,
        parsedPath: candidate.parsedPath,
        parsedEntryCount: candidate.parsedEntryCount,
        parsedEntrySample: candidate.parsedEntries.slice(0, 5),
        headers: candidate.headers,
        bodyPreview: candidate.bodyPreview,
        body: options.debug ? candidate.body : undefined,
        rawBodyBase64:
          options.debug && candidate.base64Encoded
            ? candidate.rawBodyBase64
            : undefined,
      })),
      debug: options.debug ? captureResult.debug : undefined,
    };
  }

  function buildMissingTranscriptCaptureMessage(
    captureResult,
    debugEnabled,
    captureControl = 'manual',
  ) {
    const candidateResponseCount =
      captureResult.captureFeedback?.candidateResponseCount || 0;
    const transcriptHintCount =
      captureResult.captureFeedback?.transcriptHintCount || 0;

    let guidance =
      'No transcript payload was observed after network capture started.';

    if (candidateResponseCount === 0) {
      guidance +=
        ' No transcript-related network responses were captured while the terminal was armed.';
    } else if (transcriptHintCount === 0) {
      guidance +=
        ' Some Stream responses were captured, but none looked transcript-specific.';
    } else {
      guidance +=
        ' Transcript-like traffic was detected, but none of the captured bodies parsed into transcript entries.';
    }

    if (debugEnabled) {
      return `${guidance} Review the saved .network.json capture.`;
    }

    if (captureControl === 'automatic') {
      return (
        `${guidance} Rerun with --debug to save a .network.json capture and ` +
        'print the automatic action trace.'
      );
    }

    return (
      `${guidance} Rerun with --debug to save a .network.json capture and ` +
      'auto-reload the page once capture is armed.'
    );
  }
  
  function handleError(error) {
    if (error instanceof CliError) {
      console.error(`\nError: ${error.message}`);
      return error.exitCode;
    }
  
    console.error('\nUnexpected error:');
    console.error(error);
    return 1;
  }
  
  async function run(captureControl = 'manual') {
    const options = parseArgs(process.argv.slice(2));
  
    if (options.help) {
      printHelp();
      return 0;
    }
  
    if (options.version) {
      printVersion();
      return 0;
    }
  
    ensureSupportedPlatform();
  
    const prompt = createPrompt();
    let browserProcess = null;
    let cdp = null;
    let tempDataDir = null;
    let browser = null;
    let profile = null;
    let debugPort = null;
    let targetPage = null;
  
    try {
      ({ browser, profile } = await selectBrowserAndProfile(options, prompt));
      debugPort = await findAvailablePort(options.debugPort);
  
      console.log(`\nUsing ${browser.name} / "${profile.displayName}".`);
      await ensureBrowserIsClosed(prompt, browser);
  
      tempDataDir = join(tmpdir(), `stream-transcript-extractor-network-${Date.now()}`);
      mkdirSync(tempDataDir, { recursive: true });
  
      console.log('Preparing temporary browser profile...');
      prepareTempProfile(browser.basePath, profile, tempDataDir);
  
      console.log(`Launching ${browser.name} on debug port ${debugPort}...`);
      browserProcess = launchBrowser(browser, profile, tempDataDir, debugPort);
  
      console.log('Waiting for the browser debug endpoint...');
      await waitForBrowserDebugEndpoint(debugPort);
      await waitForBrowserPageTarget(debugPort);
  
      printInitialRunInstructions(captureControl);
      await prompt.waitForEnter(
        '\nPress Enter once the Stream meeting page is open and ready...\n',
      );
  
      targetPage = await selectTranscriptPage(prompt, debugPort);
      console.log(`\nConnecting to: ${targetPage.title}`);
  
      cdp = await connectToPage(targetPage.webSocketDebuggerUrl);
      await cdp.send('Runtime.enable');
  
      console.log(
        captureControl === 'automatic'
          ? 'Capturing transcript-related network responses in automatic mode...'
          : 'Capturing transcript-related network responses...',
      );
      if (options.debug) {
        console.log(
          'Debug capture enabled: saving request/response lifecycle data, ' +
            'candidate bodies, and WebSocket frames.',
        );
      }
      const captureResult = await captureStreamNetwork(
        cdp,
        prompt,
        options.debug,
        captureControl,
      );
      const pageMetadata = await extractMeetingMetadata(cdp).catch(() => ({
        title: targetPage?.title || '',
        date: '',
        recordedBy: '',
        sourceUrl: targetPage?.url || '',
        createdBy: '',
        createdByEmail: '',
        createdByTenantId: '',
        sourceApplication: '',
        recordingStartDateTime: '',
        recordingEndDateTime: '',
        sharePointFilePath: '',
        sharePointItemUrl: '',
      }));
      const captureMetadata = extractMeetingMetadataFromCapture(captureResult);
      const metadata = mergeMeetingMetadata(
        pageMetadata,
        captureMetadata,
        targetPage,
      );
      const outputBasePath = buildOutputBasePath(
        metadata.title || targetPage?.title || 'meeting',
        options.outputName,
        options.outputDir,
      );
      let networkOutputPath = '';

      console.log(
        `Observed ${captureResult.candidates.length} potentially relevant network response` +
          `${captureResult.candidates.length === 1 ? '' : 's'}.`,
      );
      console.log(
        `Parsed ${captureResult.transcriptMatchCount || 0} transcript payload match` +
          `${captureResult.transcriptMatchCount === 1 ? '' : 'es'}.`,
      );

      if (!captureResult.matchedCandidate) {
        if (options.debug) {
          const networkCapturePayload = buildNetworkCapturePayload({
            options,
            browser,
            profile,
            debugPort,
            targetPage,
            captureResult,
          });
          networkOutputPath = saveNetworkCaptureOutput(
            networkCapturePayload,
            outputBasePath,
          );
          console.log(`Saved network capture to: ${networkOutputPath}`);
        }

        throw new CliError(
          buildMissingTranscriptCaptureMessage(
            captureResult,
            options.debug,
            captureControl,
          ),
        );
      }

      const networkCapturePayload = buildNetworkCapturePayload({
        options,
        browser,
        profile,
        debugPort,
        targetPage,
        captureResult,
      });
      networkOutputPath = saveNetworkCaptureOutput(
        networkCapturePayload,
        outputBasePath,
      );

      const entries = captureResult.matchedCandidate.parsedEntries;
      const outputPayload = buildOutputPayload(metadata, entries);
      const outputPaths = saveOutputs(
        outputPayload,
        outputBasePath,
        options.outputFormat,
      );
  
      console.log(
        `Matched transcript payload: ${captureResult.matchedCandidate.url}`,
      );
      console.log(`Parsed ${entries.length} transcript entries.`);
      for (const outputPath of outputPaths) {
        console.log(`Saved transcript to: ${outputPath}`);
      }
      console.log(`Saved network capture to: ${networkOutputPath}`);
  
      return 0;
    } catch (error) {
      return handleError(error);
    } finally {
      prompt.close();
  
      if (cdp) {
        cdp.close();
      }
  
      if (!options.keepBrowserOpen && browserProcess?.pid) {
        await shutdownBrowser(browserProcess, debugPort);
      }
  
      if (options.keepBrowserOpen && tempDataDir) {
        console.log(
          `Temporary profile preserved because --keep-browser-open is set: ${tempDataDir}`,
        );
      } else if (tempDataDir) {
        try {
          rmSync(tempDataDir, { recursive: true, force: true });
        } catch {
          // Ignore cleanup failures.
        }
      }
    }
  }

  return async function runEmbeddedNetworkMode(captureControl = 'manual') {
    return run(captureControl);
  };
})();

async function runSelectedMode() {
  const { mode, forwardedArgs } = resolveExtractorMode(process.argv.slice(2));
  const originalArgv = [...process.argv];
  process.argv = [originalArgv[0], originalArgv[1], ...forwardedArgs];

  try {
    if (forwardedArgs.includes('--help') || forwardedArgs.includes('-h')) {
      printEntrypointHelp();
      return 0;
    }

    if (forwardedArgs.includes('--version') || forwardedArgs.includes('-v')) {
      printEntrypointVersion();
      return 0;
    }

    if (mode === 'dom') {
      return await runDomMode();
    }

    if (mode === 'automatic') {
      return await runEmbeddedNetworkMode('automatic');
    }

    return await runEmbeddedNetworkMode('manual');
  } finally {
    process.argv = originalArgv;
  }
}

let exitCode = 0;

try {
  exitCode = await runSelectedMode();
} catch (error) {
  exitCode = handleEntrypointError(error);
}

process.exitCode = exitCode;
