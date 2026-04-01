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
import { dirname, join, resolve } from 'node:path';
import {
  connectCdp as connectCoreCdp,
  createCdpEventScope,
  evaluate as evaluateWithCdp,
  findPageTargets as findCorePageTargets,
  getBrowserDebuggerWebSocketUrl as getCoreBrowserDebuggerWebSocketUrl,
  navigatePageAndWait as navigateWithCdp,
  reloadPageAndWait as reloadWithCdp,
  waitForBrowserDebugEndpoint as waitForCoreBrowserDebugEndpoint,
  waitForBrowserDebugEndpointToClose as waitForCoreBrowserDebugEndpointToClose,
  waitForBrowserPageTarget as waitForCoreBrowserPageTarget,
} from './lib/cdp.js';
import {
  chooseFromList as chooseCliFromList,
  chooseManyFromList as chooseCliManyFromList,
  createPrompt as createCliPrompt,
  parseCliArgs as parseSharedCliArgs,
  parseSelectionSpec as parseCliSelectionSpec,
  printHelpScreen,
} from './lib/cli.js';
import {
  createResponseBodyRecord,
  loadResponseBody as loadCapturedResponseBody,
} from './lib/network-capture.js';

const APP_NAME = 'Stream Transcript Extractor';
const APP_DESCRIPTION =
  'Extract Microsoft Teams recording transcripts from Microsoft Stream ' +
  'using your signed-in Chrome or Edge profile.';
const DEFAULT_WORKFLOW = 'extract';
const SUPPORTED_WORKFLOWS = ['extract', 'crawl'];
const DEFAULT_EXTRACTOR_MODE = 'automatic';
const SUPPORTED_EXTRACTOR_MODES = ['network', 'automatic', 'dom'];
const SUPPORTED_BROWSER_KEYS = ['chrome', 'edge'];
const SUPPORTED_OUTPUT_FORMATS = ['json', 'md', 'both'];
const DEFAULT_TRANSCRIPT_OUTPUT_FORMAT = 'md';
const DEFAULT_OUTPUT_DIR = 'output';
const DEFAULT_CRAWL_START_URL =
  'https://m365.cloud.microsoft/launch/Stream/?auth=2&home=1';
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
 * This repository keeps the Stream-specific runtime centered in one shareable
 * source file while the reusable CLI and CDP helpers live in `lib/`.
 * The module is split into three logical sections:
 * 1. The shared CLI entrypoint and DOM fallback implementation
 * 2. The embedded Stream network extractor, kept self-contained in a closure
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

function normalizeExtractorMode(value) {
  return String(value || '').trim().toLowerCase();
}

function normalizeWorkflow(value) {
  return String(value || '').trim().toLowerCase();
}

/**
 * Resolve the top-level workflow before handing control to the selected
 * runtime. `extract` remains the default workflow, while `crawl` wraps the
 * automatic extractor flow in the batch discovery worker.
 */
function resolveWorkflowSelection(argv) {
  let workflow = DEFAULT_WORKFLOW;
  let workflowProvided = false;
  let mode = DEFAULT_EXTRACTOR_MODE;
  let modeProvided = false;
  const forwardedArgs = [];
  let startIndex = 0;

  if (argv.length > 0 && !String(argv[0]).startsWith('-')) {
    const candidate = normalizeWorkflow(argv[0]);
    if (!SUPPORTED_WORKFLOWS.includes(candidate)) {
      throw new CliError(
        `Unsupported workflow "${argv[0]}". Use one of: ` +
          `${SUPPORTED_WORKFLOWS.join(', ')}.`,
      );
    }

    workflow = candidate;
    workflowProvided = true;
    startIndex = 1;
  }

  for (let index = startIndex; index < argv.length; index += 1) {
    const arg = argv[index];
    const [flag, inlineValue] = arg.split(/=(.*)/s, 2);

    if (flag !== '--mode') {
      forwardedArgs.push(arg);
      continue;
    }

    if (workflow !== 'extract') {
      throw new CliError('--mode is only supported for the extract workflow.');
    }

    const value = readOptionValue(flag, inlineValue, argv[index + 1]);
    mode = normalizeExtractorMode(value);
    modeProvided = true;

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
    workflow,
    workflowProvided,
    mode,
    modeProvided,
    forwardedArgs,
  };
}

/**
 * @param {string[]} argv
 * @returns {CliOptions}
 */
function parseDomModeArgs(argv) {
  return parseSharedCliArgs(argv, {
    defaults: {
      browser: '',
      profile: '',
      outputName: '',
      outputDir: DEFAULT_OUTPUT_DIR,
      outputFormat: DEFAULT_TRANSCRIPT_OUTPUT_FORMAT,
      debugPort: null,
      debug: false,
      keepBrowserOpen: false,
      help: false,
      version: false,
    },
    supportedBrowserKeys: SUPPORTED_BROWSER_KEYS,
    supportedOutputFormats: SUPPORTED_OUTPUT_FORMATS,
    normalizeBrowserKey,
    errorFactory: (message) => new CliError(message),
  });
}

function printEntrypointHelp() {
  printHelpScreen({
    name: APP_NAME,
    summary: APP_DESCRIPTION,
    usage: [
      'bun ./extract.js [options]',
      'bun ./extract.js crawl [options]',
      './stream-transcript-extractor-<target> [workflow] [options]',
    ],
    sections: [
      {
        title: 'Start here',
        rows: [
          {
            label: 'bun ./extract.js',
            description:
              'Open the interactive launcher in a real terminal. It lets you choose extract or crawl, then applies the same workflow settings for source and compiled builds.',
          },
          {
            label: './stream-transcript-extractor-<target>',
            description:
              'Open the same interactive launcher from a compiled binary in a real terminal.',
          },
          {
            label: 'bun ./extract.js crawl',
            description:
              'Open Stream home, switch to Meetings, scroll the rendered list, merge the results into a persistent queue, then run automatic extraction against the items you select in the same browser session.',
          },
        ],
      },
      {
        title: 'Workflows',
        rows: [
          {
            label: 'crawl',
            description:
              'Batch flow. Opens Stream home, selects Meetings, scrolls the visible page, merges results into a *.state.json queue, then wraps automatic extraction across the items you select.',
          },
          {
            label: 'extract (default)',
            description:
              'Single-meeting extraction. Supports --mode network, automatic, or dom.',
          },
        ],
      },
      {
        title: 'Shared options',
        rows: [
          { label: '--browser <chrome|edge>', description: 'Use a specific browser.' },
          {
            label: '--profile <query>',
            description: 'Match a profile by name, email, or directory.',
          },
          {
            label: '--output <name>',
            description: 'Override the transcript filename prefix.',
          },
          {
            label: '--output-dir <path>',
            description: 'Write output files to a custom directory.',
          },
          {
            label: '--format <json|md|both>',
            description:
              'Transcript output format. Recommended and default: md. Use json for structured automation or both when you need both.',
          },
          {
            label: '--debug-port <port>',
            description: 'Force a specific remote-debugging port.',
          },
          {
            label: '--debug',
            description:
              'Write extra diagnostics. Recommended default: leave this off for normal runs. Turn it on when you are diagnosing a failed or suspicious extraction.',
          },
          {
            label: '--keep-browser-open',
            description: 'Leave the launched browser open after extraction.',
          },
          { label: '--version, -v', description: 'Print the build version.' },
          { label: '--help, -h', description: 'Show this help text.' },
        ],
      },
      {
        title: 'Workflow-specific options',
        lines: [
          `extract supports --mode <${SUPPORTED_EXTRACTOR_MODES.join('|')}>. Recommended and default: ${DEFAULT_EXTRACTOR_MODE}.`,
          `crawl supports --start-url <url>. Default: ${DEFAULT_CRAWL_START_URL}. It always wraps the automatic extractor in one browser session.`,
          `Both workflows default to --format ${DEFAULT_TRANSCRIPT_OUTPUT_FORMAT}.`,
          'Run `bun ./extract.js --mode automatic --help` or `bun ./extract.js crawl --help` for workflow-specific help.',
        ],
      },
      {
        title: 'Examples',
        lines: [
          'bun ./extract.js',
          'bun ./extract.js --browser chrome --profile Work',
          'bun ./extract.js crawl',
          'bun ./extract.js crawl --state-file ./exports/team.state.json',
          'bun ./extract.js --mode network --output-dir ./exports --format both',
          'bun ./extract.js --mode dom --debug',
        ],
      },
    ],
  });
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

function isInteractiveEntrypointLaunch(argv) {
  return (
    Array.isArray(argv) &&
    argv.length === 0 &&
    Boolean(process.stdin.isTTY) &&
    Boolean(process.stdout.isTTY)
  );
}

function renderInteractiveMenuChoice(choice) {
  const label = String(choice?.label || '').trim();
  const description = String(choice?.description || '').trim();
  return description ? `${label} | ${description}` : label;
}

async function chooseInteractiveMenuChoice(prompt, title, question, choices) {
  return chooseFromList(
    prompt,
    title,
    choices,
    renderInteractiveMenuChoice,
    question,
  );
}

async function askOptionalValue(prompt, question) {
  return String(await prompt.ask(question)).trim();
}

async function promptForInteractiveLaunchArgs() {
  const prompt = createPrompt();

  try {
    console.log('\nInteractive launch');
    console.log(
      'Choose a workflow and the common settings here. Advanced overrides still stay available as normal CLI flags.',
    );

    const workflowChoice = await chooseInteractiveMenuChoice(
      prompt,
      'Workflow',
      '\nChoose a workflow (number): ',
      [
        {
          value: 'crawl',
          label: 'crawl',
          description:
            'Open Stream home, wait 10 seconds, switch to Meetings, scroll the page, build or update the queue, then extract selected items in the same browser session.',
        },
        {
          value: 'extract',
          label: 'extract',
          description:
            'Open one recording and run the single-meeting extractor. Recommended when you only need one transcript.',
        },
      ],
    );

    const args = [];
    if (workflowChoice.value === 'crawl') {
      args.push('crawl');
    } else {
      const modeChoice = await chooseInteractiveMenuChoice(
        prompt,
        'Extract mode',
        '\nChoose an extract mode (number): ',
        [
          {
            value: 'automatic',
            label: 'automatic (Recommended)',
            description:
              'Lowest operator effort. Reloads with capture armed and tries the Transcript panel for you.',
          },
          {
            value: 'network',
            label: 'network',
            description:
              'Advanced path. You manage transcript-panel timing yourself and capture the network payload directly.',
          },
          {
            value: 'dom',
            label: 'dom',
            description:
              'Fallback path. Reads the visible transcript UI instead of relying on the transport layer.',
          },
        ],
      );
      args.push('--mode', modeChoice.value);
    }

    const formatChoice = await chooseInteractiveMenuChoice(
      prompt,
      'Transcript format',
      '\nChoose an output format (number): ',
      [
        {
          value: 'md',
          label: 'md (Recommended)',
          description:
            'Best for review, sharing, and follow-up editing.',
        },
        {
          value: 'json',
          label: 'json',
          description:
            'Structured output for automation or downstream processing.',
        },
        {
          value: 'both',
          label: 'both',
          description:
            'Write both Markdown and JSON for the same run.',
        },
      ],
    );
    args.push('--format', formatChoice.value);

    const debugChoice = await chooseInteractiveMenuChoice(
      prompt,
      'Diagnostics',
      '\nChoose a diagnostics level (number): ',
      [
        {
          value: 'off',
          label: 'debug off (Recommended)',
          description:
            'Normal operator path. Keep output and terminal noise minimal.',
        },
        {
          value: 'on',
          label: 'debug on',
          description:
            'Write deeper diagnostics and print the automatic action trace when supported.',
        },
      ],
    );
    if (debugChoice.value === 'on') {
      args.push('--debug');
    }

    const browserChoice = await chooseInteractiveMenuChoice(
      prompt,
      'Browser',
      '\nChoose a browser preference (number): ',
      [
        {
          value: '',
          label: 'auto detect (Recommended)',
          description:
            'Choose later from available local browsers if needed.',
        },
        {
          value: 'edge',
          label: 'edge',
          description:
            'Prefer Microsoft Edge.',
        },
        {
          value: 'chrome',
          label: 'chrome',
          description:
            'Prefer Google Chrome.',
        },
      ],
    );
    if (browserChoice.value) {
      args.push('--browser', browserChoice.value);
    }

    if (workflowChoice.value === 'crawl') {
      const waitChoice = await chooseInteractiveMenuChoice(
        prompt,
        'Crawl settle wait',
        '\nChoose the wait before discovery starts (number): ',
        [
          {
            value: '10000',
            label: '10 seconds (Recommended)',
            description:
              'Best default for normal auth and page load settle time.',
          },
          {
            value: '30000',
            label: '30 seconds',
            description:
              'Use when Stream auth or page hydration is unusually slow.',
          },
          {
            value: 'custom',
            label: 'custom',
            description:
              'Enter your own wait duration in milliseconds.',
          },
        ],
      );

      let waitBeforeDiscoveryMs = waitChoice.value;
      if (waitChoice.value === 'custom') {
        while (true) {
          const customWait = await askOptionalValue(
            prompt,
            '\nEnter wait-before-discovery in milliseconds: ',
          );
          const parsed = Number.parseInt(customWait, 10);
          if (Number.isInteger(parsed) && parsed >= 0) {
            waitBeforeDiscoveryMs = String(parsed);
            break;
          }
          console.log('Enter a whole number greater than or equal to 0.');
        }
      }

      args.push('--wait-before-discovery-ms', waitBeforeDiscoveryMs);
    }

    const keepOpenChoice = await chooseInteractiveMenuChoice(
      prompt,
      'Browser lifecycle',
      '\nKeep the launched browser open after the run? (number): ',
      [
        {
          value: 'no',
          label: 'no (Recommended)',
          description:
            'Close the temporary browser session when the run completes.',
        },
        {
          value: 'yes',
          label: 'yes',
          description:
            'Preserve the launched browser for manual follow-up after the run.',
        },
      ],
    );
    if (keepOpenChoice.value === 'yes') {
      args.push('--keep-browser-open');
    }

    const advancedChoice = await chooseInteractiveMenuChoice(
      prompt,
      'Advanced overrides',
      '\nConfigure advanced optional values? (number): ',
      [
        {
          value: 'no',
          label: 'no (Recommended)',
          description:
            'Use the recommended launcher path and keep the remaining values at their defaults.',
        },
        {
          value: 'yes',
          label: 'yes',
          description:
            'Set optional paths and workflow-specific overrides before launch.',
        },
      ],
    );

    if (advancedChoice.value === 'yes') {
      const outputDir = await askOptionalValue(
        prompt,
        '\nOutput directory (Enter = default "output"): ',
      );
      if (outputDir) {
        args.push('--output-dir', outputDir);
      }

      const profileQuery = await askOptionalValue(
        prompt,
        '\nProfile match (Enter = choose later or auto-pick single profile): ',
      );
      if (profileQuery) {
        args.push('--profile', profileQuery);
      }

      const outputName = await askOptionalValue(
        prompt,
        '\nOutput filename prefix (Enter = default naming): ',
      );
      if (outputName) {
        args.push('--output', outputName);
      }

      const debugPort = await askOptionalValue(
        prompt,
        '\nDebug port (Enter = automatic): ',
      );
      if (debugPort) {
        args.push('--debug-port', debugPort);
      }

      if (workflowChoice.value === 'crawl') {
        const startUrl = await askOptionalValue(
          prompt,
          '\nStart URL (Enter = default Stream home): ',
        );
        if (startUrl) {
          args.push('--start-url', startUrl);
        }

        const stateFile = await askOptionalValue(
          prompt,
          '\nState file path (Enter = generated/default path): ',
        );
        if (stateFile) {
          args.push('--state-file', stateFile);
        }

        const selectionSpec = await askOptionalValue(
          prompt,
          '\nSelection spec (Enter = choose later in the queue menu): ',
        );
        if (selectionSpec) {
          args.push('--select', selectionSpec);
        }
      }
    }

    return args;
  } finally {
    prompt.close();
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

async function connectCdp(websocketUrl) {
  return connectCoreCdp(websocketUrl, {
    defaultTimeoutMs: DEFAULT_CDP_TIMEOUT_MS,
    errorFactory: (message) => new CliError(message),
  });
}

async function findPageTargets(port) {
  return findCorePageTargets({
    host: DEFAULT_CDP_HOST,
    port,
    errorFactory: (message) => new CliError(message),
  });
}

async function connectToPage(pageWebsocketUrl) {
  const cdp = await connectCdp(pageWebsocketUrl);
  await cdp.send('Runtime.enable');
  return cdp;
}

async function evaluate(cdp, expression, timeoutMs = DEFAULT_CDP_TIMEOUT_MS) {
  return evaluateWithCdp(cdp, expression, {
    timeoutMs,
    errorFactory: (message) => new CliError(message),
  });
}

async function waitForBrowserDebugEndpoint(port) {
  return waitForCoreBrowserDebugEndpoint({
    host: DEFAULT_CDP_HOST,
    port,
    timeoutMs: DEFAULT_BROWSER_READY_TIMEOUT_MS,
    errorFactory: (message) => new CliError(message),
  });
}

async function waitForBrowserPageTarget(
  port,
  timeoutMs = DEFAULT_BROWSER_READY_TIMEOUT_MS,
) {
  return waitForCoreBrowserPageTarget({
    host: DEFAULT_CDP_HOST,
    port,
    timeoutMs,
    errorFactory: (message) => new CliError(message),
  });
}

async function waitForBrowserDebugEndpointToClose(
  port,
  timeoutMs = DEFAULT_BROWSER_SHUTDOWN_TIMEOUT_MS,
) {
  return waitForCoreBrowserDebugEndpointToClose({
    host: DEFAULT_CDP_HOST,
    port,
    timeoutMs,
  });
}

async function getBrowserDebuggerWebSocketUrl(port) {
  return getCoreBrowserDebuggerWebSocketUrl({
    host: DEFAULT_CDP_HOST,
    port,
  });
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
const networkModeRuntime = (() => {
  const APP_NAME = 'Stream Transcript Extractor (Network Mode)';
  const APP_DESCRIPTION =
    'Extract Microsoft Teams recording transcripts from Microsoft Stream ' +
    'using your signed-in Chrome or Edge profile.';
  const SUPPORTED_BROWSER_KEYS = ['chrome', 'edge'];
  const SUPPORTED_OUTPUT_FORMATS = ['json', 'md', 'both'];
  const DEFAULT_OUTPUT_DIR = 'output';
  const DEFAULT_CRAWL_START_URL =
    'https://m365.cloud.microsoft/launch/Stream/?auth=2&home=1';
  const DEFAULT_CDP_HOST = '127.0.0.1';
  const DEFAULT_BROWSER_READY_TIMEOUT_MS = 15_000;
  const DEFAULT_CDP_TIMEOUT_MS = 30_000;
  const DEFAULT_PAGE_NAVIGATION_TIMEOUT_MS = 45_000;
  const DEFAULT_CAPTURE_SETTLE_MS = 1_500;
  const DEFAULT_CRAWL_SCROLL_SETTLE_MS = 1_000;
  const DEFAULT_WAIT_BEFORE_DISCOVERY_MS = 10_000;
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
  const MAX_DISCOVERY_RESPONSE_BODIES = 200;
  const CRAWL_STATE_VERSION = 1;

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

  /**
   * @param {string[]} argv
   * @returns {CliOptions}
   */
  function parseArgs(argv) {
    return parseSharedCliArgs(argv, {
      defaults: {
        browser: '',
        profile: '',
        outputName: '',
        outputDir: DEFAULT_OUTPUT_DIR,
        outputFormat: DEFAULT_TRANSCRIPT_OUTPUT_FORMAT,
        debugPort: null,
        debug: false,
        keepBrowserOpen: false,
        help: false,
        version: false,
      },
      supportedBrowserKeys: SUPPORTED_BROWSER_KEYS,
      supportedOutputFormats: SUPPORTED_OUTPUT_FORMATS,
      normalizeBrowserKey,
      errorFactory: (message) => new CliError(message),
    });
  }

  function printHelp() {
    printHelpScreen({
      name: APP_NAME,
      summary: APP_DESCRIPTION,
      usage: [
        'bun ./extract.js --mode <network|automatic> [options]',
        './stream-transcript-extractor-<target> [options]',
      ],
      sections: [
        {
          title: 'Recommended path',
          rows: [
            {
              label: 'automatic (recommended default)',
              description:
                'Recommended for first success. Reload with capture armed, try the Transcript panel automatically, and only ask for manual help outside batch mode when the page still needs it.',
              },
            {
              label: 'network (advanced)',
              description:
                'Use when you want the lowest-level network capture flow and are comfortable opening the Transcript panel yourself after arming capture.',
            },
          ],
        },
        {
          title: 'Recommended settings',
          rows: [
            {
              label: '--format md',
              description:
                'Recommended and default transcript output for review, sharing, and follow-up editing.',
            },
            {
              label: '--debug off',
              description:
                'Recommended normal path. Turn on --debug only when you want deeper diagnostics or a saved network trace for triage.',
            },
          ],
        },
        {
          title: 'Options',
          rows: [
            { label: '--browser <chrome|edge>', description: 'Use a specific browser.' },
            {
              label: '--profile <query>',
              description: 'Match a profile by name, email, or directory.',
            },
            {
              label: '--output <name>',
              description: 'Override the transcript filename prefix.',
            },
            {
              label: '--output-dir <path>',
              description: 'Write files to a custom directory.',
            },
            {
              label: '--format <json|md|both>',
              description:
                'Choose JSON, Markdown, or both outputs. Recommended and default: md.',
            },
            {
              label: '--debug-port <port>',
              description: 'Force a specific remote-debugging port.',
            },
            {
              label: '--debug',
              description:
                'Save extended network diagnostics, including candidate bodies and the automatic action trace. Recommended default: off.',
            },
            {
              label: '--keep-browser-open',
              description: 'Leave the launched browser open after extraction.',
            },
            { label: '--version, -v', description: 'Print the build version.' },
            { label: '--help, -h', description: 'Show this help text.' },
          ],
        },
        {
          title: 'Examples',
          lines: [
            'bun ./extract.js --mode automatic --browser chrome --profile Work',
            'bun ./extract.js --mode automatic --format md',
            'bun ./extract.js --mode network --debug --output-dir ./exports',
          ],
        },
      ],
    });
  }

  function parseCrawlerArgs(argv) {
    return parseSharedCliArgs(argv, {
      defaults: {
        browser: '',
        profile: '',
        outputName: '',
        outputDir: DEFAULT_OUTPUT_DIR,
        outputFormat: DEFAULT_TRANSCRIPT_OUTPUT_FORMAT,
        debugPort: null,
        debug: false,
        keepBrowserOpen: false,
        help: false,
        version: false,
        startUrl: DEFAULT_CRAWL_START_URL,
        stateFile: '',
        selectionSpec: '',
        waitBeforeDiscoveryMs: DEFAULT_WAIT_BEFORE_DISCOVERY_MS,
      },
      supportedBrowserKeys: SUPPORTED_BROWSER_KEYS,
      supportedOutputFormats: SUPPORTED_OUTPUT_FORMATS,
      normalizeBrowserKey,
      errorFactory: (message) => new CliError(message),
      extraOptions: [
        {
          flag: '--start-url',
          key: 'startUrl',
          normalize: (value) => String(value).trim(),
          validate: (value, _options, errorFactory) => {
            if (!value) {
              throw errorFactory
                ? errorFactory('Missing value for --start-url.')
                : new Error('Missing value for --start-url.');
            }
          },
        },
        {
          flag: '--state-file',
          key: 'stateFile',
          normalize: (value) => String(value).trim(),
        },
        {
          flag: '--select',
          key: 'selectionSpec',
          normalize: (value) => String(value).trim(),
        },
        {
          flag: '--wait-before-discovery-ms',
          key: 'waitBeforeDiscoveryMs',
          normalize: (value) => {
            const parsed = Number.parseInt(String(value).trim(), 10);
            if (!Number.isInteger(parsed) || parsed < 0) {
              throw new CliError(
                `Invalid wait duration "${value}" for --wait-before-discovery-ms.`,
              );
            }
            return parsed;
          },
        },
      ],
    });
  }

  function printCrawlerHelp() {
    printHelpScreen({
      name: 'Stream Transcript Extractor (Crawl Workflow)',
      summary:
        'Discover meeting pages from the Stream Meetings view, queue them in a persistent state file, and reuse one browser session to batch-extract transcripts with the automatic extractor.',
      usage: [
        'bun ./extract.js crawl [options]',
        './stream-transcript-extractor-<target> crawl [options]',
      ],
      sections: [
        {
          title: 'Recommended path',
          rows: [
            {
              label: 'bun ./extract.js crawl',
              description:
                'Open Stream home, wait the default 10-second settle window, switch to Meetings, scroll the rendered list, merge the results into *.state.json, let you select items from the queue, then extract each selected meeting with the automatic extractor in the same browser session.',
            },
          ],
        },
        {
          title: 'How crawl works',
          lines: [
            '1. Open Stream home and wait the default 10-second settle window for auth and page load.',
            '2. Select the Meetings view and scroll until the rendered list stops growing.',
            '3. Merge discovered recordings into the persistent *.state.json queue and let you select items in the terminal.',
            '4. Reuse the same browser debug session to run automatic extraction against each selected meeting, then write updated queue state and a *.batch.json run summary.',
          ],
        },
        {
          title: 'Selection syntax',
          lines: [
            'Use comma-separated indexes and ranges such as 1-5,8,10.',
            'Press Enter to select items that are not yet successful, or use keywords such as new, failed, done, or all.',
            'The crawler keeps a persistent *.state.json queue, keeps running after item-level failures, and also writes a *.batch.json status summary for the current run.',
          ],
        },
        {
          title: 'Recommended settings',
          rows: [
            {
              label: '--format md',
              description:
                'Recommended and default transcript output for the crawl workflow. Use json for automation or both when you need both forms.',
            },
            {
              label: '--debug off',
              description:
                'Recommended normal path. Turn on --debug only when diagnosing failed items or validating suspicious captures.',
            },
          ],
        },
        {
          title: 'Options',
          rows: [
            { label: '--browser <chrome|edge>', description: 'Use a specific browser.' },
            {
              label: '--profile <query>',
              description: 'Match a profile by name, email, or directory.',
            },
            {
              label: '--start-url <url>',
              description: `Override the Stream entry URL. Default: ${DEFAULT_CRAWL_START_URL}.`,
            },
            {
              label: '--state-file <path>',
              description: 'Override the persistent crawl queue state file.',
            },
            {
              label: '--select <spec>',
              description:
                'Select items non-interactively with pending, new, failed, done, all, or numeric ranges.',
            },
            {
              label: '--wait-before-discovery-ms <ms>',
              description:
                `Delay discovery after the Stream URL opens so auth and page load can settle. Default: ${DEFAULT_WAIT_BEFORE_DISCOVERY_MS}.`,
            },
            {
              label: '--output <name>',
              description:
                'Override the batch status filename prefix. Transcript filenames still use each meeting title.',
            },
            {
              label: '--output-dir <path>',
              description: 'Write transcripts and batch status files here.',
            },
            {
              label: '--format <json|md|both>',
              description:
                'Choose JSON, Markdown, or both transcript outputs. Recommended and default: md.',
            },
            {
              label: '--debug-port <port>',
              description: 'Force a specific remote-debugging port.',
            },
            {
              label: '--debug',
              description:
                'Save extended network diagnostics for each item. Recommended default: off.',
            },
            {
              label: '--keep-browser-open',
              description: 'Leave the launched browser open after the batch.',
            },
            { label: '--version, -v', description: 'Print the build version.' },
            { label: '--help, -h', description: 'Show this help text.' },
          ],
        },
        {
          title: 'Examples',
          lines: [
            'bun ./extract.js crawl --browser chrome --profile Work',
            'bun ./extract.js crawl --state-file ./exports/team.state.json',
            'bun ./extract.js crawl --select pending --browser edge',
            `bun ./extract.js crawl --wait-before-discovery-ms ${DEFAULT_WAIT_BEFORE_DISCOVERY_MS} --browser edge`,
            'bun ./extract.js crawl --output-dir ./exports --format md',
            'bun ./extract.js crawl --output-dir ./exports --format both --debug',
          ],
        },
      ],
    });
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

  function trimForTerminal(value, maxLength = 120) {
    const text = String(value || '').replace(/\s+/g, ' ').trim();

    if (!text || text.length <= maxLength) {
      return text;
    }

    return `${text.slice(0, maxLength - 3)}...`;
  }

  function parseSelectionSpec(value, itemCount) {
    return parseCliSelectionSpec(
      value,
      itemCount,
      (message) => new CliError(message),
    );
  }

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
      (message) => new CliError(message),
      config,
    );
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
    return connectCoreCdp(websocketUrl, {
      defaultTimeoutMs: DEFAULT_CDP_TIMEOUT_MS,
      errorFactory: (message) => new CliError(message),
    });
  }

  async function findPageTargets(port) {
    return findCorePageTargets({
      host: DEFAULT_CDP_HOST,
      port,
      errorFactory: (message) => new CliError(message),
    });
  }

  async function connectToPage(pageWebsocketUrl) {
    return connectCdp(pageWebsocketUrl);
  }

  async function evaluate(cdp, expression, timeoutMs = DEFAULT_CDP_TIMEOUT_MS) {
    return evaluateWithCdp(cdp, expression, {
      timeoutMs,
      errorFactory: (message) => new CliError(message),
    });
  }

  async function waitForBrowserDebugEndpoint(port) {
    return waitForCoreBrowserDebugEndpoint({
      host: DEFAULT_CDP_HOST,
      port,
      timeoutMs: DEFAULT_BROWSER_READY_TIMEOUT_MS,
      errorFactory: (message) => new CliError(message),
    });
  }

  async function waitForBrowserPageTarget(
    port,
    timeoutMs = DEFAULT_BROWSER_READY_TIMEOUT_MS,
  ) {
    return waitForCoreBrowserPageTarget({
      host: DEFAULT_CDP_HOST,
      port,
      timeoutMs,
      errorFactory: (message) => new CliError(message),
    });
  }

  async function waitForBrowserDebugEndpointToClose(
    port,
    timeoutMs = DEFAULT_BROWSER_SHUTDOWN_TIMEOUT_MS,
  ) {
    return waitForCoreBrowserDebugEndpointToClose({
      host: DEFAULT_CDP_HOST,
      port,
      timeoutMs,
    });
  }

  async function getBrowserDebuggerWebSocketUrl(port) {
    return getCoreBrowserDebuggerWebSocketUrl({
      host: DEFAULT_CDP_HOST,
      port,
    });
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
    return /stream|microsoftstream|m365\.cloud\.microsoft|substrate\.office\.com|officeapps\.live\.com|office\.com|office365\.com|office\.net|sharepoint\.com|teams\.microsoft\.com|graph\.microsoft\.com|meetingcatchupportal|cortana\.ai/.test(
      String(url || '').toLowerCase(),
    );
  }

  function containsTranscriptSignal(value) {
    return /\btranscript(?:s)?\b|\bcaptions?\b|\bsubtitles?\b|\butterances?\b|\bspeakers?\b|\bspeech\b|\bvtt\b|\bclosedcaption(?:s)?\b/.test(
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
    if (url.startsWith('data:')) {
      return false;
    }

    const mimeType = String(response?.mimeType || '').toLowerCase();
    const headers = lowerCaseHeaderMap(response?.headers);
    const contentType = String(headers['content-type'] || '').toLowerCase();
    const contentDisposition = String(headers['content-disposition'] || '').toLowerCase();
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
    if (String(url || '').startsWith('data:')) {
      return 0;
    }

    let score = 0;
    const haystack = [url, mimeType, contentType, contentDisposition].join(' ');

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

  function isUsableTranscriptCandidate(candidate) {
    const entryCount = Number(candidate?.parsedEntryCount || 0);
    if (entryCount <= 0) {
      return false;
    }

    const entries = Array.isArray(candidate?.parsedEntries)
      ? candidate.parsedEntries
      : [];
    const speakerCount = entries.filter((entry) =>
      normalizeInlineText(entry?.speaker || ''),
    ).length;
    const timestampCount = entries.filter((entry) =>
      normalizeInlineText(entry?.timestamp || ''),
    ).length;
    const signalText = [
      candidate?.url || '',
      candidate?.mimeType || '',
      candidate?.contentType || '',
      candidate?.parsedPath || '',
    ]
      .join(' ')
      .toLowerCase();

    if (candidate?.parsedFormat === 'vtt' || isEncryptedTranscriptUrl(candidate?.url || '')) {
      return true;
    }

    if (containsTranscriptSignal(signalText)) {
      return speakerCount > 0 || timestampCount > 0 || entryCount >= 8;
    }

    return speakerCount >= 2 && timestampCount >= 1;
  }

  function scoreTranscriptCandidate(candidate) {
    const entries = Array.isArray(candidate?.parsedEntries)
      ? candidate.parsedEntries
      : [];
    const speakerCount = entries.filter((entry) =>
      normalizeInlineText(entry?.speaker || ''),
    ).length;
    const timestampCount = entries.filter((entry) =>
      normalizeInlineText(entry?.timestamp || ''),
    ).length;
    const signalText = [
      candidate?.url || '',
      candidate?.mimeType || '',
      candidate?.contentType || '',
      candidate?.parsedPath || '',
    ]
      .join(' ')
      .toLowerCase();

    let score = Number(candidate?.score || 0);
    score += Number(candidate?.parsedEntryCount || 0) * 3;
    score += speakerCount * 2;
    score += timestampCount * 2;

    if (containsTranscriptSignal(signalText)) {
      score += 20;
    }
    if (candidate?.parsedFormat === 'vtt') {
      score += 15;
    }
    if (isEncryptedTranscriptUrl(candidate?.url || '')) {
      score += 15;
    }

    return score;
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

  function saveBatchStatusOutput(payload, outputBasePath) {
    const outputPath = `${outputBasePath}.batch.json`;
    writeFileSync(outputPath, JSON.stringify(payload, null, 2));
    return outputPath;
  }

  function resolveCrawlerStatePath(options) {
    if (options.stateFile) {
      return resolve(options.stateFile);
    }

    return `${buildOutputBasePath('crawl', options.outputName, options.outputDir)}.state.json`;
  }

  function saveCrawlerStateOutput(payload, stateFilePath) {
    mkdirSync(dirname(stateFilePath), { recursive: true });
    writeFileSync(stateFilePath, JSON.stringify(payload, null, 2));
    return stateFilePath;
  }

  function saveCrawlerDiscoveryDebugOutput(payload, stateFilePath) {
    const outputPath = stateFilePath.endsWith('.state.json')
      ? stateFilePath.replace(/\.state\.json$/, '.discovery.debug.json')
      : `${stateFilePath}.discovery.debug.json`;
    mkdirSync(dirname(outputPath), { recursive: true });
    writeFileSync(outputPath, JSON.stringify(payload, null, 2));
    return outputPath;
  }

  function loadCrawlerState(stateFilePath) {
    if (!existsSync(stateFilePath)) {
      return {
        stateVersion: CRAWL_STATE_VERSION,
        items: [],
      };
    }

    try {
      const payload = JSON.parse(readFileSync(stateFilePath, 'utf-8'));
      return {
        stateVersion:
          typeof payload?.stateVersion === 'number'
            ? payload.stateVersion
            : CRAWL_STATE_VERSION,
        items: Array.isArray(payload?.items) ? payload.items : [],
      };
    } catch (error) {
      throw new CliError(
        `Unable to read crawl state file "${stateFilePath}": ` +
          `${error instanceof Error ? error.message : 'Unknown parse failure.'}`,
      );
    }
  }

  function uniqueStrings(values) {
    return [...new Set((values || []).map((value) => String(value || '').trim()).filter(Boolean))];
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

  function normalizeCrawlerStateItem(item, fallbackDiscoveredAt = '') {
    const candidateUrls = sortMeetingUrlsForExtraction(
      getCrawlerItemLookupUrls(item),
    );
    const url = candidateUrls[0] || normalizeMeetingUrl(item?.url);

    return {
      url,
      identityKey: normalizeInlineText(item?.identityKey || ''),
      title: String(item?.title || ''),
      subtitle: String(item?.subtitle || ''),
      candidateUrls,
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
        name: 'Stream Transcript Extractor',
        version: BUILD_VERSION,
        buildTime: BUILD_TIME,
        platform: CURRENT_PLATFORM,
      },
      stateVersion: CRAWL_STATE_VERSION,
      updatedAt,
      stateFilePath,
      options: {
        startUrl: options.startUrl,
        outputDir: resolveOutputDirectory(options.outputDir),
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
      url === DEFAULT_CRAWL_START_URL.toLowerCase() ||
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

  function buildCrawlerExtractionTargetUrls(item) {
    return sortMeetingUrlsForExtraction([
      ...(Array.isArray(item?.candidateUrls) ? item.candidateUrls : []),
      item?.url || '',
    ]);
  }

  async function navigatePageAndWait(
    cdp,
    url,
    timeoutMs = DEFAULT_PAGE_NAVIGATION_TIMEOUT_MS,
  ) {
    await navigateWithCdp(cdp, url, {
      timeoutMs,
      settleMs: DEFAULT_CRAWL_SCROLL_SETTLE_MS,
      errorFactory: (message) => new CliError(message),
    });
    return getCurrentPageSnapshot(cdp).catch(() => ({
      title: '',
      url,
    }));
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
    await sleep(DEFAULT_CAPTURE_SETTLE_MS);
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
    return loadCapturedResponseBody(cdp, requestId, bodyCache, {
      trimForPreview,
    });
  }

  function createEmptyResponseBodyRecord(bodyError = '') {
    return createResponseBodyRecord(
      {
        body: '',
        bodyError,
      },
      { trimForPreview },
    );
  }

  function normalizeMeetingUrl(value) {
    const text = String(value || '').trim();
    if (!text) {
      return '';
    }

    if (!/^(https?:)?\/\//i.test(text) && !text.startsWith('/')) {
      return '';
    }

    try {
      const url = new URL(text, DEFAULT_CRAWL_START_URL);
      if (!/^https?:$/.test(url.protocol)) {
        return '';
      }

      url.hash = '';
      const sortedSearchParams = [...url.searchParams.entries()].sort(
        ([leftKey, leftValue], [rightKey, rightValue]) =>
          leftKey.localeCompare(rightKey) || leftValue.localeCompare(rightValue),
      );
      url.search = '';
      for (const [key, paramValue] of sortedSearchParams) {
        url.searchParams.append(key, paramValue);
      }

      return url.href;
    } catch {
      return '';
    }
  }

  function buildMeetingExtractionUrl(value) {
    const normalizedUrl = normalizeMeetingUrl(value);
    if (!normalizedUrl) {
      return '';
    }

    try {
      const url = new URL(normalizedUrl, DEFAULT_CRAWL_START_URL);
      const lowerHost = url.hostname.toLowerCase();
      const lowerPath = url.pathname.toLowerCase();
      const isSharePointVideoFile =
        /\.sharepoint\.com$/.test(lowerHost) &&
        /\.(mp4|m4v|mov|webm)$/.test(lowerPath) &&
        !/\/_layouts\/15\/stream\.aspx$/.test(lowerPath);

      if (!isSharePointVideoFile) {
        return normalizedUrl;
      }

      if (String(url.searchParams.get('web') || '').trim() !== '1') {
        url.searchParams.set('web', '1');
      }

      return normalizeMeetingUrl(url.toString());
    } catch {
      return normalizedUrl;
    }
  }

  function buildMeetingExtractionComparisonKey(value) {
    const executionUrl = buildMeetingExtractionUrl(value);
    if (!executionUrl) {
      return '';
    }

    try {
      const url = new URL(executionUrl, DEFAULT_CRAWL_START_URL);
      const lowerOrigin = url.origin.toLowerCase();
      const lowerPath = url.pathname.toLowerCase();

      if (
        /\.sharepoint\.com$/.test(url.hostname.toLowerCase()) &&
        /\.(mp4|m4v|mov|webm)$/.test(lowerPath)
      ) {
        return `${lowerOrigin}${lowerPath}`;
      }

      if (
        /\.sharepoint\.com$/.test(url.hostname.toLowerCase()) &&
        /\/_layouts\/15\/stream\.aspx$/.test(lowerPath)
      ) {
        const streamId = String(url.searchParams.get('id') || '').trim();
        if (streamId) {
          return `${lowerOrigin}${decodeURIComponent(streamId).toLowerCase()}`;
        }
      }

      return executionUrl.toLowerCase();
    } catch {
      return executionUrl.toLowerCase();
    }
  }

  function countMeetingUrlSearchParams(value) {
    const executionUrl = buildMeetingExtractionUrl(value);
    if (!executionUrl) {
      return 0;
    }

    try {
      return [...new URL(executionUrl, DEFAULT_CRAWL_START_URL).searchParams.keys()]
        .length;
    } catch {
      return 0;
    }
  }

  function isLikelyMeetingItemUrl(value) {
    const normalizedUrl = normalizeMeetingUrl(value);
    if (!normalizedUrl) {
      return false;
    }

    const lowerUrl = normalizedUrl.toLowerCase();
    if (
      lowerUrl === DEFAULT_CRAWL_START_URL.toLowerCase() ||
      /[?&]home=1(?:&|$)/.test(lowerUrl)
    ) {
      return false;
    }

    if (
      !/(stream|sharepoint|m365\.cloud\.microsoft|meetingcatchupportal|cortana\.ai)/.test(
        lowerUrl,
      )
    ) {
      return false;
    }

    if (/meetingcatchupportal|meetingcatchup|stream\.aspx|watch|oneplayer|clip/.test(lowerUrl)) {
      return true;
    }

    return /\/recordings?\/.*\.(mp4|m4v|mov|webm)(?:$|[?#])/.test(lowerUrl);
  }

  function isLikelyMeetingVideoFileUrl(value) {
    const normalizedUrl = normalizeMeetingUrl(value);
    if (!normalizedUrl) {
      return false;
    }

    const lowerUrl = normalizedUrl.toLowerCase();
    if (
      !/(stream|sharepoint|m365\.cloud\.microsoft|meetingcatchupportal|cortana\.ai)/.test(
        lowerUrl,
      )
    ) {
      return false;
    }

    return /\.(mp4|m4v|mov|webm)(?:$|[?#])/.test(lowerUrl);
  }

  function scoreMeetingTargetUrl(value) {
    const lowerUrl = normalizeMeetingUrl(value).toLowerCase();
    if (!lowerUrl) {
      return 0;
    }

    if (/meetingcatchupportal|meetingcatchup/.test(lowerUrl)) {
      return 3;
    }

    if (/\/recordings?\/.*\.(mp4|m4v|mov|webm)(?:$|[?#])/.test(lowerUrl)) {
      return 2;
    }

    if (/stream\.aspx|watch|oneplayer|clip/.test(lowerUrl)) {
      return 1;
    }

    return 0;
  }

  function scoreMeetingExtractionTargetUrl(value) {
    const lowerUrl = buildMeetingExtractionUrl(value).toLowerCase();
    if (!lowerUrl) {
      return 0;
    }

    if (/stream\.aspx|watch|oneplayer|clip/.test(lowerUrl)) {
      return 4;
    }

    if (isLikelyMeetingVideoFileUrl(lowerUrl)) {
      return 3;
    }

    if (/meetingcatchupportal|meetingcatchup/.test(lowerUrl)) {
      return 2;
    }

    return isLikelyMeetingItemUrl(lowerUrl) ? 1 : 0;
  }

  function uniqueNormalizedMeetingUrls(values) {
    return [
      ...new Set(
        (values || [])
          .map((value) => normalizeMeetingUrl(value))
          .filter(Boolean),
      ),
    ];
  }

  function sortMeetingUrlsForExtraction(values) {
    const sortedUrls = (values || [])
      .map((value) => buildMeetingExtractionUrl(value))
      .filter((value) => scoreMeetingExtractionTargetUrl(value) > 0)
      .filter(Boolean)
      .sort((left, right) => {
      const scoreDelta =
        scoreMeetingExtractionTargetUrl(right) -
        scoreMeetingExtractionTargetUrl(left);
      if (scoreDelta !== 0) {
        return scoreDelta;
      }

      const searchParamDelta =
        countMeetingUrlSearchParams(right) - countMeetingUrlSearchParams(left);
      if (searchParamDelta !== 0) {
        return searchParamDelta;
      }

      return scoreMeetingTargetUrl(right) - scoreMeetingTargetUrl(left);
    });

    const dedupedUrls = [];
    const seenComparisonKeys = new Set();

    for (const value of sortedUrls) {
      const comparisonKey = buildMeetingExtractionComparisonKey(value) || value;
      if (seenComparisonKeys.has(comparisonKey)) {
        continue;
      }

      seenComparisonKeys.add(comparisonKey);
      dedupedUrls.push(value);
    }

    return dedupedUrls;
  }

  function extractUrlsFromText(value) {
    const text = String(value || '');
    const matches = text.match(/https?:\/\/[^\s"'<>\\]+/g) || [];

    return matches
      .map((match) => match.split(/(?:&quot;|&#34;|&apos;|&#39;)/i, 1)[0])
      .map((match) => match.replace(/[),.;]+$/g, ''))
      .filter(Boolean);
  }

  function extractMeetingDiscoveryItemsFromValue(value, sourceUrl = '') {
    const matches = [];
    const seenMatches = new Set();

    function addMatch(candidateUrl, metadata, path) {
      const normalizedUrl = normalizeMeetingUrl(candidateUrl);
      if (!isLikelyMeetingItemUrl(normalizedUrl)) {
        return;
      }

      const key = `${normalizedUrl}|${path}`;
      if (seenMatches.has(key)) {
        return;
      }

      seenMatches.add(key);
      matches.push({
        url: normalizedUrl,
        title: normalizeInlineText(metadata?.title || ''),
        subtitle: normalizeInlineText(metadata?.subtitle || ''),
        identityKey: normalizeInlineText(metadata?.identityKey || ''),
        candidateUrls: [normalizedUrl],
        sourceUrl,
        path,
      });
    }

    function walk(current, path, depth = 0) {
      if (depth > 10 || current == null) {
        return;
      }

      if (typeof current === 'string') {
        for (const candidateUrl of extractUrlsFromText(current)) {
          addMatch(candidateUrl, {}, path);
        }
        return;
      }

      if (Array.isArray(current)) {
        current.forEach((entry, index) => {
          walk(entry, `${path}[${index}]`, depth + 1);
        });
        return;
      }

      if (typeof current !== 'object') {
        return;
      }

      const title = getFirstStringValue(current, [
        'title',
        'name',
        'displayName',
        'meetingTitle',
        'subject',
        'fileName',
        'label',
      ]);
      const subtitle = getFirstStringValue(current, [
        'description',
        'sharedBy',
        'sharedByName',
        'createdBy',
        'createdByName',
        'ownerName',
      ]);
      const extension = getFirstStringValue(current, [
        'extension',
        'fileExtension',
        'file_extension',
      ]).toLowerCase();
      const itemType = getFirstStringValue(current, ['type']).toLowerCase();
      const fileType = getFirstStringValue(current, [
        'fileType',
        'file_type',
        'app',
      ]).toLowerCase();
      const isMeetingRecording =
        current?.is_meeting_recording === true ||
        current?.isMeetingRecording === true ||
        String(
          current?.is_meeting_recording ?? current?.isMeetingRecording ?? '',
        ).toLowerCase() === 'true';
      const looksLikeVideoFile =
        ['mp4', 'm4v', 'mov', 'webm'].includes(extension) ||
        itemType === 'video' ||
        fileType === 'stream';
      const identityKey =
        normalizeInlineText(current?.resource_id || current?.resourceId || '') ||
        normalizeInlineText(current?.doc_id || current?.docId || '') ||
        normalizeInlineText(current?.file_id || current?.fileId || '') ||
        normalizeInlineText(current?.sharepoint_info?.unique_id || '') ||
        normalizeInlineText(current?.sharepointIds?.listItemUniqueId || '') ||
        normalizeInlineText(current?.sharepointIds?.driveItemId || '') ||
        normalizeInlineText(current?.sharepointIds?.itemId || '') ||
        normalizeInlineText(current?.onedrive_info?.item_id || '');

      if (isMeetingRecording || looksLikeVideoFile) {
        const candidateKeys = [
          'meeting_catchup_link',
          'meetingCatchupLink',
          'canonicalUrl',
          'canonical_url',
          'webUrl',
          'web_url',
          'url',
          'mru_url',
          'shareUrl',
          'share_url',
          'playerUrl',
          'player_url',
          'downloadUrl',
          'download_url',
        ];
        candidateKeys.forEach((key) => {
          if (typeof current[key] === 'string') {
            addMatch(
              current[key],
              { title, subtitle, identityKey },
              `${path}.${key}`,
            );
          }
        });
        return;
      }

      for (const [key, child] of Object.entries(current)) {
        const childPath = `${path}.${key}`;

        if (typeof child === 'string') {
          if (
            /url|uri|href|link|path|weburl|shareurl|watchurl|playerurl/i.test(key) ||
            isLikelyMeetingItemUrl(child)
          ) {
            addMatch(child, { title, subtitle }, childPath);
          } else {
            for (const candidateUrl of extractUrlsFromText(child)) {
              addMatch(candidateUrl, { title, subtitle }, childPath);
            }
          }
          continue;
        }

        walk(child, childPath, depth + 1);
      }
    }

    walk(value, '$');
    return matches;
  }

  function extractMeetingDiscoveryItemsFromBody(body, sourceUrl = '') {
    const trimmed = stripUtf8Bom(String(body || '')).trim();
    if (!trimmed) {
      return [];
    }

    const jsonValue = tryParseJsonLenient(trimmed);
    if (jsonValue != null) {
      return extractMeetingDiscoveryItemsFromValue(jsonValue, sourceUrl);
    }

    return extractUrlsFromText(trimmed)
      .map((candidateUrl) => ({
        url: normalizeMeetingUrl(candidateUrl),
        title: '',
        subtitle: '',
        sourceUrl,
        path: '$',
      }))
      .filter((item) => isLikelyMeetingItemUrl(item.url));
  }

  function normalizeMeetingComparisonTitle(value) {
    return normalizeInlineText(
      String(value || '')
        .replace(/^stream\s+/i, '')
        .replace(
          /-\d{8}_\d{6}-meeting\s+(?:recording|transcript)\.(?:mp4|m4v|mov|webm)$/i,
          '',
        )
        .replace(/\.(?:mp4|m4v|mov|webm)$/i, '')
        .replace(/\{placeholder\}/gi, 'placeholder')
        .replace(/[|/]+/g, ' ')
        .replace(/[^a-z0-9]+/gi, ' '),
    ).toLowerCase();
  }

  function buildMeetingComparisonKey(value) {
    return normalizeMeetingComparisonTitle(value).replace(/\s+/g, '');
  }

  function parseGraphThumbnailDriveItem(value) {
    const normalizedUrl = normalizeMeetingUrl(value);
    if (!normalizedUrl) {
      return null;
    }

    const match = normalizedUrl.match(
      /^https:\/\/graph\.microsoft\.com\/v1\.0\/drives\/([^/]+)\/items\/([^/]+)\/thumbnails\//i,
    );
    if (!match) {
      return null;
    }

    return {
      driveId: decodeURIComponent(match[1]),
      itemId: decodeURIComponent(match[2]),
      thumbnailUrl: normalizedUrl,
    };
  }

  function getGraphAuthorizationHeader(graphThumbnailCandidates) {
    if (
      !(graphThumbnailCandidates instanceof Map) ||
      graphThumbnailCandidates.size === 0
    ) {
      return '';
    }

    return (
      [...new Set(
        [...graphThumbnailCandidates.values()]
          .map((candidate) => String(candidate?.authorization || '').trim())
          .filter(Boolean),
      )][0] || ''
    );
  }

  function buildGraphItemMetadataUrl(driveId, itemId) {
    const metadataUrl = new URL(
      `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}`,
    );
    metadataUrl.searchParams.set(
      '$select',
      'id,name,webUrl,file,video,sharepointIds,parentReference,@microsoft.graph.downloadUrl',
    );
    return metadataUrl.toString();
  }

  async function fetchGraphItemMetadata(driveId, itemId, authorization) {
    if (!driveId || !itemId || !authorization) {
      return null;
    }

    let response;
    try {
      response = await fetch(buildGraphItemMetadataUrl(driveId, itemId), {
        headers: {
          Authorization: authorization,
          Accept: 'application/json',
        },
      });
    } catch {
      return null;
    }

    if (!response.ok) {
      return null;
    }

    try {
      return await response.json();
    } catch {
      return null;
    }
  }

  async function resolveGraphThumbnailMeetingItems(
    graphThumbnailCandidates,
    {
      visibleTitles = [],
      existingNetworkItems = [],
    } = {},
  ) {
    if (!(graphThumbnailCandidates instanceof Map) || graphThumbnailCandidates.size === 0) {
      return [];
    }

    const authorization = getGraphAuthorizationHeader(graphThumbnailCandidates);
    if (!authorization) {
      return [];
    }

    const visibleQueueByKey = new Map();
    for (const title of visibleTitles) {
      const key = buildMeetingComparisonKey(title);
      if (!key) {
        continue;
      }

      if (!visibleQueueByKey.has(key)) {
        visibleQueueByKey.set(key, []);
      }
      visibleQueueByKey.get(key).push(normalizeInlineText(title));
    }

    const existingNetworkCounts = new Map();
    for (const item of existingNetworkItems) {
      const key = buildMeetingComparisonKey(item?.title || '');
      if (!key) {
        continue;
      }

      existingNetworkCounts.set(
        key,
        (existingNetworkCounts.get(key) || 0) + 1,
      );
    }

    const outstandingVisibleByKey = new Map();
    for (const [key, titles] of visibleQueueByKey.entries()) {
      const alreadyCovered = existingNetworkCounts.get(key) || 0;
      const remainingTitles = titles.slice(alreadyCovered);
      if (remainingTitles.length > 0) {
        outstandingVisibleByKey.set(key, remainingTitles);
      }
    }

    if (outstandingVisibleByKey.size === 0) {
      return [];
    }

    const resolvedGraphItems = [];
    for (const candidate of graphThumbnailCandidates.values()) {
      const payload = await fetchGraphItemMetadata(
        candidate.driveId,
        candidate.itemId,
        authorization,
      );
      if (!payload) {
        continue;
      }

      const itemName = normalizeInlineText(payload?.name || '');
      const key = buildMeetingComparisonKey(itemName);
      if (!key || !outstandingVisibleByKey.has(key)) {
        continue;
      }

      const queue = outstandingVisibleByKey.get(key);
      const title = queue.shift() || itemName;
      if (queue.length === 0) {
        outstandingVisibleByKey.delete(key);
      }

      const targetUrl =
        normalizeMeetingUrl(
          payload?.webUrl ||
            payload?.['@microsoft.graph.downloadUrl'] ||
            '',
        ) || '';
      if (!targetUrl || !isLikelyMeetingItemUrl(targetUrl)) {
        continue;
      }

      resolvedGraphItems.push({
        url: targetUrl,
        title,
        subtitle: normalizeInlineText(
          payload?.parentReference?.path || '',
        ),
        identityKey:
          normalizeInlineText(payload?.sharepointIds?.listItemUniqueId || '') ||
          normalizeInlineText(payload?.id || ''),
        sourceUrl: normalizeMeetingUrl(candidate.thumbnailUrl),
        path: '$graphThumbnail',
      });
    }

    return resolvedGraphItems;
  }

  async function resolveDomGraphMeetingItems(
    domItems,
    {
      graphThumbnailCandidates,
      existingNetworkItems = [],
    } = {},
  ) {
    if (!Array.isArray(domItems) || domItems.length === 0) {
      return [];
    }

    const authorization = getGraphAuthorizationHeader(graphThumbnailCandidates);
    if (!authorization) {
      return [];
    }

    const existingNetworkUrls = new Set(
      existingNetworkItems
        .map((item) => normalizeMeetingUrl(item?.url || ''))
        .filter(Boolean),
    );
    const existingNetworkIdentityKeys = new Set(
      existingNetworkItems
        .map((item) => normalizeInlineText(item?.identityKey || ''))
        .filter(Boolean),
    );
    const requestedGraphItems = new Set();
    const resolvedGraphItems = [];

    for (const item of domItems) {
      const existingUrl = normalizeMeetingUrl(item?.url || '');
      const identityKey = normalizeInlineText(item?.identityKey || '');
      const driveId = normalizeInlineText(item?.graphDriveId || '');
      const itemId = normalizeInlineText(item?.graphItemId || '');
      if (!driveId || !itemId) {
        continue;
      }

      if (
        (existingUrl && existingNetworkUrls.has(existingUrl)) ||
        (identityKey && existingNetworkIdentityKeys.has(identityKey))
      ) {
        continue;
      }

      const requestKey = `${driveId}\t${itemId}`;
      if (requestedGraphItems.has(requestKey)) {
        continue;
      }
      requestedGraphItems.add(requestKey);

      const payload = await fetchGraphItemMetadata(
        driveId,
        itemId,
        authorization,
      );
      if (!payload) {
        continue;
      }

      const resolvedUrl =
        normalizeMeetingUrl(
          payload?.webUrl || payload?.['@microsoft.graph.downloadUrl'] || '',
        ) || existingUrl;
      if (
        !resolvedUrl ||
        (!isLikelyMeetingItemUrl(resolvedUrl) &&
          !isLikelyMeetingVideoFileUrl(resolvedUrl))
      ) {
        continue;
      }

      resolvedGraphItems.push({
        url: resolvedUrl,
        title: normalizeInlineText(item?.title || '') || normalizeInlineText(payload?.name || '') || resolvedUrl,
        subtitle:
          normalizeInlineText(item?.subtitle || '') ||
          normalizeInlineText(payload?.parentReference?.path || ''),
        identityKey:
          identityKey ||
          normalizeInlineText(payload?.sharepointIds?.listItemUniqueId || '') ||
          normalizeInlineText(payload?.id || ''),
        sourceUrl: buildGraphItemMetadataUrl(driveId, itemId),
        path: '$graphDomItem',
      });
      existingNetworkUrls.add(resolvedUrl);
      if (identityKey) {
        existingNetworkIdentityKeys.add(identityKey);
      }
    }

    return resolvedGraphItems;
  }

  function mergeDiscoveredMeetingItems(networkItems = [], domItems = []) {
    const mergedItems = new Map();

    function buildCandidateUrls(item) {
      return uniqueNormalizedMeetingUrls([
        item?.url || '',
        ...(Array.isArray(item?.candidateUrls) ? item.candidateUrls : []),
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
          sources: item.sources,
          discoveryPaths: item.discoveryPaths,
          discoverySourceUrls: item.discoverySourceUrls,
        };
      });
  }

  function buildMeetingDiscoveryExpression() {
    const settleMs = JSON.stringify(DEFAULT_CRAWL_SCROLL_SETTLE_MS);

    return `
      (async () => {
        const settleMs = ${settleMs};

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
            href === ${JSON.stringify(DEFAULT_CRAWL_START_URL.toLowerCase())} ||
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

  function shouldTrackMeetingDiscoveryResponse(response, resourceType) {
    const url = String(response?.url || '').toLowerCase();
    if (!url || !isStreamRelatedUrl(url)) {
      return false;
    }

    const resource = String(resourceType || '');
    if (/Image|Media|Font|Stylesheet/.test(resource)) {
      return false;
    }

    const mimeType = String(response?.mimeType || '').toLowerCase();
    const contentType = String(
      lowerCaseHeaderMap(response?.headers || {})['content-type'] || '',
    ).toLowerCase();
    const contentHint = `${mimeType} ${contentType}`;

    return /json|text|html|xml|javascript/.test(contentHint);
  }

  async function discoverMeetingPages(
    cdp,
    startUrl,
    {
      navigate = true,
      initialPage = null,
      waitBeforeDomDiscoveryMs = 0,
      debugEnabled = false,
    } = {},
  ) {
    await cdp.send('Network.enable', {
      maxTotalBufferSize: 100_000_000,
      maxResourceBufferSize: 10_000_000,
    });
    await cdp
      .send('Network.setCacheDisabled', { cacheDisabled: true })
      .catch(() => {});
    await cdp
      .send('Network.setBypassServiceWorker', { bypass: true })
      .catch(() => {});

    const eventScope = createCdpEventScope(cdp);
    const trackedResponses = new Map();
    const finishedRequests = new Set();
    const failedRequests = new Set();
    const bodyCache = new Map();
    const requestMetadata = new Map();
    const graphThumbnailCandidates = new Map();

    eventScope.on('Network.requestWillBeSent', (params) => {
      requestMetadata.set(String(params.requestId), {
        requestId: String(params.requestId),
        url: String(params.request?.url || ''),
        method: String(params.request?.method || ''),
      });
    });

    eventScope.on('Network.responseReceived', (params) => {
      if (
        trackedResponses.size >= MAX_DEBUG_RESPONSES ||
        !shouldTrackMeetingDiscoveryResponse(params.response, params.type)
      ) {
        return;
      }

      const requestId = String(params.requestId);
      trackedResponses.set(requestId, {
        requestId,
        url: String(params.response.url || ''),
        status: params.response.status,
        mimeType: String(params.response.mimeType || ''),
        resourceType: String(params.type || ''),
      });
    });

    eventScope.on('Network.requestWillBeSentExtraInfo', (params) => {
      const requestId = String(params.requestId);
      const request = requestMetadata.get(requestId) || {};
      const graphTarget = parseGraphThumbnailDriveItem(request.url || '');
      if (!graphTarget || String(request.method || '').toUpperCase() !== 'GET') {
        return;
      }

      const headers = params.headers || {};
      const authorization = String(
        headers.authorization || headers.Authorization || '',
      ).trim();
      const key = `${graphTarget.driveId}\t${graphTarget.itemId}`;
      if (!graphThumbnailCandidates.has(key)) {
        graphThumbnailCandidates.set(key, {
          ...graphTarget,
          authorization,
        });
        return;
      }

      const existing = graphThumbnailCandidates.get(key);
      if (!existing.authorization && authorization) {
        existing.authorization = authorization;
      }
    });

    eventScope.on('Network.loadingFinished', (params) => {
      finishedRequests.add(String(params.requestId));
    });

    eventScope.on('Network.loadingFailed', (params) => {
      failedRequests.add(String(params.requestId));
    });

    try {
      const pageSnapshot = navigate
        ? await navigatePageAndWait(cdp, startUrl)
        : initialPage ||
          (await getCurrentPageSnapshot(cdp).catch(() => ({
            title: '',
            url: startUrl,
          })));

      if (waitBeforeDomDiscoveryMs > 0) {
        await sleep(waitBeforeDomDiscoveryMs);
      }

      const domDiscovery = await evaluate(
        cdp,
        buildMeetingDiscoveryExpression(),
        DEFAULT_PAGE_NAVIGATION_TIMEOUT_MS * 2,
      );

      if (domDiscovery?.error) {
        throw new CliError(domDiscovery.error);
      }

      eventScope.stopAll();

      const networkItems = [];
      const debugResponseBodies = [];
      const bodyTargets = Array.from(trackedResponses.values())
        .filter(
          (responseRecord) =>
            finishedRequests.has(responseRecord.requestId) &&
            !failedRequests.has(responseRecord.requestId),
        )
        .slice(0, MAX_DISCOVERY_RESPONSE_BODIES);

      for (const responseRecord of bodyTargets) {
        const bodyRecord = await loadResponseBody(
          cdp,
          responseRecord.requestId,
          bodyCache,
        );

        if (bodyRecord.bodyError || !bodyRecord.body) {
          if (debugEnabled) {
            debugResponseBodies.push({
              ...responseRecord,
              bodyError: bodyRecord.bodyError,
              bodyPreview: bodyRecord.bodyPreview,
              matchedItems: [],
            });
          }
          continue;
        }

        const extractedItems = extractMeetingDiscoveryItemsFromBody(
          bodyRecord.body,
          responseRecord.url,
        );
        networkItems.push(...extractedItems);

        if (debugEnabled) {
          debugResponseBodies.push({
            ...responseRecord,
            bodyError: bodyRecord.bodyError,
            bodyPreview: bodyRecord.bodyPreview,
            matchedItems: extractedItems,
          });
        }
      }

      const graphResolvedItems = await resolveGraphThumbnailMeetingItems(
        graphThumbnailCandidates,
        {
          visibleTitles: Array.isArray(domDiscovery?.visibleCardTitles)
            ? domDiscovery.visibleCardTitles
            : [],
          existingNetworkItems: networkItems,
        },
      );
      networkItems.push(...graphResolvedItems);

      const domGraphResolvedItems = await resolveDomGraphMeetingItems(
        Array.isArray(domDiscovery?.items) ? domDiscovery.items : [],
        {
          graphThumbnailCandidates,
          existingNetworkItems: networkItems,
        },
      );
      networkItems.push(...domGraphResolvedItems);

      const mergedItems = mergeDiscoveredMeetingItems(
        networkItems,
        Array.isArray(domDiscovery?.items) ? domDiscovery.items : [],
      );

      const totalGraphResolvedItemCount =
        graphResolvedItems.length + domGraphResolvedItems.length;

      return {
        startUrl,
        page: domDiscovery?.page || pageSnapshot,
        meetingsControlLabel: String(domDiscovery?.controlLabel || ''),
        meetingsSelectedBeforeClick: Boolean(domDiscovery?.selectedBeforeClick),
        scrollIterations:
          typeof domDiscovery?.scrollIterations === 'number'
            ? domDiscovery.scrollIterations
            : 0,
        trackedResponseCount: trackedResponses.size,
        networkItemCount: networkItems.length,
        graphResolvedItemCount: totalGraphResolvedItemCount,
        graphThumbnailResolvedItemCount: graphResolvedItems.length,
        domGraphResolvedItemCount: domGraphResolvedItems.length,
        domItemCount: Array.isArray(domDiscovery?.items)
          ? domDiscovery.items.length
          : 0,
        visibleCardCount:
          typeof domDiscovery?.visibleCardCount === 'number'
            ? domDiscovery.visibleCardCount
            : Array.isArray(domDiscovery?.visibleCardTitles)
              ? domDiscovery.visibleCardTitles.length
              : 0,
        visibleCardTitles: Array.isArray(domDiscovery?.visibleCardTitles)
          ? domDiscovery.visibleCardTitles
          : [],
        items: mergedItems,
        debug: debugEnabled
          ? {
              trackedResponses: Array.from(trackedResponses.values()),
              responseBodies: debugResponseBodies,
              graphThumbnailCandidateCount: graphThumbnailCandidates.size,
              graphResolvedItemCount: totalGraphResolvedItemCount,
              graphThumbnailResolvedItemCount: graphResolvedItems.length,
              domGraphResolvedItemCount: domGraphResolvedItems.length,
            }
          : null,
      };
    } finally {
      eventScope.stopAll();
      await cdp
        .send('Network.setBypassServiceWorker', { bypass: false })
        .catch(() => {});
      await cdp
        .send('Network.setCacheDisabled', { cacheDisabled: false })
        .catch(() => {});
      await cdp.send('Network.disable').catch(() => {});
    }
  }

  async function refetchBinaryBody(url) {
    const response = await fetch(url, {
      redirect: 'follow',
    });

    if (!response.ok) {
      throw new Error(`Refetch failed with status ${response.status}.`);
    }

    const rawBytes = Buffer.from(await response.arrayBuffer());
    return createResponseBodyRecord(
      {
        body: rawBytes.toString('utf8'),
        bodyError: '',
        base64Encoded: true,
        rawBodyBase64: rawBytes.toString('base64'),
        rawBodyByteLength: rawBytes.length,
        rawBodyPreviewHex: rawBytes.subarray(0, 32).toString('hex'),
        refetchedExternally: true,
        refetchError: '',
      },
      { trimForPreview },
    );
  }

  async function reloadPageAndWait(cdp, timeoutMs = DEFAULT_CDP_TIMEOUT_MS) {
    return reloadWithCdp(cdp, {
      timeoutMs,
      settleMs: 1_000,
      errorFactory: (message) => new CliError(message),
    });
  }

  async function captureStreamNetwork(
    cdp,
    prompt,
    debugEnabled = false,
    captureControl = 'manual',
    { allowManualAssist = true } = {},
  ) {
    await cdp.send('Network.enable', {
      maxTotalBufferSize: 100_000_000,
      maxResourceBufferSize: 10_000_000,
    });

    const eventScope = createCdpEventScope(cdp);
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

    eventScope.on('Network.responseReceived', (params) => {
      if (trackedResponses.size >= MAX_TRACKED_RESPONSES) {
        // Continue recording extended debug traffic even after candidate capture is full.
      } else if (shouldTrackNetworkResponse(params.response, params.type)) {
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

        trackedResponses.set(requestId, {
          requestId,
          url,
          status: params.response.status,
          mimeType,
          resourceType,
          headers,
          score,
        });
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

    eventScope.on('Network.requestWillBeSent', (params) => {
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

    eventScope.on('Network.loadingFinished', (params) => {
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

    eventScope.on('Network.loadingFailed', (params) => {
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

    eventScope.on(
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

    eventScope.on(
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

    eventScope.on(
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

    eventScope.on(
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

    eventScope.on(
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

    eventScope.on(
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

    eventScope.on(
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

    try {
      if (captureControl === 'automatic') {
        await runAutomaticTranscriptCaptureFlow(
          cdp,
          prompt,
          captureFeedback,
          debugEnabled,
          {
            allowManualAssist,
          },
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

      eventScope.stopAll();

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
          bodyRecord = createEmptyResponseBodyRecord(failed);
        } else {
          bodyRecord = createEmptyResponseBodyRecord(
            'Request did not finish before capture ended.',
          );
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
          bodyRecordsByRequestId.get(requestId) ||
            createEmptyResponseBodyRecord('Response body was not captured.'),
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
            bodyRecordsByRequestId.get(responseRecord.requestId) ||
              createEmptyResponseBodyRecord('Response body was not captured.'),
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

      candidates.sort((left, right) => {
        const scoreDelta =
          scoreTranscriptCandidate(right) - scoreTranscriptCandidate(left);
        if (scoreDelta !== 0) {
          return scoreDelta;
        }

        if (right.parsedEntryCount !== left.parsedEntryCount) {
          return right.parsedEntryCount - left.parsedEntryCount;
        }

        return right.score - left.score;
      });

      const transcriptMatches = candidates.filter(
        (candidate) => isUsableTranscriptCandidate(candidate),
      );

      const matchedCandidate = transcriptMatches[0] || null;

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
    } finally {
      eventScope.stopAll();
      await cdp.send('Network.disable').catch(() => {});
    }
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
        name: 'Stream Transcript Extractor',
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

  function buildFallbackMeetingMetadata(targetPage) {
    return {
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
    };
  }

  function attachCliErrorDetails(error, details) {
    if (error && typeof error === 'object') {
      error.details = {
        ...(error.details || {}),
        ...details,
      };
    }

    return error;
  }

  function buildDiscoveredMeetingLabel(item) {
    const title = trimForTerminal(item?.title || item?.url || 'Untitled meeting', 56);
    const subtitle = trimForTerminal(item?.subtitle || '', 40);
    const progress = getCrawlerItemProgress(item);
    const progressParts = [
      progress,
      item?.seenInCurrentDiscovery ? 'seen now' : 'saved only',
    ];
    if (Number(item?.attemptCount || 0) > 0) {
      progressParts.push(`${item.attemptCount}x`);
    }
    if (Number(item?.lastEntryCount || 0) > 0 && progress === 'done') {
      progressParts.push(`${item.lastEntryCount} entries`);
    }
    const sourceLabel =
      Array.isArray(item?.sources) && item.sources.length > 0
        ? item.sources.join('+')
        : 'unknown';
    const progressLabel = progressParts.join(', ');

    return subtitle
      ? `${title} | ${progressLabel} | ${subtitle} | ${sourceLabel}`
      : `${title} | ${progressLabel} | ${sourceLabel}`;
  }

  function filterCrawlerSelectionIndexes(items, predicate, emptyMessage) {
    const indexes = [];

    for (let index = 0; index < items.length; index += 1) {
      if (predicate(items[index], index)) {
        indexes.push(index);
      }
    }

    if (indexes.length === 0) {
      throw new CliError(emptyMessage);
    }

    return indexes;
  }

  function resolveCrawlerSelectionIndexes(answer, items) {
    const normalizedAnswer = String(answer || '').trim().toLowerCase();

    if (!normalizedAnswer || normalizedAnswer === 'pending') {
      return filterCrawlerSelectionIndexes(
        items,
        (item) => getCrawlerItemProgress(item) !== 'done',
        'No pending, new, or failed items are queued. Enter all, done, or a numeric selection.',
      );
    }

    if (normalizedAnswer === 'new') {
      return filterCrawlerSelectionIndexes(
        items,
        (item) => getCrawlerItemProgress(item) === 'new',
        'No new items were discovered in this run.',
      );
    }

    if (normalizedAnswer === 'failed') {
      return filterCrawlerSelectionIndexes(
        items,
        (item) => getCrawlerItemProgress(item) === 'failed',
        'No failed items are recorded in the crawl state.',
      );
    }

    if (
      normalizedAnswer === 'done' ||
      normalizedAnswer === 'completed' ||
      normalizedAnswer === 'success'
    ) {
      return filterCrawlerSelectionIndexes(
        items,
        (item) => getCrawlerItemProgress(item) === 'done',
        'No completed items are recorded in the crawl state.',
      );
    }

    if (normalizedAnswer === 'all') {
      return items.map((_item, index) => index);
    }

    return parseSelectionSpec(answer, items.length);
  }

  function updateCrawlerStateItem(item, patch) {
    Object.assign(item, patch);
    return item;
  }

  function buildBatchStatusPayload({
    options,
    browser,
    profile,
    discovery,
    results,
    startedAt,
    completedAt,
    stateFilePath,
    stateSummary,
  }) {
    return {
      app: {
        name: 'Stream Transcript Extractor',
        version: BUILD_VERSION,
        buildTime: BUILD_TIME,
        platform: CURRENT_PLATFORM,
      },
      startedAt,
      completedAt,
      options: {
        startUrl: options.startUrl,
        outputDir: resolveOutputDirectory(options.outputDir),
        outputFormat: options.outputFormat,
        debug: options.debug,
        stateFilePath,
      },
      browser: {
        key: browser.key,
        name: browser.name,
      },
      profile: {
        dirName: profile.dirName,
        displayName: profile.displayName,
        email: profile.email,
      },
      discovery: {
        page: discovery.page,
        meetingsControlLabel: discovery.meetingsControlLabel,
        meetingsSelectedBeforeClick: discovery.meetingsSelectedBeforeClick,
        scrollIterations: discovery.scrollIterations,
        trackedResponseCount: discovery.trackedResponseCount,
        networkItemCount: discovery.networkItemCount,
        graphResolvedItemCount: discovery.graphResolvedItemCount || 0,
        graphThumbnailResolvedItemCount:
          discovery.graphThumbnailResolvedItemCount || 0,
        domGraphResolvedItemCount: discovery.domGraphResolvedItemCount || 0,
        domItemCount: discovery.domItemCount,
        visibleCardCount:
          typeof discovery.visibleCardCount === 'number'
            ? discovery.visibleCardCount
            : Array.isArray(discovery.visibleCardTitles)
              ? discovery.visibleCardTitles.length
              : 0,
        visibleCardTitles: discovery.visibleCardTitles || [],
        discoveredItemCount: discovery.items.length,
        items: discovery.items,
      },
      stateSummary,
      selectedItemCount: results.length,
      results,
    };
  }

  function buildCrawlerDiscoveryDebugPayload({
    options,
    browser,
    profile,
    discovery,
    stateFilePath,
    discoveredAt,
  }) {
    return {
      app: {
        name: 'Stream Transcript Extractor',
        version: BUILD_VERSION,
        buildTime: BUILD_TIME,
        platform: CURRENT_PLATFORM,
      },
      discoveredAt,
      stateFilePath,
      options: {
        startUrl: options.startUrl,
        outputDir: resolveOutputDirectory(options.outputDir),
        outputFormat: options.outputFormat,
        debug: options.debug,
      },
      browser: {
        key: browser.key,
        name: browser.name,
      },
      profile: {
        dirName: profile.dirName,
        displayName: profile.displayName,
        email: profile.email,
      },
      discovery: {
        page: discovery.page,
        meetingsControlLabel: discovery.meetingsControlLabel,
        meetingsSelectedBeforeClick: discovery.meetingsSelectedBeforeClick,
        scrollIterations: discovery.scrollIterations,
        trackedResponseCount: discovery.trackedResponseCount,
        networkItemCount: discovery.networkItemCount,
        graphResolvedItemCount: discovery.graphResolvedItemCount || 0,
        graphThumbnailResolvedItemCount:
          discovery.graphThumbnailResolvedItemCount || 0,
        domGraphResolvedItemCount: discovery.domGraphResolvedItemCount || 0,
        domItemCount: discovery.domItemCount,
        visibleCardCount:
          typeof discovery.visibleCardCount === 'number'
            ? discovery.visibleCardCount
            : Array.isArray(discovery.visibleCardTitles)
              ? discovery.visibleCardTitles.length
              : 0,
        visibleCardTitles: discovery.visibleCardTitles || [],
        discoveredItemCount: discovery.items.length,
        items: discovery.items,
        debug: discovery.debug || null,
      },
    };
  }

  async function extractTranscriptFromConnectedPage({
    cdp,
    prompt,
    options,
    browser,
    profile,
    debugPort,
    targetPage,
    captureControl = 'manual',
    allowManualAssist = true,
  }) {
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
      {
        allowManualAssist,
      },
    );
    const pageMetadata = await extractMeetingMetadata(cdp).catch(() =>
      buildFallbackMeetingMetadata(targetPage),
    );
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

      throw attachCliErrorDetails(
        new CliError(
          buildMissingTranscriptCaptureMessage(
            captureResult,
            options.debug,
            captureControl,
          ),
        ),
        {
          networkOutputPath,
          outputBasePath,
        },
      );
    }

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
    }

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
    if (networkOutputPath) {
      console.log(`Saved network capture to: ${networkOutputPath}`);
    }

    return {
      metadata,
      outputPayload,
      outputBasePath,
      outputPaths,
      networkOutputPath,
      matchedCandidateUrl: captureResult.matchedCandidate.url,
      entryCount: entries.length,
      captureSummary: {
        candidateResponseCount: captureResult.candidates.length,
        transcriptMatchCount: captureResult.transcriptMatchCount || 0,
      },
    };
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
      await extractTranscriptFromConnectedPage({
        cdp,
        prompt,
        options,
        browser,
        profile,
        debugPort,
        targetPage,
        captureControl,
      });

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

  async function runCrawler() {
    const options = parseCrawlerArgs(process.argv.slice(2));

    if (options.help) {
      printCrawlerHelp();
      return 0;
    }

    if (options.version) {
      printVersion();
      return 0;
    }

    ensureSupportedPlatform();

    const prompt = createPrompt();
    const stateFilePath = resolveCrawlerStatePath(options);
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

      tempDataDir = join(tmpdir(), `stream-transcript-extractor-crawl-${Date.now()}`);
      mkdirSync(tempDataDir, { recursive: true });

      console.log('Preparing temporary browser profile...');
      prepareTempProfile(browser.basePath, profile, tempDataDir);

      console.log(`Launching ${browser.name} on debug port ${debugPort}...`);
      browserProcess = launchBrowser(browser, profile, tempDataDir, debugPort);

      console.log('Waiting for the browser debug endpoint...');
      await waitForBrowserDebugEndpoint(debugPort);
      const pages = await waitForBrowserPageTarget(debugPort);
      targetPage = pages[0];

      if (!targetPage) {
        throw new CliError('No browser page target is available for crawling.');
      }

      cdp = await connectToPage(targetPage.webSocketDebuggerUrl);
      await cdp.send('Runtime.enable');

      console.log(`Opening Stream: ${options.startUrl}`);
      if (options.waitBeforeDiscoveryMs > 0) {
        console.log(
          `Waiting ${options.waitBeforeDiscoveryMs} ms on the Stream page so auth and page load can settle...`,
        );
      }

      console.log('\nDiscovering meetings from Stream network payloads and the current page...');
      const discovery = await discoverMeetingPages(cdp, options.startUrl, {
        navigate: true,
        waitBeforeDomDiscoveryMs: options.waitBeforeDiscoveryMs,
        debugEnabled: options.debug,
      });
      const discoveredAt = new Date().toISOString();
      const existingState = loadCrawlerState(stateFilePath);
      const crawlItems = mergeCrawlerStateItems(
        existingState.items,
        discovery.items,
        discoveredAt,
      );
      const crawlItemsByUrl = new Map(
        crawlItems.map((item) => [item.url, item]),
      );
      let stateSummary = buildCrawlerStateSummary(crawlItems);

      function persistCrawlerState(updatedAt) {
        stateSummary = buildCrawlerStateSummary(crawlItems);
        saveCrawlerStateOutput(
          buildCrawlerStatePayload({
            options,
            browser,
            profile,
            stateFilePath,
            items: crawlItems,
            updatedAt,
          }),
          stateFilePath,
        );
      }

      persistCrawlerState(discoveredAt);

      if (options.debug) {
        const discoveryDebugPath = saveCrawlerDiscoveryDebugOutput(
          buildCrawlerDiscoveryDebugPayload({
            options,
            browser,
            profile,
            discovery,
            stateFilePath,
            discoveredAt,
          }),
          stateFilePath,
        );
        console.log(`Discovery debug file: ${discoveryDebugPath}`);
      }

      console.log(
        `Selected Meetings filter "${discovery.meetingsControlLabel || 'Meetings'}".`,
      );
      console.log(
        `Scrolled ${discovery.scrollIterations} time` +
          `${discovery.scrollIterations === 1 ? '' : 's'} and discovered ` +
          `${discovery.items.length} meeting page` +
          `${discovery.items.length === 1 ? '' : 's'}.`,
      );
      console.log(
        `Discovery sources: ${discovery.domItemCount} DOM item` +
          `${discovery.domItemCount === 1 ? '' : 's'}, ` +
          `${discovery.networkItemCount} network match` +
          `${discovery.networkItemCount === 1 ? '' : 'es'}.`,
      );
      if (Number(discovery.graphResolvedItemCount || 0) > 0) {
        console.log(
          `Network fallback resolved ${discovery.graphResolvedItemCount} item` +
            `${discovery.graphResolvedItemCount === 1 ? '' : 's'} via Microsoft Graph ` +
            `(${discovery.graphThumbnailResolvedItemCount || 0} from thumbnail traffic, ` +
            `${discovery.domGraphResolvedItemCount || 0} from card metadata).`,
        );
      }
      if (
        typeof discovery.visibleCardCount === 'number'
          ? discovery.visibleCardCount > 0
          : Array.isArray(discovery.visibleCardTitles) &&
            discovery.visibleCardTitles.length > 0
      ) {
        console.log(
          `Visible meeting cards in the rendered list: ${
            typeof discovery.visibleCardCount === 'number'
              ? discovery.visibleCardCount
              : discovery.visibleCardTitles.length
          }.`,
        );
      }
      console.log(
        `Tracked ${discovery.trackedResponseCount} discovery response` +
          `${discovery.trackedResponseCount === 1 ? '' : 's'}.`,
      );
      console.log(`Queue state file: ${stateFilePath}`);
      console.log(
        `Queue status: ${stateSummary.newItemCount} new, ` +
          `${stateSummary.pendingItemCount} pending, ` +
          `${stateSummary.failedItemCount} failed, ` +
          `${stateSummary.successItemCount} done.`,
      );

      if (crawlItems.length === 0) {
        throw new CliError(
          'No meeting pages were discovered from the Stream Meetings view.',
        );
      }
      if (discovery.items.length === 0) {
        console.log(
          'No current meetings were discovered in this run. Using the saved crawl queue instead.',
        );
      }

      let selection;
      if (options.selectionSpec) {
        let indexes;
        try {
          indexes = resolveCrawlerSelectionIndexes(
            options.selectionSpec,
            crawlItems,
          );
        } catch (error) {
          if (
            error instanceof CliError &&
            String(error.message || '').startsWith('No ')
          ) {
            console.log(error.message);
            return 0;
          }
          throw error;
        }
        selection = {
          indexes,
          items: indexes.map((index) => crawlItems[index]),
        };
        console.log(
          `Selection "${options.selectionSpec}" matched ` +
            `${selection.items.length} queue item` +
            `${selection.items.length === 1 ? '' : 's'}.`,
        );
      } else {
        selection = await chooseManyFromList(
          prompt,
          'Crawler queue',
          crawlItems,
          buildDiscoveredMeetingLabel,
          '\nSelect meetings to extract (Enter = pending queue; or use 1-5, new, failed, done, all): ',
          {
            resolveSelection: resolveCrawlerSelectionIndexes,
          },
        );
      }
      const selectedItems = selection.items;
      if (selectedItems.length === 0) {
        console.log('No crawl items matched the requested selection.');
        return 0;
      }
      const selectedAt = new Date().toISOString();
      for (const item of selectedItems) {
        const stateItem = crawlItemsByUrl.get(item.url);
        if (stateItem) {
          updateCrawlerStateItem(stateItem, {
            lastSelectedAt: selectedAt,
          });
        }
      }
      persistCrawlerState(selectedAt);
      const itemOptions = {
        ...options,
        outputName: '',
      };
      const startedAt = new Date().toISOString();
      const results = [];

      console.log(
        `\nStarting batch extraction for ${selectedItems.length} meeting` +
          `${selectedItems.length === 1 ? '' : 's'}.`,
      );
      console.log(
        'Reusing the same browser debug session for discovery and all extraction attempts.',
      );

      for (let index = 0; index < selectedItems.length; index += 1) {
        const selectedItem = selectedItems[index];
        const stateItem = crawlItemsByUrl.get(selectedItem.url);
        console.log(
          `\n[${index + 1}/${selectedItems.length}] ` +
            `${selectedItem.title || selectedItem.url}`,
        );

        const attemptStartedAt = new Date().toISOString();
        if (stateItem) {
          updateCrawlerStateItem(stateItem, {
            lastSelectedAt: attemptStartedAt,
            lastAttemptedAt: attemptStartedAt,
          });
          persistCrawlerState(attemptStartedAt);
        }

        try {
          const candidateTargetUrls = buildCrawlerExtractionTargetUrls(selectedItem);
          let extractionResult = null;
          let extractionTargetUrl = '';
          let extractionPageSnapshot = null;
          let lastExtractionError = null;

          for (
            let candidateIndex = 0;
            candidateIndex < candidateTargetUrls.length;
            candidateIndex += 1
          ) {
            const candidateTargetUrl = candidateTargetUrls[candidateIndex];
            console.log(
              `Trying playback URL ${candidateIndex + 1}/${candidateTargetUrls.length}: ${candidateTargetUrl}`,
            );

            const pageSnapshot = await navigatePageAndWait(cdp, candidateTargetUrl);
            if (isClearlyWrongMeetingLandingPage(pageSnapshot, candidateTargetUrl)) {
              lastExtractionError = new CliError(
                `Navigation landed on an unexpected page instead of the meeting recording: ` +
                  `${pageSnapshot.title || pageSnapshot.url || candidateTargetUrl}`,
              );
              console.log(
                `Skipping URL because it landed on "${pageSnapshot.title || pageSnapshot.url}".`,
              );
              continue;
            }

            const currentTargetPage = {
              ...targetPage,
              title:
                pageSnapshot.title || selectedItem.title || targetPage.title,
              url: pageSnapshot.url || candidateTargetUrl,
            };

            try {
              extractionResult = await extractTranscriptFromConnectedPage({
                cdp,
                prompt,
                options: itemOptions,
                browser,
                profile,
                debugPort,
                targetPage: currentTargetPage,
                captureControl: 'automatic',
                allowManualAssist: false,
              });
              extractionTargetUrl = candidateTargetUrl;
              extractionPageSnapshot = pageSnapshot;
              break;
            } catch (error) {
              lastExtractionError = error;
              if (candidateIndex < candidateTargetUrls.length - 1) {
                console.log(
                  `Extraction did not succeed from this URL. Trying the next candidate...`,
                );
              }
            }
          }

          if (!extractionResult) {
            throw (
              lastExtractionError ||
              new CliError(
                'No usable playback URL produced a transcript for this meeting.',
              )
            );
          }

          results.push({
            index: index + 1,
            title: extractionResult.metadata.title || selectedItem.title,
            url: selectedItem.url,
            extractionTargetUrl:
              extractionTargetUrl || extractionPageSnapshot?.url || selectedItem.url,
            status: 'success',
            entryCount: extractionResult.entryCount,
            matchedCandidateUrl: extractionResult.matchedCandidateUrl,
            outputPaths: extractionResult.outputPaths,
            networkOutputPath: extractionResult.networkOutputPath,
            captureSummary: extractionResult.captureSummary,
          });

          if (stateItem) {
            updateCrawlerStateItem(stateItem, {
              title:
                extractionResult.metadata.title ||
                selectedItem.title ||
                stateItem.title,
              lastStatus: 'success',
              attemptCount: stateItem.attemptCount + 1,
              successCount: stateItem.successCount + 1,
              lastSucceededAt: new Date().toISOString(),
              lastError: '',
              lastEntryCount: extractionResult.entryCount,
              lastMatchedCandidateUrl: extractionResult.matchedCandidateUrl,
              outputPaths: extractionResult.outputPaths,
              networkOutputPath: extractionResult.networkOutputPath,
              candidateUrls: candidateTargetUrls,
              isNewThisRun: false,
            });
            persistCrawlerState(new Date().toISOString());
          }
        } catch (error) {
          const message =
            error instanceof Error ? error.message : 'Unknown extraction failure.';
          console.error(`\nItem failed: ${message}`);
          results.push({
            index: index + 1,
            title: selectedItem.title,
            url: selectedItem.url,
            extractionTargetUrl: '',
            status: 'failed',
            entryCount: 0,
            matchedCandidateUrl: '',
            outputPaths: [],
            networkOutputPath: error?.details?.networkOutputPath || '',
            captureSummary: null,
            error: message,
          });

          if (stateItem) {
            updateCrawlerStateItem(stateItem, {
              lastStatus: 'failed',
              attemptCount: stateItem.attemptCount + 1,
              failureCount: stateItem.failureCount + 1,
              lastFailedAt: new Date().toISOString(),
              lastError: message,
              lastEntryCount: stateItem.lastEntryCount,
              lastMatchedCandidateUrl: '',
              outputPaths: stateItem.outputPaths,
              networkOutputPath:
                error?.details?.networkOutputPath || stateItem.networkOutputPath,
              isNewThisRun: false,
            });
            persistCrawlerState(new Date().toISOString());
          }
        }
      }

      const completedAt = new Date().toISOString();
      persistCrawlerState(completedAt);
      const batchStatusPayload = buildBatchStatusPayload({
        options,
        browser,
        profile,
        discovery,
        results,
        startedAt,
        completedAt,
        stateFilePath,
        stateSummary,
      });
      const batchStatusPath = saveBatchStatusOutput(
        batchStatusPayload,
        buildOutputBasePath('crawl', options.outputName, options.outputDir),
      );
      const successCount = results.filter((result) => result.status === 'success').length;
      const failureCount = results.length - successCount;

      console.log('\nBatch summary:');
      console.log(`Successful items: ${successCount}`);
      console.log(`Failed items: ${failureCount}`);
      console.log(
        `Queue after run: ${stateSummary.newItemCount} new, ` +
          `${stateSummary.pendingItemCount} pending, ` +
          `${stateSummary.failedItemCount} failed, ` +
          `${stateSummary.successItemCount} done.`,
      );
      console.log(`Updated crawl state: ${stateFilePath}`);
      console.log(`Saved batch status to: ${batchStatusPath}`);

      return failureCount === 0 ? 0 : 1;
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

  return {
    runEmbeddedNetworkMode: async (captureControl = 'manual') => run(captureControl),
    runCrawler,
    parseSelectionSpec,
    mergeDiscoveredMeetingItems,
  };
})();

async function runSelectedWorkflow() {
  const originalArgv = [...process.argv];
  let rawArgv = process.argv.slice(2);

  if (isInteractiveEntrypointLaunch(rawArgv)) {
    rawArgv = await promptForInteractiveLaunchArgs();
  }

  const { workflow, mode, modeProvided, forwardedArgs } = resolveWorkflowSelection(
    rawArgv,
  );
  process.argv = [originalArgv[0], originalArgv[1], ...forwardedArgs];

  try {
    if (workflow === 'crawl') {
      return await networkModeRuntime.runCrawler();
    }

    if (
      !modeProvided &&
      (forwardedArgs.includes('--help') || forwardedArgs.includes('-h'))
    ) {
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
      return await networkModeRuntime.runEmbeddedNetworkMode('automatic');
    }

    return await networkModeRuntime.runEmbeddedNetworkMode('manual');
  } finally {
    process.argv = originalArgv;
  }
}

export async function runExtractCli() {
  return runSelectedWorkflow();
}

export async function runCrawlerCli() {
  return networkModeRuntime.runCrawler();
}

export function parseCrawlerSelectionSpec(value, itemCount) {
  return networkModeRuntime.parseSelectionSpec(value, itemCount);
}

export function mergeCrawlerDiscoveryItems(networkItems, domItems) {
  return networkModeRuntime.mergeDiscoveredMeetingItems(networkItems, domItems);
}

if (import.meta.main) {
  let exitCode = 0;

  try {
    exitCode = await runSelectedWorkflow();
  } catch (error) {
    exitCode = handleEntrypointError(error);
  }

  process.exitCode = exitCode;
}
