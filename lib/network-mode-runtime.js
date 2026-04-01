import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import {
  createCdpEventScope,
  navigatePageAndWait as navigateWithCdp,
  reloadPageAndWait as reloadWithCdp,
} from './cdp.js';
import { parseCliArgs as parseSharedCliArgs, printHelpScreen } from './cli.js';
import { createCrawlerStateRuntime } from './crawl-state.js';
import { createMeetingDiscoveryRuntime } from './meeting-discovery.js';
import { buildMeetingDiscoveryExpression } from './meeting-discovery-dom.js';
import { createMeetingNavigationRuntime } from './meeting-navigation.js';
import { createMeetingPageDiscoveryRuntime } from './meeting-page-discovery.js';
import {
  createResponseBodyRecord,
  loadResponseBody as loadCapturedResponseBody,
} from './network-capture.js';
import {
  createAutomaticTranscriptCaptureRuntime,
  logAutomaticAction,
} from './automatic-transcript-capture.js';
import { createConnectedSessionExtractionRuntime } from './connected-session-extraction.js';
import { createTranscriptNetworkRuntime } from './transcript-network.js';
import { createNetworkTranscriptCaptureRuntime } from './network-transcript-capture.js';
import {
  buildTranscriptOutputBasePath,
  buildTranscriptOutputPayload,
  resolveOutputDirectory as resolveSharedOutputDirectory,
  sanitizeOutputFilename,
  saveTranscriptOutputs,
  writeJsonOutput,
} from './transcript-output.js';
import { selectTranscriptPage as selectSharedTranscriptPage } from './transcript-pages.js';
import {
  APP_NAME,
  BUILD_TIME,
  BUILD_VERSION,
  CURRENT_PLATFORM,
  DEFAULT_CDP_TIMEOUT_MS,
  DEFAULT_CRAWL_START_URL,
  DEFAULT_OUTPUT_DIR,
  DEFAULT_TRANSCRIPT_OUTPUT_FORMAT,
  SUPPORTED_BROWSER_KEYS,
  SUPPORTED_OUTPUT_FORMATS,
} from './extractor-config.js';
import { CliError, createCliRuntime } from './extractor-cli-runtime.js';
import {
  extractMeetingMetadata as extractMeetingMetadataFromPage,
  extractMeetingMetadataFromCapture,
  mergeMeetingMetadata,
  renderTranscriptMarkdown,
} from './meeting-metadata.js';

const NETWORK_APP_NAME = `${APP_NAME} (Network Mode)`;
const NETWORK_APP_DESCRIPTION =
  'Extract Microsoft Teams recording transcripts from Microsoft Stream ' +
  'using your signed-in Chrome or Edge profile.';
const DEFAULT_PAGE_NAVIGATION_TIMEOUT_MS = 45_000;
const DEFAULT_CAPTURE_SETTLE_MS = 1_500;
const DEFAULT_CRAWL_SCROLL_SETTLE_MS = 1_000;
const DEFAULT_WAIT_BEFORE_DISCOVERY_MS = 10_000;
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
const DEFAULT_CRAWLER_AUTO_SELECTION_SPEC = 'pending';

const cliRuntime = createCliRuntime({ appName: APP_NAME });
const {
  errorFactory,
  normalizeBrowserKey,
  ensureSupportedPlatform,
  createPrompt,
  chooseFromList,
  chooseManyFromList,
  parseSelectionSpec,
  findAvailablePort,
  connectToPage,
  evaluate,
  waitForBrowserDebugEndpoint,
  waitForBrowserPageTarget,
  findPageTargets,
  ensureBrowserIsClosed,
  selectBrowserAndProfile,
  prepareTempProfile,
  launchBrowser,
  shutdownBrowser,
} = cliRuntime;

function sleep(ms) {
  return new Promise((resolvePromise) => setTimeout(resolvePromise, ms));
}

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
    errorFactory,
  });
}

function printHelp() {
  printHelpScreen({
    name: 'Stream Transcript Extractor (Automatic and network modes)',
    summary: NETWORK_APP_DESCRIPTION,
    usage: [
      'bun ./cli.js --mode <network|automatic> [options]',
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
          'bun ./cli.js --mode automatic --browser chrome --profile Work',
          'bun ./cli.js --mode automatic --format md',
          'bun ./cli.js --mode network --debug --output-dir ./exports',
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
    errorFactory,
    extraOptions: [
      {
        flag: '--start-url',
        key: 'startUrl',
        normalize: (value) => String(value).trim(),
        validate: (value) => {
          if (!value) {
            throw errorFactory('Missing value for --start-url.');
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
      'bun ./cli.js crawl [options]',
      './stream-transcript-extractor-<target> crawl [options]',
    ],
    sections: [
      {
        title: 'Recommended path',
        rows: [
          {
            label: 'bun ./cli.js crawl',
            description:
              'Open Stream home, wait the default 10-second settle window, switch to Meetings, scroll the rendered list, merge the results into *.state.json, then extract the actionable queue with the automatic extractor in the same browser session. When a saved state or batch exists, reruns resume automatically.',
          },
        ],
      },
      {
        title: 'How crawl works',
        lines: [
          '1. Open Stream home and wait the default 10-second settle window for auth and page load.',
          '2. Select the Meetings view and scroll until the rendered list stops growing.',
          '3. Merge discovered recordings into the persistent *.state.json queue, or restore the saved queue from *.batch.json when needed.',
          '4. Reuse the same browser debug session to run automatic extraction against the actionable queue. Fresh runs can still use the terminal selection menu, while reruns resume automatically from saved queue state.',
          '5. Write the updated queue state and a *.batch.json run summary. By default those files stay stable in the output directory so reruns reopen the same queue.',
        ],
      },
      {
        title: 'Selection syntax',
        lines: [
          'Use comma-separated indexes and ranges such as 1-5,8,10.',
          'Press Enter to select items that are not yet successful. Failed items are retried first, then newly discovered items, then any remaining pending items. You can also use keywords such as new, failed, done, or all.',
          'When a saved *.state.json or *.batch.json queue already exists and you do not pass --select, the crawler automatically retries failed items and picks up newly discovered meetings.',
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
          'bun ./cli.js crawl --browser chrome --profile Work',
          'bun ./cli.js crawl --state-file ./exports/team.state.json',
          'bun ./cli.js crawl --select pending --browser edge',
          `bun ./cli.js crawl --wait-before-discovery-ms ${DEFAULT_WAIT_BEFORE_DISCOVERY_MS} --browser edge`,
          'bun ./cli.js crawl --output-dir ./exports --format md',
          'bun ./cli.js crawl --output-dir ./exports --format both --debug',
        ],
      },
    ],
  });
}

function printVersion() {
  const buildSuffix = BUILD_TIME ? ` (${BUILD_TIME})` : '';
  console.log(`${NETWORK_APP_NAME} ${BUILD_VERSION}${buildSuffix}`);
}

function trimForTerminal(value, maxLength = 120) {
  const text = String(value || '').replace(/\s+/g, ' ').trim();

  if (!text || text.length <= maxLength) {
    return text;
  }

  return `${text.slice(0, maxLength - 3)}...`;
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
  return ['Font', 'Image', 'Media', 'Stylesheet', 'Manifest'].includes(
    String(resourceType || ''),
  );
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

const {
  buildProtectionKeyMap,
  finalizeBodyRecordForResponse,
  getFirstStringValue,
  getUrlSearchParam,
  isEncryptedTranscriptUrl,
  isUsableTranscriptCandidate,
  scoreTranscriptCandidate,
  summarizeCapturedBody,
} = createTranscriptNetworkRuntime({
  containsTranscriptSignal,
  formatTimestampValue,
  normalizeInlineText,
  stripHtmlTags,
  stripUtf8Bom,
  trimForPreview,
  tryParseJson,
  tryParseJsonLenient,
});

function saveNetworkCaptureOutput(payload, outputBasePath) {
  return writeJsonOutput(`${outputBasePath}.network.json`, payload);
}

function saveBatchStatusOutput(payload, outputBasePath) {
  return writeJsonOutput(`${outputBasePath}.batch.json`, payload);
}

function resolveCrawlerArtifactDirectory(options) {
  return resolveSharedOutputDirectory(options.outputDir, {
    defaultOutputDir: DEFAULT_OUTPUT_DIR,
  });
}

function resolveCrawlerArtifactBaseName(options) {
  return sanitizeOutputFilename(options.outputName || 'crawl', {
    fallbackName: 'crawl',
  });
}

function resolveCrawlerOutputBasePath(options) {
  if (options.stateFile) {
    const explicitStateFilePath = resolve(options.stateFile);
    if (explicitStateFilePath.endsWith('.state.json')) {
      return explicitStateFilePath.replace(/\.state\.json$/, '');
    }

    return explicitStateFilePath;
  }

  return join(
    resolveCrawlerArtifactDirectory(options),
    resolveCrawlerArtifactBaseName(options),
  );
}

function resolveLegacyCrawlerArtifactPath(options, kind) {
  const outputDir = resolveCrawlerArtifactDirectory(options);
  const baseName = resolveCrawlerArtifactBaseName(options);
  const suffix = `.${kind}.json`;

  if (!existsSync(outputDir)) {
    return '';
  }

  const legacyFileName = readdirSync(outputDir)
    .filter(
      (entry) =>
        entry.startsWith(`${baseName}_`) &&
        entry.endsWith(suffix),
    )
    .sort()
    .at(-1);

  if (!legacyFileName) {
    return '';
  }

  return join(outputDir, legacyFileName);
}

function resolveCrawlerStatePath(options) {
  if (options.stateFile) {
    return resolve(options.stateFile);
  }

  return `${resolveCrawlerOutputBasePath(options)}.state.json`;
}

function resolveCrawlerBatchPath(options) {
  return `${resolveCrawlerOutputBasePath(options)}.batch.json`;
}

function loadCrawlerBatchStatus(batchFilePath) {
  if (!existsSync(batchFilePath)) {
    return null;
  }

  try {
    return JSON.parse(readFileSync(batchFilePath, 'utf-8'));
  } catch (error) {
    throw errorFactory(
      `Unable to read crawl batch file "${batchFilePath}": ` +
        `${error instanceof Error ? error.message : 'Unknown parse failure.'}`,
    );
  }
}

function resolvePersistedStateFilePath(value, fallback = '') {
  const text = String(value || '').trim();
  if (!text) {
    return fallback;
  }

  return resolve(text);
}

function createEmptyCrawlerState() {
  return {
    stateVersion: CRAWL_STATE_VERSION,
    items: [],
  };
}

function resolveCrawlerPersistenceContext(options) {
  const requestedStateFilePath = resolveCrawlerStatePath(options);
  const batchFilePath = resolveCrawlerBatchPath(options);
  const legacyStateFilePath = options.stateFile
    ? ''
    : resolveLegacyCrawlerArtifactPath(options, 'state');
  const legacyBatchFilePath = options.stateFile
    ? ''
    : resolveLegacyCrawlerArtifactPath(options, 'batch');

  if (existsSync(requestedStateFilePath)) {
    return {
      stateFilePath: requestedStateFilePath,
      batchFilePath,
      existingState: loadCrawlerState(requestedStateFilePath),
      resumeSource: 'state',
      batchPayload: null,
      loadedStateFilePath: requestedStateFilePath,
      loadedBatchFilePath: '',
    };
  }

  if (legacyStateFilePath && existsSync(legacyStateFilePath)) {
    return {
      stateFilePath: requestedStateFilePath,
      batchFilePath,
      existingState: loadCrawlerState(legacyStateFilePath),
      resumeSource: 'legacy-state',
      batchPayload: null,
      loadedStateFilePath: legacyStateFilePath,
      loadedBatchFilePath: '',
    };
  }

  const batchPayload = loadCrawlerBatchStatus(batchFilePath);
  if (!batchPayload) {
    const legacyBatchPayload = legacyBatchFilePath
      ? loadCrawlerBatchStatus(legacyBatchFilePath)
      : null;

    if (!legacyBatchPayload) {
      return {
        stateFilePath: requestedStateFilePath,
        batchFilePath,
        existingState: createEmptyCrawlerState(),
        resumeSource: 'none',
        batchPayload: null,
        loadedStateFilePath: '',
        loadedBatchFilePath: '',
      };
    }

    const legacyLinkedStateFilePath = resolvePersistedStateFilePath(
      legacyBatchPayload?.state?.stateFilePath ||
        legacyBatchPayload?.options?.stateFilePath,
      '',
    );

    if (legacyLinkedStateFilePath && existsSync(legacyLinkedStateFilePath)) {
      return {
        stateFilePath: requestedStateFilePath,
        batchFilePath,
        existingState: loadCrawlerState(legacyLinkedStateFilePath),
        resumeSource: 'legacy-batch-linked-state',
        batchPayload: legacyBatchPayload,
        loadedStateFilePath: legacyLinkedStateFilePath,
        loadedBatchFilePath: legacyBatchFilePath,
      };
    }

    if (Array.isArray(legacyBatchPayload?.state?.items)) {
      return {
        stateFilePath: requestedStateFilePath,
        batchFilePath,
        existingState: {
          stateVersion:
            typeof legacyBatchPayload?.state?.stateVersion === 'number'
              ? legacyBatchPayload.state.stateVersion
              : CRAWL_STATE_VERSION,
          items: legacyBatchPayload.state.items,
        },
        resumeSource: 'legacy-batch-snapshot',
        batchPayload: legacyBatchPayload,
        loadedStateFilePath: '',
        loadedBatchFilePath: legacyBatchFilePath,
      };
    }

    return {
      stateFilePath: requestedStateFilePath,
      batchFilePath,
      existingState: createEmptyCrawlerState(),
      resumeSource: 'none',
      batchPayload: null,
      loadedStateFilePath: '',
      loadedBatchFilePath: '',
    };
  }

  const linkedStateFilePath = resolvePersistedStateFilePath(
    batchPayload?.state?.stateFilePath || batchPayload?.options?.stateFilePath,
    '',
  );

  if (linkedStateFilePath && existsSync(linkedStateFilePath)) {
    return {
      stateFilePath: requestedStateFilePath,
      batchFilePath,
      existingState: loadCrawlerState(linkedStateFilePath),
      resumeSource: 'batch-linked-state',
      batchPayload,
      loadedStateFilePath: linkedStateFilePath,
      loadedBatchFilePath: batchFilePath,
    };
  }

  if (Array.isArray(batchPayload?.state?.items)) {
    return {
      stateFilePath: requestedStateFilePath,
      batchFilePath,
      existingState: {
        stateVersion:
          typeof batchPayload?.state?.stateVersion === 'number'
            ? batchPayload.state.stateVersion
            : CRAWL_STATE_VERSION,
        items: batchPayload.state.items,
      },
      resumeSource: 'batch-snapshot',
      batchPayload,
      loadedStateFilePath: '',
      loadedBatchFilePath: batchFilePath,
    };
  }

  return {
    stateFilePath: requestedStateFilePath,
    batchFilePath,
    existingState: createEmptyCrawlerState(),
    resumeSource: 'none',
    batchPayload,
    loadedStateFilePath: '',
    loadedBatchFilePath: batchFilePath,
  };
}

function saveCrawlerStateOutput(payload, stateFilePath) {
  return writeJsonOutput(stateFilePath, payload, { ensureDirectory: true });
}

function saveCrawlerDiscoveryDebugOutput(payload, stateFilePath) {
  const outputPath = stateFilePath.endsWith('.state.json')
    ? stateFilePath.replace(/\.state\.json$/, '.discovery.debug.json')
    : `${stateFilePath}.discovery.debug.json`;
  return writeJsonOutput(outputPath, payload, { ensureDirectory: true });
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

function normalizeMeetingDownloadUrls(values) {
  return uniqueNormalizedMeetingUrls(values).filter((value) =>
    isLikelyMeetingVideoFileUrl(value),
  );
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

const { mergeDiscoveredMeetingItems } = createMeetingDiscoveryRuntime({
  normalizeMeetingUrl,
  uniqueNormalizedMeetingUrls,
  normalizeMeetingDownloadUrls,
  sortMeetingUrlsForExtraction,
  normalizeInlineText,
  scoreMeetingTargetUrl,
});

const { loadCrawlerState, getCrawlerItemProgress, mergeCrawlerStateItems, buildCrawlerStateSummary, buildCrawlerStatePayload, buildCrawlerExtractionTargetUrls, buildCrawlerDownloadTargetUrls } =
  createCrawlerStateRuntime({
    stateVersion: CRAWL_STATE_VERSION,
    appName: APP_NAME,
    buildVersion: BUILD_VERSION,
    buildTime: BUILD_TIME,
    platform: CURRENT_PLATFORM,
    defaultOutputDir: DEFAULT_OUTPUT_DIR,
    resolveOutputDirectory: resolveSharedOutputDirectory,
    normalizeMeetingUrl,
    normalizeMeetingDownloadUrls,
    uniqueNormalizedMeetingUrls,
    sortMeetingUrlsForExtraction,
    normalizeInlineText,
    errorFactory,
  });

const {
  isClearlyWrongMeetingLandingPage,
  getCurrentPageSnapshot,
  navigatePageAndWait,
} = createMeetingNavigationRuntime({
  defaultCrawlStartUrl: DEFAULT_CRAWL_START_URL,
  defaultPageNavigationTimeoutMs: DEFAULT_PAGE_NAVIGATION_TIMEOUT_MS,
  defaultCrawlScrollSettleMs: DEFAULT_CRAWL_SCROLL_SETTLE_MS,
  navigatePage: navigateWithCdp,
  evaluate,
  normalizeMeetingUrl,
  errorFactory,
});

const { discoverMeetingPages } = createMeetingPageDiscoveryRuntime({
  defaultPageNavigationTimeoutMs: DEFAULT_PAGE_NAVIGATION_TIMEOUT_MS,
  maxDebugResponses: MAX_DEBUG_RESPONSES,
  maxDiscoveryResponseBodies: MAX_DISCOVERY_RESPONSE_BODIES,
  createCdpEventScope,
  loadResponseBody,
  navigatePageAndWait,
  getCurrentPageSnapshot,
  evaluate,
  sleep,
  buildMeetingDiscoveryExpression: () =>
    buildMeetingDiscoveryExpression({
      settleMs: DEFAULT_CRAWL_SCROLL_SETTLE_MS,
      defaultCrawlStartUrl: DEFAULT_CRAWL_START_URL,
    }),
  mergeDiscoveredMeetingItems,
  normalizeMeetingUrl,
  normalizeInlineText,
  stripUtf8Bom,
  tryParseJsonLenient,
  getFirstStringValue,
  uniqueNormalizedMeetingUrls,
  normalizeMeetingDownloadUrls,
  isLikelyMeetingItemUrl,
  isLikelyMeetingVideoFileUrl,
  isStreamRelatedUrl,
  lowerCaseHeaderMap,
  errorFactory,
});

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

const { runAutomaticTranscriptCaptureFlow } =
  createAutomaticTranscriptCaptureRuntime({
    evaluate,
    reloadPageAndWait,
    sleep,
    defaultAutomaticUiPollMs: DEFAULT_AUTOMATIC_UI_POLL_MS,
    defaultAutomaticPanelOpenTimeoutMs:
      DEFAULT_AUTOMATIC_PANEL_OPEN_TIMEOUT_MS,
    defaultAutomaticClickSettleMs: DEFAULT_AUTOMATIC_CLICK_SETTLE_MS,
    defaultAutomaticSignalTimeoutMs: DEFAULT_AUTOMATIC_SIGNAL_TIMEOUT_MS,
    defaultAutomaticSignalRetryTimeoutMs:
      DEFAULT_AUTOMATIC_SIGNAL_RETRY_TIMEOUT_MS,
    defaultAutomaticRequestSettleTimeoutMs:
      DEFAULT_AUTOMATIC_REQUEST_SETTLE_TIMEOUT_MS,
    defaultCaptureSettleMs: DEFAULT_CAPTURE_SETTLE_MS,
  });

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

async function refetchBinaryBody(url) {
  const response = await fetch(url, { redirect: 'follow' });

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
    errorFactory,
  });
}

const { captureStreamNetwork } = createNetworkTranscriptCaptureRuntime({
  createCdpEventScope,
  createEmptyResponseBodyRecord,
  loadResponseBody,
  refetchBinaryBody,
  reloadPageAndWait,
  runAutomaticTranscriptCaptureFlow,
  sleep,
  printCaptureInstructions,
  logAutomaticAction,
  sanitizeHeaders,
  lowerCaseHeaderMap,
  trimForPreview,
  containsTranscriptSignal,
  shouldTrackDebugTraffic,
  shouldFetchDebugResponseBody,
  summarizeWebSocketPayload,
  shouldTrackNetworkResponse,
  scoreNetworkResponse,
  isLikelyTranscriptResponseSignal,
  formatNetworkResponseForTerminal,
  buildProtectionKeyMap,
  isEncryptedTranscriptUrl,
  getUrlSearchParam,
  finalizeBodyRecordForResponse,
  summarizeCapturedBody,
  scoreTranscriptCandidate,
  isUsableTranscriptCandidate,
  maxTrackedResponses: MAX_TRACKED_RESPONSES,
  maxDebugResponses: MAX_DEBUG_RESPONSES,
  maxDebugRequests: MAX_DEBUG_REQUESTS,
  maxDebugBodies: MAX_DEBUG_BODIES,
  maxDebugWebSockets: MAX_DEBUG_WEBSOCKETS,
  maxDebugWebSocketFrames: MAX_DEBUG_WEBSOCKET_FRAMES,
  maxDebugPostDataPreviewLength: MAX_DEBUG_POST_DATA_PREVIEW_LENGTH,
  maxDebugFramePreviewLength: MAX_DEBUG_FRAME_PREVIEW_LENGTH,
  defaultCaptureSettleMs: DEFAULT_CAPTURE_SETTLE_MS,
});

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
    return resolveAutoResumeCrawlerSelectionIndexes(items);
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

function resolveAutoResumeCrawlerSelectionIndexes(items) {
  const indexedItems = [];

  for (let index = 0; index < items.length; index += 1) {
    const progress = getCrawlerItemProgress(items[index]);
    if (progress === 'done') {
      continue;
    }

    indexedItems.push({
      index,
      priority:
        progress === 'failed'
          ? 0
          : progress === 'new'
            ? 1
            : 2,
    });
  }

  if (indexedItems.length === 0) {
    throw new CliError(
      'No pending, new, or failed items are queued. Nothing needs to be retried.',
    );
  }

  indexedItems.sort(
    (left, right) => left.priority - right.priority || left.index - right.index,
  );

  return indexedItems.map((item) => item.index);
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
  statePayload,
  appliedSelectionSpec,
  autoSelected,
  resumeSource,
}) {
  return {
    app: {
      name: APP_NAME,
      version: BUILD_VERSION,
      buildTime: BUILD_TIME,
      platform: CURRENT_PLATFORM,
    },
    startedAt,
    completedAt,
    options: {
      startUrl: options.startUrl,
      outputDir: resolveSharedOutputDirectory(options.outputDir, {
        defaultOutputDir: DEFAULT_OUTPUT_DIR,
      }),
      outputFormat: options.outputFormat,
      debug: options.debug,
      stateFilePath,
    },
    resume: {
      source: resumeSource,
    },
    selection: {
      requestedSpec: String(options.selectionSpec || ''),
      appliedSpec: String(appliedSelectionSpec || ''),
      autoSelected: Boolean(autoSelected),
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
    state: statePayload
      ? {
          stateVersion: statePayload.stateVersion,
          updatedAt: statePayload.updatedAt,
          stateFilePath: statePayload.stateFilePath,
          summary: statePayload.summary,
          items: statePayload.items,
        }
      : null,
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
      name: APP_NAME,
      version: BUILD_VERSION,
      buildTime: BUILD_TIME,
      platform: CURRENT_PLATFORM,
    },
    discoveredAt,
    stateFilePath,
    options: {
      startUrl: options.startUrl,
      outputDir: resolveSharedOutputDirectory(options.outputDir, {
        defaultOutputDir: DEFAULT_OUTPUT_DIR,
      }),
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

const buildMarkdownOutput = (payload) =>
  renderTranscriptMarkdown(payload, { includeExtendedMetadata: true });
const extractMeetingMetadata = (cdp) =>
  extractMeetingMetadataFromPage(cdp, evaluate);

const {
  extractTranscriptFromConnectedPage,
  extractCrawlerItemInConnectedSession,
} = createConnectedSessionExtractionRuntime({
  appName: APP_NAME,
  buildVersion: BUILD_VERSION,
  buildTime: BUILD_TIME,
  platform: CURRENT_PLATFORM,
  defaultOutputDir: DEFAULT_OUTPUT_DIR,
  maxSavedCandidates: MAX_SAVED_CANDIDATES,
  buildTranscriptOutputBasePath,
  buildTranscriptOutputPayload,
  saveTranscriptOutputs,
  buildMarkdownOutput,
  saveNetworkCaptureOutput,
  extractMeetingMetadata,
  extractMeetingMetadataFromCapture: (captureResult) =>
    extractMeetingMetadataFromCapture(captureResult, { isStreamRelatedUrl }),
  mergeMeetingMetadata,
  captureStreamNetwork,
  buildCrawlerExtractionTargetUrls,
  navigatePageAndWait,
  isClearlyWrongMeetingLandingPage,
  errorFactory,
});

function handleError(error) {
  if (error instanceof CliError) {
    console.error(`\nError: ${error.message}`);
    return error.exitCode;
  }

  console.error('\nUnexpected error:');
  console.error(error);
  return 1;
}

async function selectTranscriptPage(prompt, debugPort) {
  try {
    return await selectSharedTranscriptPage({
      prompt,
      debugPort,
      findPageTargets,
      chooseFromList,
      selectionQuestion:
        '\nWhich page contains the Microsoft Stream meeting? (number): ',
    });
  } catch (error) {
    throw error instanceof Error ? new CliError(error.message) : error;
  }
}

export async function runEmbeddedNetworkMode(
  captureControl = 'manual',
  argv = process.argv.slice(2),
) {
  const options = parseArgs(argv);

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

export async function runCrawler(argv = process.argv.slice(2)) {
  const options = parseCrawlerArgs(argv);

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
  const persistenceContext = resolveCrawlerPersistenceContext(options);
  const stateFilePath = persistenceContext.stateFilePath;
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

    console.log(`Opening Stream: ${options.startUrl}`);
    if (options.waitBeforeDiscoveryMs > 0) {
      console.log(
        `Waiting ${options.waitBeforeDiscoveryMs} ms on the Stream page so auth and page load can settle...`,
      );
    }

    if (persistenceContext.resumeSource === 'state') {
      console.log(`Loaded saved crawl state: ${stateFilePath}`);
    } else if (persistenceContext.resumeSource === 'legacy-state') {
      console.log(
        `Loaded saved crawl state from legacy artifact: ${persistenceContext.loadedStateFilePath}`,
      );
      console.log(`Migrating queue state to: ${stateFilePath}`);
    } else if (persistenceContext.resumeSource === 'batch-linked-state') {
      console.log(
        `Loaded saved crawl state via batch metadata: ${persistenceContext.loadedStateFilePath}`,
      );
      console.log(`Migrating queue state to: ${stateFilePath}`);
    } else if (
      persistenceContext.resumeSource === 'legacy-batch-linked-state'
    ) {
      console.log(
        `Loaded saved crawl state via legacy batch metadata: ${persistenceContext.loadedStateFilePath}`,
      );
      console.log(`Migrating queue state to: ${stateFilePath}`);
    } else if (persistenceContext.resumeSource === 'batch-snapshot') {
      console.log(
        `Restored saved crawl queue from batch metadata: ${persistenceContext.loadedBatchFilePath}`,
      );
      console.log(`Migrating queue state to: ${stateFilePath}`);
    } else if (
      persistenceContext.resumeSource === 'legacy-batch-snapshot'
    ) {
      console.log(
        `Restored saved crawl queue from legacy batch metadata: ${persistenceContext.loadedBatchFilePath}`,
      );
      console.log(`Migrating queue state to: ${stateFilePath}`);
    }

    console.log('\nDiscovering meetings from Stream network payloads and the current page...');
    const discovery = await discoverMeetingPages(cdp, options.startUrl, {
      navigate: true,
      waitBeforeDomDiscoveryMs: options.waitBeforeDiscoveryMs,
      debugEnabled: options.debug,
    });
    const discoveredAt = new Date().toISOString();
    const existingState = persistenceContext.existingState;
    const crawlItems = mergeCrawlerStateItems(
      existingState.items,
      discovery.items,
      discoveredAt,
    );
    const crawlItemsByUrl = new Map(crawlItems.map((item) => [item.url, item]));
    let stateSummary = buildCrawlerStateSummary(crawlItems);
    let currentStatePayload = null;

    function persistCrawlerState(updatedAt) {
      stateSummary = buildCrawlerStateSummary(crawlItems);
      currentStatePayload = buildCrawlerStatePayload({
        options,
        browser,
        profile,
        stateFilePath,
        items: crawlItems,
        updatedAt,
      });
      saveCrawlerStateOutput(
        currentStatePayload,
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
    let appliedSelectionSpec = String(options.selectionSpec || '');
    let autoSelected = false;
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
    } else if (persistenceContext.resumeSource !== 'none') {
      let indexes;
      try {
        indexes = resolveAutoResumeCrawlerSelectionIndexes(crawlItems);
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

      appliedSelectionSpec = DEFAULT_CRAWLER_AUTO_SELECTION_SPEC;
      autoSelected = true;
      selection = {
        indexes,
        items: indexes.map((index) => crawlItems[index]),
      };
      console.log(
        `Resuming from saved crawl queue: automatically selected ` +
          `${selection.items.length} actionable meeting` +
          `${selection.items.length === 1 ? '' : 's'} ` +
          '(failed first, then new, then remaining pending items).',
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
        const extractionResult = await extractCrawlerItemInConnectedSession({
          cdp,
          prompt,
          options: itemOptions,
          browser,
          profile,
          debugPort,
          targetPage,
          item: selectedItem,
          captureControl: 'automatic',
          allowManualAssist: false,
        });

        results.push({
          index: index + 1,
          title: extractionResult.metadata.title || selectedItem.title,
          url: selectedItem.url,
          extractionTargetUrl:
            extractionResult.extractionTargetUrl || selectedItem.url,
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
            candidateUrls: extractionResult.candidateTargetUrls,
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
      statePayload: currentStatePayload,
      appliedSelectionSpec,
      autoSelected,
      resumeSource: persistenceContext.resumeSource,
    });
    const batchStatusPath = saveBatchStatusOutput(
      batchStatusPayload,
      resolveCrawlerOutputBasePath(options),
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

export function parseCrawlerSelectionSpec(value, itemCount) {
  return parseSelectionSpec(value, itemCount);
}

export function mergeCrawlerDiscoveryItems(networkItems, domItems) {
  return mergeDiscoveredMeetingItems(networkItems, domItems);
}

export async function discoverCrawlerMeetingPages(
  cdp,
  startUrl = DEFAULT_CRAWL_START_URL,
  options = {},
) {
  return discoverMeetingPages(cdp, startUrl, options);
}

export function buildCrawlerExtractionTargets(item) {
  return buildCrawlerExtractionTargetUrls(item);
}

export function buildCrawlerDownloadTargets(item) {
  return buildCrawlerDownloadTargetUrls(item);
}

export function getCrawlerItemProgressStatus(item) {
  return getCrawlerItemProgress(item);
}

export async function extractCrawlerItemTranscriptInSession({
  cdp,
  item,
  options = {},
  browser = null,
  profile = null,
  debugPort = null,
  targetPage = null,
  prompt = null,
  captureControl = 'automatic',
  allowManualAssist = false,
}) {
  return extractCrawlerItemInConnectedSession({
    cdp,
    item,
    options: {
      outputName: '',
      outputDir: DEFAULT_OUTPUT_DIR,
      outputFormat: DEFAULT_TRANSCRIPT_OUTPUT_FORMAT,
      debug: false,
      ...options,
    },
    browser,
    profile,
    debugPort,
    targetPage: targetPage || { title: item?.title || '', url: item?.url || '' },
    prompt,
    captureControl,
    allowManualAssist,
  });
}
