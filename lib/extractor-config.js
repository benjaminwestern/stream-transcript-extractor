import { platform } from 'node:os';

export const APP_NAME = 'Stream Transcript Extractor';
export const APP_DESCRIPTION =
  'Extract Microsoft Teams recording transcripts from Microsoft Stream ' +
  'using your signed-in Chrome or Edge profile.';
export const DEFAULT_WORKFLOW = 'extract';
export const SUPPORTED_WORKFLOWS = ['extract', 'crawl'];
export const DEFAULT_EXTRACTOR_MODE = 'automatic';
export const SUPPORTED_EXTRACTOR_MODES = ['network', 'automatic', 'dom'];
export const SUPPORTED_BROWSER_KEYS = ['chrome', 'edge'];
export const SUPPORTED_OUTPUT_FORMATS = ['json', 'md', 'both'];
export const DEFAULT_TRANSCRIPT_OUTPUT_FORMAT = 'md';
export const DEFAULT_OUTPUT_DIR = 'output';
export const DEFAULT_CRAWL_START_URL =
  'https://m365.cloud.microsoft/launch/Stream/?auth=2&home=1';
export const DEFAULT_CDP_HOST = '127.0.0.1';
export const DEFAULT_BROWSER_READY_TIMEOUT_MS = 15_000;
export const DEFAULT_CDP_TIMEOUT_MS = 30_000;
export const EXTRACTION_TIMEOUT_MS = 10 * 60_000;
export const DEFAULT_BROWSER_SHUTDOWN_TIMEOUT_MS = 8_000;

export const EXTRACTION_SETTINGS = Object.freeze({
  dedupePrefixLength: 50,
  maxScrollIterations: 2_000,
  scrollSettleMs: 300,
  stableScrollPasses: 8,
  viewportChunkRatio: 0.8,
});

export const BUILD_VERSION =
  typeof __BUILD_VERSION__ === 'string' ? __BUILD_VERSION__ : 'dev';
export const BUILD_TIME =
  typeof __BUILD_TIME__ === 'string' ? __BUILD_TIME__ : '';

export const CURRENT_PLATFORM = platform();
export const IS_WINDOWS = CURRENT_PLATFORM === 'win32';
export const IS_MACOS = CURRENT_PLATFORM === 'darwin';
