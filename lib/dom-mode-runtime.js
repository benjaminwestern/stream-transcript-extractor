import { mkdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { parseCliArgs as parseSharedCliArgs } from './cli.js';
import {
  APP_NAME,
  BUILD_TIME,
  BUILD_VERSION,
  CURRENT_PLATFORM,
  DEFAULT_OUTPUT_DIR,
  DEFAULT_TRANSCRIPT_OUTPUT_FORMAT,
  EXTRACTION_SETTINGS,
  EXTRACTION_TIMEOUT_MS,
  SUPPORTED_BROWSER_KEYS,
  SUPPORTED_OUTPUT_FORMATS,
} from './extractor-config.js';
import { CliError, createCliRuntime } from './extractor-cli-runtime.js';
import {
  extractMeetingMetadata,
  renderTranscriptMarkdown,
} from './meeting-metadata.js';
import { buildTranscriptExtractionExpression } from './transcript-dom.js';
import { selectTranscriptPage as selectSharedTranscriptPage } from './transcript-pages.js';
import {
  buildTranscriptOutputBasePath,
  buildTranscriptOutputPayload,
  saveTranscriptOutputs,
  writeJsonOutput,
} from './transcript-output.js';

const cliRuntime = createCliRuntime({ appName: APP_NAME });
const {
  normalizeBrowserKey,
  ensureSupportedPlatform,
  createPrompt,
  chooseFromList,
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
 * @param {string[]} argv
 * @returns {CliOptions}
 */
export function parseDomModeArgs(argv) {
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

function printRunInstructions() {
  console.log('\nBrowser is ready.');
  console.log('1. Open the meeting in Microsoft Stream.');
  console.log('2. Open the Transcript panel.');
  console.log('3. Return to this terminal and continue.');
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

export async function runDomMode(argv = process.argv.slice(2), controls = {}) {
  const options = parseDomModeArgs(argv);
  const {
    printHelp = () => {},
    printVersion = () => {},
  } = controls;

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

    targetPage = await selectSharedTranscriptPage({
      prompt,
      debugPort,
      findPageTargets,
      chooseFromList,
    });
    console.log(`\nConnecting to: ${targetPage.title}`);

    cdp = await connectToPage(targetPage.webSocketDebuggerUrl);

    console.log('Scrolling the transcript and extracting entries...');
    extractionResult = await evaluate(
      cdp,
      buildTranscriptExtractionExpression({
        extractionSettings: EXTRACTION_SETTINGS,
        debugEnabled: options.debug,
      }),
      EXTRACTION_TIMEOUT_MS,
    );

    if (extractionResult.error) {
      if (options.debug) {
        const failureOutputBasePath = buildTranscriptOutputBasePath({
          defaultName: targetPage?.title || 'meeting',
          outputName: options.outputName,
          outputDir: options.outputDir,
          defaultOutputDir: DEFAULT_OUTPUT_DIR,
        });
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
        const debugPath = writeJsonOutput(
          `${failureOutputBasePath}.debug.json`,
          debugPayload,
        );
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
        const failureOutputBasePath = buildTranscriptOutputBasePath({
          defaultName: targetPage?.title || 'meeting',
          outputName: options.outputName,
          outputDir: options.outputDir,
          defaultOutputDir: DEFAULT_OUTPUT_DIR,
        });
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
        const debugPath = writeJsonOutput(
          `${failureOutputBasePath}.debug.json`,
          debugPayload,
        );
        console.log(`Saved debug output to: ${debugPath}`);
      }

      throw new CliError(
        'No transcript entries were found. Confirm the transcript panel is open and visible.',
      );
    }

    const metadata = await extractMeetingMetadata(cdp, evaluate);
    const outputPayload = buildTranscriptOutputPayload(metadata, entries);
    const outputBasePath = buildTranscriptOutputBasePath({
      defaultName: outputPayload.meeting.title,
      outputName: options.outputName,
      outputDir: options.outputDir,
      defaultOutputDir: DEFAULT_OUTPUT_DIR,
    });
    const outputPaths = saveTranscriptOutputs({
      payload: outputPayload,
      outputBasePath,
      outputFormat: options.outputFormat,
      renderMarkdown: (payload) => renderTranscriptMarkdown(payload),
    });
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
      debugPath = writeJsonOutput(`${outputBasePath}.debug.json`, debugPayload);
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
