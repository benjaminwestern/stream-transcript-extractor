#!/usr/bin/env bun

import { spawnSync } from 'node:child_process';
import { dirname } from 'node:path';
import { printHelpScreen } from './lib/cli.js';
import { discoverBrowsers } from './lib/browser-session.js';
import {
  APP_DESCRIPTION,
  APP_NAME,
  BUILD_TIME,
  BUILD_VERSION,
  DEFAULT_CRAWL_START_URL,
  DEFAULT_EXTRACTOR_MODE,
  DEFAULT_TRANSCRIPT_OUTPUT_FORMAT,
  DEFAULT_WORKFLOW,
  IS_MACOS,
  IS_WINDOWS,
  SUPPORTED_EXTRACTOR_MODES,
  SUPPORTED_WORKFLOWS,
} from './lib/extractor-config.js';
import { CliError, createCliRuntime } from './lib/extractor-cli-runtime.js';
import { runDomMode } from './lib/dom-mode-runtime.js';
import {
  runCrawler,
  runEmbeddedNetworkMode,
} from './lib/network-mode-runtime.js';

const { createPrompt, chooseFromList } = createCliRuntime({ appName: APP_NAME });
const TERMINAL_HANDOFF_ENV = 'STREAM_TRANSCRIPT_EXTRACTOR_TERMINAL_HANDOFF';

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

function printEntrypointHelp() {
  printHelpScreen({
    name: APP_NAME,
    summary: APP_DESCRIPTION,
    usage: [
      'bun ./cli.js [options]',
      'bun ./cli.js crawl [options]',
      './stream-transcript-extractor-<target> [workflow] [options]',
    ],
    sections: [
      {
        title: 'Start here',
        rows: [
          {
            label: 'bun ./cli.js',
            description:
              'Open the interactive launcher in a real terminal. The recommended path starts crawl with default settings and only asks you to choose a browser. The custom path keeps the full workflow menu.',
          },
          {
            label: './stream-transcript-extractor-<target>',
            description:
              'Open the same interactive launcher from a compiled binary in a real terminal.',
          },
          {
            label: 'bun ./cli.js crawl',
            description:
              'Open Stream home, switch to Meetings, scroll the rendered list, merge the results into a persistent queue, then run automatic extraction against the actionable queue in the same browser session.',
          },
        ],
      },
      {
        title: 'Workflows',
        rows: [
          {
            label: 'crawl',
            description:
              'Batch flow. Opens Stream home, selects Meetings, scrolls the visible page, merges results into a *.state.json queue, then wraps automatic extraction across new and failed items by default when resuming.',
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
          'The interactive recommended path launches crawl with the pending queue. Failed items are retried first, then newly discovered items, then any remaining pending items.',
          'Run `bun ./cli.js --mode automatic --help` or `bun ./cli.js crawl --help` for workflow-specific help.',
        ],
      },
      {
        title: 'Examples',
        lines: [
          'bun ./cli.js',
          'bun ./cli.js --browser chrome --profile Work',
          'bun ./cli.js crawl',
          'bun ./cli.js crawl --state-file ./exports/team.state.json',
          'bun ./cli.js --mode network --output-dir ./exports --format both',
          'bun ./cli.js --mode dom --debug',
        ],
      },
    ],
  });
}

function printEntrypointVersion() {
  const buildSuffix = BUILD_TIME ? ` (${BUILD_TIME})` : '';
  console.log(`${APP_NAME} ${BUILD_VERSION}${buildSuffix}`);
}

function isInteractiveEntrypointLaunch(argv) {
  return (
    Array.isArray(argv) &&
    argv.length === 0 &&
    Boolean(process.stdin.isTTY) &&
    Boolean(process.stdout.isTTY)
  );
}

function isCompiledBinaryEntrypoint() {
  const entrypoint = String(process.argv[1] || '').trim().toLowerCase();
  return Boolean(
    entrypoint &&
      !entrypoint.endsWith('.js') &&
      !entrypoint.endsWith('.mjs') &&
      !entrypoint.endsWith('.cjs'),
  );
}

function escapeShellArg(value) {
  return `'${String(value).replace(/'/g, `'\\''`)}'`;
}

function escapeAppleScriptString(value) {
  return String(value).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
}

function quoteWindowsCmdArg(value) {
  return `"${String(value).replace(/"/g, '""')}"`;
}

function handoffBinaryLaunchToMacTerminal(binaryPath) {
  const command = [
    `cd ${escapeShellArg(dirname(binaryPath))}`,
    `${TERMINAL_HANDOFF_ENV}=1 ${escapeShellArg(binaryPath)}`,
  ].join('; ');
  const appleScript = [
    'tell application "Terminal"',
    'activate',
    `do script "${escapeAppleScriptString(command)}"`,
    'end tell',
  ].join('\n');

  return spawnSync('/usr/bin/osascript', ['-e', appleScript], {
    stdio: 'ignore',
  }).status;
}

function handoffBinaryLaunchToWindowsConsole(binaryPath) {
  const command = [
    `set ${TERMINAL_HANDOFF_ENV}=1`,
    `start "" /d ${quoteWindowsCmdArg(dirname(binaryPath))} ${quoteWindowsCmdArg(binaryPath)}`,
  ].join(' && ');

  return spawnSync(process.env.ComSpec || 'cmd.exe', ['/d', '/c', command], {
    stdio: 'ignore',
  }).status;
}

function handleNoTtyBinaryLaunch(argv) {
  if (
    (!IS_MACOS && !IS_WINDOWS) ||
    !Array.isArray(argv) ||
    argv.length > 0 ||
    process.env[TERMINAL_HANDOFF_ENV] === '1' ||
    isInteractiveEntrypointLaunch(argv) ||
    !isCompiledBinaryEntrypoint()
  ) {
    return null;
  }

  const binaryPath = String(process.argv[1] || process.execPath || '').trim();
  if (!binaryPath) {
    return null;
  }

  const handoffStatus = IS_MACOS
    ? handoffBinaryLaunchToMacTerminal(binaryPath)
    : handoffBinaryLaunchToWindowsConsole(binaryPath);

  if (handoffStatus === 0) {
    return 0;
  }

  console.error(
    'This binary needs a terminal window. Run it from Terminal and try again.',
  );
  return 1;
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

function getDetectedBrowserChoices() {
  const browsers = discoverBrowsers();

  if (browsers.length === 0) {
    throw new CliError('No supported Chrome or Edge profiles were found on this machine.');
  }

  return browsers.map((browser) => ({
    value: browser.key,
    label: browser.name,
    description:
      browser.profiles.length === 1
        ? '1 local profile detected.'
        : `${browser.profiles.length} local profiles detected.`,
  }));
}

async function promptForRecommendedLaunchArgs(prompt) {
  console.log('\nRecommended settings');
  console.log(
    'Run the crawl workflow with the default settings and automatically process the actionable queue.',
  );

  const browserChoices = getDetectedBrowserChoices();
  let browserKey = browserChoices[0]?.value || '';

  if (browserChoices.length === 1) {
    console.log(`Using ${browserChoices[0].label}.`);
  } else {
    const browserChoice = await chooseInteractiveMenuChoice(
      prompt,
      'Browser',
      '\nChoose a browser (number): ',
      browserChoices,
    );
    browserKey = browserChoice.value;
  }

  return ['crawl', '--browser', browserKey, '--select', 'pending'];
}

async function promptForCustomLaunchArgs(prompt) {
  console.log('\nCustom settings');
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
}

async function promptForInteractiveLaunchArgs() {
  const prompt = createPrompt();

  try {
    console.log('\nInteractive launch');
    console.log(
      'Choose the fast recommended crawl path or open the full custom settings flow.',
    );

    const launchChoice = await chooseInteractiveMenuChoice(
      prompt,
      'Launch mode',
      '\nChoose launch mode (number): ',
      [
        {
          value: 'recommended',
          label: 'recommended settings',
          description:
            'Run the crawl workflow with the default settings, automatically process the pending queue, and only ask you to choose the browser up front.',
        },
        {
          value: 'custom',
          label: 'custom settings',
          description:
            'Open the existing full launcher so you can choose workflow, formats, diagnostics, and overrides.',
        },
      ],
    );

    if (launchChoice.value === 'recommended') {
      return await promptForRecommendedLaunchArgs(prompt);
    }

    return await promptForCustomLaunchArgs(prompt);
  } finally {
    prompt.close();
  }
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

async function runSelectedWorkflow(argv = process.argv.slice(2)) {
  let rawArgv = [...argv];
  const noTtyLaunchResult = handleNoTtyBinaryLaunch(rawArgv);

  if (noTtyLaunchResult != null) {
    return noTtyLaunchResult;
  }

  if (isInteractiveEntrypointLaunch(rawArgv)) {
    rawArgv = await promptForInteractiveLaunchArgs();
  }

  const { workflow, mode, modeProvided, forwardedArgs } = resolveWorkflowSelection(
    rawArgv,
  );
  if (workflow === 'crawl') {
    return runCrawler(forwardedArgs);
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
    return runDomMode(forwardedArgs, {
      printHelp: printEntrypointHelp,
      printVersion: printEntrypointVersion,
    });
  }

  if (mode === 'automatic') {
    return runEmbeddedNetworkMode('automatic', forwardedArgs);
  }

  return runEmbeddedNetworkMode('manual', forwardedArgs);
}

if (import.meta.main) {
  let exitCode = 0;

  try {
    exitCode = await runSelectedWorkflow(process.argv.slice(2));
  } catch (error) {
    exitCode = handleEntrypointError(error);
  }

  process.exitCode = exitCode;
}
