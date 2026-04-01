import { spawn, execSync } from 'node:child_process';
import {
  cpSync,
  existsSync,
  readFileSync,
  readdirSync,
  statSync,
  symlinkSync,
} from 'node:fs';
import { createServer } from 'node:net';
import { homedir, platform } from 'node:os';
import { join } from 'node:path';
import {
  connectCdp,
  getBrowserDebuggerWebSocketUrl,
  waitForBrowserDebugEndpointToClose,
} from './cdp.js';

const CURRENT_PLATFORM = platform();
const IS_WINDOWS = CURRENT_PLATFORM === 'win32';

function buildError(message, errorFactory) {
  return errorFactory ? errorFactory(message) : new Error(message);
}

function sleep(ms) {
  return new Promise((resolvePromise) => setTimeout(resolvePromise, ms));
}

export async function findAvailablePort({
  requestedPort,
  host = '127.0.0.1',
  errorFactory,
} = {}) {
  return new Promise((resolvePort, rejectPort) => {
    const server = createServer();

    server.unref();
    server.on('error', (error) => {
      const reason =
        requestedPort == null
          ? 'Unable to allocate a remote-debugging port.'
          : `Port ${requestedPort} is not available.`;
      rejectPort(buildError(`${reason} ${error.message}`, errorFactory));
    });

    server.listen(requestedPort ?? 0, host, () => {
      const address = server.address();
      if (address == null || typeof address === 'string') {
        server.close(() =>
          rejectPort(buildError('Unable to determine the debug port.', errorFactory)),
        );
        return;
      }

      const { port } = address;
      server.close((closeError) => {
        if (closeError) {
          rejectPort(buildError(closeError.message, errorFactory));
          return;
        }

        resolvePort(port);
      });
    });
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

function discoverProfiles(basePath) {
  const infoCache = readProfileInfoCache(basePath);
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

export function discoverBrowsers() {
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

export async function ensureBrowserIsClosed(prompt, browser) {
  while (isBrowserRunning(browser.processName)) {
    console.log(`\n${browser.name} is still running.`);
    console.log(
      'Close it so the extractor can relaunch the selected profile in debug mode.',
    );
    await prompt.waitForEnter('\nPress Enter after the browser is closed...\n');
  }
}

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

export async function selectBrowserAndProfile(
  options,
  prompt,
  {
    selectFromList,
    errorFactory,
  } = {},
) {
  const browsers = discoverBrowsers();

  if (browsers.length === 0) {
    throw buildError(
      'No supported Chrome or Edge profiles were found on this machine.',
      errorFactory,
    );
  }

  let browser = null;

  if (options.browser) {
    browser = browsers.find(
      (candidate) => candidate.key === options.browser,
    );
    if (!browser) {
      throw buildError(
        `Browser "${options.browser}" was not found. Available browsers: ` +
          `${browsers.map((candidate) => candidate.key).join(', ')}.`,
        errorFactory,
      );
    }
  } else if (browsers.length === 1) {
    browser = browsers[0];
    console.log(`Using ${browser.name}.`);
  } else {
    browser = await selectFromList(
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
      throw buildError(
        `Profile "${options.profile}" was not found for ${browser.name}.`,
        errorFactory,
      );
    }
  } else if (browser.profiles.length === 1) {
    profile = browser.profiles[0];
    console.log(`Using profile: ${profile.displayName}.`);
  } else {
    profile = await selectFromList(
      prompt,
      `${browser.name} profiles`,
      browser.profiles,
      (candidate) => candidate.displayName,
      '\nSelect profile (number): ',
    );
  }

  return { browser, profile };
}

export function prepareTempProfile(basePath, profile, tempDir) {
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

export function launchBrowser(browser, profile, tempDataDir, debugPort) {
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
  timeoutMs = 8_000,
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

async function closeBrowserViaCdp({
  debugPort,
  host = '127.0.0.1',
  defaultTimeoutMs = 30_000,
  errorFactory,
  shutdownTimeoutMs = 8_000,
} = {}) {
  const browserWebSocketUrl = await getBrowserDebuggerWebSocketUrl({
    host,
    port: debugPort,
  });
  if (!browserWebSocketUrl) {
    return false;
  }

  let browserCdp = null;

  try {
    browserCdp = await connectCdp(browserWebSocketUrl, {
      defaultTimeoutMs,
      errorFactory,
    });
    await browserCdp.send('Browser.close', {}, 5_000);
  } catch {
    return false;
  } finally {
    if (browserCdp) {
      browserCdp.close();
    }
  }

  return waitForBrowserDebugEndpointToClose({
    host,
    port: debugPort,
    timeoutMs: shutdownTimeoutMs,
  });
}

export async function shutdownBrowser({
  browserProcess,
  debugPort,
  host = '127.0.0.1',
  defaultTimeoutMs = 30_000,
  errorFactory,
  shutdownTimeoutMs = 8_000,
} = {}) {
  let closed = false;

  if (debugPort != null) {
    closed = await closeBrowserViaCdp({
      debugPort,
      host,
      defaultTimeoutMs,
      errorFactory,
      shutdownTimeoutMs,
    });
  }

  if (!closed && browserProcess?.pid) {
    stopProcess(browserProcess.pid);
  }

  if (browserProcess?.pid) {
    await waitForProcessExit(browserProcess.pid, shutdownTimeoutMs);
  }

  if (!closed && debugPort != null) {
    await waitForBrowserDebugEndpointToClose({
      host,
      port: debugPort,
      timeoutMs: shutdownTimeoutMs,
    });
  }
}
