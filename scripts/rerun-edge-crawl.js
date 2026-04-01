#!/usr/bin/env bun

import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { spawn } from 'node:child_process';

function printUsage() {
  console.log('Usage: bun ./scripts/rerun-edge-crawl.js [options] [-- <extra crawl args>]');
  console.log('');
  console.log('Options:');
  console.log('  --output-dir <path>  Crawl output directory to inspect. Default: ./output');
  console.log('  --browser <key>      Browser to use for the crawl. Default: edge');
  console.log('  --dry-run            Print the resolved artifacts and crawl command only.');
}

function parseArgs(argv) {
  const options = {
    outputDir: 'output',
    browser: 'edge',
    dryRun: false,
    extraArgs: [],
  };

  for (let index = 0; index < argv.length; index += 1) {
    const arg = String(argv[index] || '');

    if (arg === '--') {
      options.extraArgs = argv.slice(index + 1);
      break;
    }

    if (arg === '--dry-run') {
      options.dryRun = true;
      continue;
    }

    if (arg === '--output-dir') {
      const value = argv[index + 1];
      if (!value || String(value).startsWith('--')) {
        throw new Error('Missing value for --output-dir.');
      }
      options.outputDir = String(value);
      index += 1;
      continue;
    }

    if (arg === '--browser') {
      const value = argv[index + 1];
      if (!value || String(value).startsWith('--')) {
        throw new Error('Missing value for --browser.');
      }
      options.browser = String(value).trim().toLowerCase();
      index += 1;
      continue;
    }

    throw new Error(`Unknown option "${arg}".`);
  }

  return options;
}

function tryReadJson(path) {
  try {
    return JSON.parse(readFileSync(path, 'utf-8'));
  } catch {
    return null;
  }
}

function findLatestArtifact(outputDir, suffix) {
  if (!existsSync(outputDir)) {
    return '';
  }

  const name = readdirSync(outputDir)
    .filter((entry) => entry.endsWith(suffix))
    .sort()
    .at(-1);

  return name ? join(outputDir, name) : '';
}

function summarizeState(statePath) {
  const payload = tryReadJson(statePath);
  if (!payload) {
    return null;
  }

  const items = Array.isArray(payload.items) ? payload.items : [];
  const failedItems = items.filter(
    (item) => String(item?.lastStatus || '').toLowerCase() === 'failed',
  );

  return {
    path: statePath,
    summary: payload.summary || null,
    failedItems: failedItems.map((item) => ({
      title: item.title || item.url || 'Untitled meeting',
      url: item.url || '',
      error: item.lastError || '',
    })),
  };
}

function summarizeBatch(batchPath) {
  const payload = tryReadJson(batchPath);
  if (!payload) {
    return null;
  }

  const failedResults = Array.isArray(payload.results)
    ? payload.results.filter(
        (item) => String(item?.status || '').toLowerCase() === 'failed',
      )
    : [];

  return {
    path: batchPath,
    stateFilePath:
      String(payload?.state?.stateFilePath || payload?.options?.stateFilePath || ''),
    failedResults: failedResults.map((item) => ({
      title: item.title || item.url || 'Untitled meeting',
      url: item.url || '',
      error: item.error || '',
    })),
  };
}

function printSummary({ outputDir, stateSummary, batchSummary, command }) {
  console.log(`Output directory: ${outputDir}`);
  console.log(`Resolved command: ${command.join(' ')}`);

  if (stateSummary) {
    console.log(`State file: ${stateSummary.path}`);
    if (stateSummary.summary) {
      const summary = stateSummary.summary;
      console.log(
        `Queue summary: ${summary.totalItemCount || 0} total, ` +
          `${summary.failedItemCount || 0} failed, ` +
          `${summary.successItemCount || 0} done, ` +
          `${summary.pendingItemCount || 0} pending, ` +
          `${summary.newItemCount || 0} new.`,
      );
    }

    if (stateSummary.failedItems.length > 0) {
      console.log('Failed items from state:');
      stateSummary.failedItems.forEach((item, index) => {
        console.log(`  ${index + 1}. ${item.title}`);
      });
    }
  } else {
    console.log('State file: none found');
  }

  if (batchSummary) {
    console.log(`Batch file: ${batchSummary.path}`);
    if (batchSummary.stateFilePath) {
      console.log(`Batch-linked state: ${batchSummary.stateFilePath}`);
    }
    if (batchSummary.failedResults.length > 0) {
      console.log(`Failed items from batch: ${batchSummary.failedResults.length}`);
    }
  } else {
    console.log('Batch file: none found');
  }
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const outputDir = resolve(process.cwd(), options.outputDir);

  const stableStatePath = join(outputDir, 'crawl.state.json');
  const stableBatchPath = join(outputDir, 'crawl.batch.json');
  const statePath = existsSync(stableStatePath)
    ? stableStatePath
    : findLatestArtifact(outputDir, '.state.json');
  const batchPath = existsSync(stableBatchPath)
    ? stableBatchPath
    : findLatestArtifact(outputDir, '.batch.json');

  const command = [
    'bun',
    './cli.js',
    'crawl',
    '--browser',
    options.browser,
    '--output-dir',
    outputDir,
    ...options.extraArgs,
  ];

  printSummary({
    outputDir,
    stateSummary: statePath ? summarizeState(statePath) : null,
    batchSummary: batchPath ? summarizeBatch(batchPath) : null,
    command,
  });

  if (options.dryRun) {
    return;
  }

  await new Promise((resolvePromise, rejectPromise) => {
    const child = spawn(command[0], command.slice(1), {
      cwd: process.cwd(),
      stdio: 'inherit',
    });

    child.on('exit', (code) => {
      process.exitCode = code ?? 1;
      resolvePromise();
    });
    child.on('error', rejectPromise);
  });
}

try {
  await main();
} catch (error) {
  console.error(error instanceof Error ? error.message : String(error));
  printUsage();
  process.exitCode = 1;
}
