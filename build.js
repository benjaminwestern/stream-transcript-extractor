#!/usr/bin/env bun

import { mkdirSync, statSync } from 'node:fs';

const DIST_DIR = 'dist';
const DEFAULT_BUILD_VERSION = '0.0.0-dev';
const BINARY = Object.freeze({
  name: 'extractor',
  entrypoint: './extract.js',
  baseName: 'stream-transcript-extractor',
});

const TARGETS = Object.freeze([
  {
    name: 'macos-arm64',
    target: 'bun-darwin-arm64',
    suffix: 'macos-arm64',
  },
  {
    name: 'macos-x64',
    target: 'bun-darwin-x64',
    suffix: 'macos-x64',
  },
  {
    name: 'windows-x64',
    target: 'bun-windows-x64-baseline',
    suffix: 'windows-x64.exe',
  },
]);

function formatBytes(bytes) {
  const units = ['B', 'KB', 'MB', 'GB'];
  let value = bytes;
  let unitIndex = 0;

  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }

  return `${value.toFixed(unitIndex === 0 ? 0 : 1)} ${units[unitIndex]}`;
}

function readBuildVersion() {
  if (process.env.BUILD_VERSION) {
    return String(process.env.BUILD_VERSION);
  }

  return DEFAULT_BUILD_VERSION;
}

function printHelp() {
  const targetNames = TARGETS.map((target) => target.name).join(', ');

  console.log('Stream Transcript Extractor build script');
  console.log('');
  console.log('Usage:');
  console.log('  bun ./build.js');
  console.log('  bun ./build.js <target> [<target> ...]');
  console.log('');
  console.log('Targets:');
  console.log(`  ${targetNames}`);
  console.log('');
  console.log('Binary:');
  console.log(`  ${BINARY.baseName}`);
  console.log('  Run the interactive launcher with: <binary>');
  console.log('  Run crawl with: <binary> crawl [options]');
  console.log('');
  console.log('Environment:');
  console.log(
    `  BUILD_VERSION  Override the embedded version (default: ${DEFAULT_BUILD_VERSION}).`,
  );
}

async function main() {
  const rawArgs = process.argv.slice(2);

  if (rawArgs.includes('--help') || rawArgs.includes('-h')) {
    printHelp();
    return;
  }

  const requestedTargets = new Set(rawArgs);
  const selectedTargets =
    requestedTargets.size === 0
      ? TARGETS
      : TARGETS.filter((target) => requestedTargets.has(target.name));

  if (requestedTargets.size > 0 && selectedTargets.length === 0) {
    console.error(
      `Unknown build target. Use one of: ${TARGETS.map((target) => target.name).join(', ')}`,
    );
    process.exitCode = 1;
    return;
  }

  const buildVersion = readBuildVersion();
  const buildTime = new Date().toISOString();

  mkdirSync(DIST_DIR, { recursive: true });

  console.log(`Building ${selectedTargets.length} binary target(s)...`);

  for (const buildTarget of selectedTargets) {
    console.log(`\n- ${buildTarget.name} (${buildTarget.target})`);

    const outfile = `${DIST_DIR}/${BINARY.baseName}-${buildTarget.suffix}`;
    const result = await Bun.build({
      entrypoints: [BINARY.entrypoint],
      minify: true,
      compile: {
        target: buildTarget.target,
        outfile,
      },
      define: {
        __BUILD_VERSION__: JSON.stringify(buildVersion),
        __BUILD_TIME__: JSON.stringify(buildTime),
      },
    });

    if (!result.success) {
      console.error(`Build failed for ${BINARY.name} / ${buildTarget.name}.`);
      for (const log of result.logs) {
        console.error(log.message);
      }
      process.exitCode = 1;
      return;
    }

    const fileSize = statSync(outfile).size;
    console.log(`  Created ${outfile} (${formatBytes(fileSize)})`);
  }
}

await main();
