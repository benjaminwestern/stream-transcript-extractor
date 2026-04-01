import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';

export function buildTranscriptOutputPayload(
  metadata,
  entries,
  { extractedAt = new Date().toISOString() } = {},
) {
  const speakers = [
    ...new Set(entries.map((entry) => entry.speaker).filter(Boolean)),
  ];

  return {
    meeting: metadata,
    extractedAt,
    entryCount: entries.length,
    speakers,
    entries,
  };
}

export function sanitizeOutputFilename(
  value,
  { fallbackName = 'meeting', maxLength = 80 } = {},
) {
  const sanitized = String(value || '')
    .replace(/[^a-zA-Z0-9-_ ]/g, '')
    .replace(/\s+/g, '_')
    .slice(0, maxLength);

  return sanitized || fallbackName;
}

export function resolveOutputDirectory(
  outputDir,
  { cwd = process.cwd(), defaultOutputDir = 'output' } = {},
) {
  return resolve(cwd, outputDir || defaultOutputDir);
}

export function buildTranscriptOutputBasePath({
  defaultName,
  outputName,
  outputDir,
  cwd = process.cwd(),
  defaultOutputDir = 'output',
  timestamp = new Date().toISOString(),
} = {}) {
  const absoluteOutputDir = resolveOutputDirectory(outputDir, {
    cwd,
    defaultOutputDir,
  });
  mkdirSync(absoluteOutputDir, { recursive: true });

  const timestampLabel = String(timestamp)
    .replace(/[:.]/g, '-')
    .slice(0, 19);
  const baseName = sanitizeOutputFilename(outputName || defaultName || 'meeting');

  return join(absoluteOutputDir, `${baseName}_${timestampLabel}`);
}

export function writeJsonOutput(
  outputPath,
  payload,
  { ensureDirectory = false } = {},
) {
  if (ensureDirectory) {
    mkdirSync(dirname(outputPath), { recursive: true });
  }

  writeFileSync(outputPath, JSON.stringify(payload, null, 2));
  return outputPath;
}

export function writeTextOutput(
  outputPath,
  content,
  { ensureDirectory = false } = {},
) {
  if (ensureDirectory) {
    mkdirSync(dirname(outputPath), { recursive: true });
  }

  writeFileSync(outputPath, content);
  return outputPath;
}

export function saveTranscriptOutputs({
  payload,
  outputBasePath,
  outputFormat,
  renderMarkdown,
}) {
  const outputPaths = [];

  if (outputFormat === 'json' || outputFormat === 'both') {
    outputPaths.push(writeJsonOutput(`${outputBasePath}.json`, payload));
  }

  if (outputFormat === 'md' || outputFormat === 'both') {
    if (typeof renderMarkdown !== 'function') {
      throw new Error('Markdown output requires a renderMarkdown function.');
    }

    outputPaths.push(
      writeTextOutput(`${outputBasePath}.md`, renderMarkdown(payload)),
    );
  }

  return outputPaths;
}
