import { createDecipheriv } from 'node:crypto';

export function createTranscriptNetworkRuntime({
  containsTranscriptSignal,
  formatTimestampValue,
  normalizeInlineText,
  stripHtmlTags,
  stripUtf8Bom,
  trimForPreview,
  tryParseJson,
  tryParseJsonLenient,
}) {
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

      if (
        lines.length === 0 ||
        lines[0] === 'WEBVTT' ||
        lines[0].startsWith('NOTE')
      ) {
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

    const kid = String(
      protectionData.Kid || protectionEnvelope.keyId || '',
    ).trim();
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

      const matches = collectProtectionKeysFromBody(
        source.body,
        source.url || '',
      );
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

    if (
      candidate?.parsedFormat === 'vtt' ||
      isEncryptedTranscriptUrl(candidate?.url || '')
    ) {
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

  return {
    buildProtectionKeyMap,
    finalizeBodyRecordForResponse,
    getFirstStringValue,
    getUrlSearchParam,
    isEncryptedTranscriptUrl,
    isUsableTranscriptCandidate,
    scoreTranscriptCandidate,
    summarizeCapturedBody,
  };
}
