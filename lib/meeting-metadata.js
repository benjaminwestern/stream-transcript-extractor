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

export function buildMeetingMetadataExtractionExpression() {
  return `
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
        date: dateMatch?.[1]?.trim() || dateMatch?.[0]?.trim() || '',
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
  `;
}

export async function extractMeetingMetadata(cdp, evaluate) {
  return evaluate(cdp, buildMeetingMetadataExtractionExpression());
}

export function formatIsoDateForDisplay(value) {
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

export function formatIsoDateTimeForDisplay(value) {
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

function isDefaultStreamRelatedUrl(url) {
  return /stream|microsoftstream|m365\.cloud\.microsoft|substrate\.office\.com|officeapps\.live\.com|office\.com|office365\.com|office\.net|sharepoint\.com|teams\.microsoft\.com|graph\.microsoft\.com|meetingcatchupportal|cortana\.ai/.test(
    String(url || '').toLowerCase(),
  );
}

export function extractMeetingMetadataFromCapture(
  captureResult,
  { isStreamRelatedUrl = isDefaultStreamRelatedUrl } = {},
) {
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

export function mergeMeetingMetadata(pageMetadata, captureMetadata, targetPage) {
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

export function renderTranscriptMarkdown(
  payload,
  { includeExtendedMetadata = false } = {},
) {
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
    includeExtendedMetadata && startDateTime
      ? `Start date/time: ${startDateTime}`
      : '',
    includeExtendedMetadata && endDateTime
      ? `End date/time: ${endDateTime}`
      : '',
    includeExtendedMetadata && createdByLabel
      ? `Created by: ${createdByLabel}`
      : '',
    payload.meeting.recordedBy ? `Recorded by: ${payload.meeting.recordedBy}` : '',
    includeExtendedMetadata && speakerList ? `Speakers: ${speakerList}` : '',
    includeExtendedMetadata && payload.meeting.sourceUrl
      ? `Source URL: ${payload.meeting.sourceUrl}`
      : '',
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
