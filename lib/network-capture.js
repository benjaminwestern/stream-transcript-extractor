export function createResponseBodyRecord(
  {
    body = '',
    bodyError = '',
    base64Encoded = false,
    rawBodyBase64 = '',
    rawBodyByteLength = 0,
    rawBodyPreviewHex = '',
    refetchedExternally = false,
    refetchError = '',
  } = {},
  { trimForPreview },
) {
  const normalizedBody = String(body || '');

  return {
    body: normalizedBody,
    bodyLength: normalizedBody.length,
    bodyPreview: trimForPreview ? trimForPreview(normalizedBody) : normalizedBody,
    bodyError,
    base64Encoded,
    rawBodyBase64,
    rawBodyByteLength,
    rawBodyPreviewHex,
    refetchedExternally,
    refetchError,
  };
}

export async function loadResponseBody(
  cdp,
  requestId,
  bodyCache,
  { trimForPreview },
) {
  if (bodyCache.has(requestId)) {
    return bodyCache.get(requestId);
  }

  let body = '';
  let bodyError = '';
  let base64Encoded = false;
  let rawBodyBase64 = '';
  let rawBodyByteLength = 0;
  let rawBodyPreviewHex = '';

  try {
    const result = await cdp.send('Network.getResponseBody', { requestId });
    base64Encoded = Boolean(result.base64Encoded);
    rawBodyBase64 = base64Encoded ? result.body : '';

    const rawBytes = base64Encoded
      ? Buffer.from(result.body, 'base64')
      : Buffer.from(result.body, 'utf8');
    rawBodyByteLength = rawBytes.length;
    rawBodyPreviewHex = rawBytes.subarray(0, 32).toString('hex');
    body = base64Encoded ? rawBytes.toString('utf8') : result.body;
  } catch (error) {
    bodyError =
      error instanceof Error ? error.message : 'Unable to capture response body.';
  }

  const record = createResponseBodyRecord(
    {
      body,
      bodyError,
      base64Encoded,
      rawBodyBase64,
      rawBodyByteLength,
      rawBodyPreviewHex,
      refetchedExternally: false,
      refetchError: '',
    },
    { trimForPreview },
  );
  bodyCache.set(requestId, record);
  return record;
}
