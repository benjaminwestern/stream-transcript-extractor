function buildError(message, errorFactory) {
  return errorFactory ? errorFactory(message) : new Error(message);
}

function sleep(ms) {
  return new Promise((resolvePromise) => setTimeout(resolvePromise, ms));
}

export function createCdpEventScope(cdp) {
  const stopHandlers = [];

  return {
    on(method, handler) {
      const stop = cdp.on(method, handler);
      stopHandlers.push(stop);
      return stop;
    },
    stopAll() {
      while (stopHandlers.length > 0) {
        const stop = stopHandlers.pop();
        try {
          stop();
        } catch {
          // Ignore listener cleanup failures.
        }
      }
    },
  };
}

export async function connectCdp(
  websocketUrl,
  {
    defaultTimeoutMs = 30_000,
    connectionTimeoutMs = 10_000,
    errorFactory,
  } = {},
) {
  const websocket = new WebSocket(websocketUrl);
  const pendingRequests = new Map();
  const eventHandlers = new Map();
  let nextMessageId = 0;

  await new Promise((resolveConnection, rejectConnection) => {
    const timeout = setTimeout(() => {
      rejectConnection(buildError('WebSocket connection timeout.', errorFactory));
    }, connectionTimeoutMs);

    websocket.onopen = () => {
      clearTimeout(timeout);
      resolveConnection();
    };

    websocket.onerror = (event) => {
      clearTimeout(timeout);
      rejectConnection(
        buildError(
          `WebSocket connection failed: ${event.message || 'unknown error'}.`,
          errorFactory,
        ),
      );
    };
  });

  websocket.onmessage = (event) => {
    const payload = JSON.parse(event.data);

    if (payload.id && pendingRequests.has(payload.id)) {
      const { resolveRequest, rejectRequest, timeout } =
        pendingRequests.get(payload.id);
      clearTimeout(timeout);
      pendingRequests.delete(payload.id);

      if (payload.error) {
        rejectRequest(
          buildError(`CDP request failed: ${payload.error.message}`, errorFactory),
        );
        return;
      }

      resolveRequest(payload.result);
      return;
    }

    if (!payload.method) {
      return;
    }

    const handlers = eventHandlers.get(payload.method);
    if (!handlers || handlers.size === 0) {
      return;
    }

    for (const handler of handlers) {
      try {
        handler(payload.params || {});
      } catch {
        // Ignore individual event handler failures so capture can continue.
      }
    }
  };

  websocket.onclose = () => {
    for (const [requestId, request] of pendingRequests.entries()) {
      clearTimeout(request.timeout);
      request.rejectRequest(
        buildError(
          `CDP connection closed before request ${requestId} completed.`,
          errorFactory,
        ),
      );
      pendingRequests.delete(requestId);
    }
  };

  return {
    send(method, params = {}, timeoutMs = defaultTimeoutMs) {
      const messageId = nextMessageId + 1;
      nextMessageId = messageId;

      return new Promise((resolveRequest, rejectRequest) => {
        const timeout = setTimeout(() => {
          pendingRequests.delete(messageId);
          rejectRequest(
            buildError(`CDP ${method} timed out.`, errorFactory),
          );
        }, timeoutMs);

        pendingRequests.set(messageId, {
          resolveRequest,
          rejectRequest,
          timeout,
        });

        websocket.send(
          JSON.stringify({
            id: messageId,
            method,
            params,
          }),
        );
      });
    },
    on(method, handler) {
      if (!eventHandlers.has(method)) {
        eventHandlers.set(method, new Set());
      }

      const handlers = eventHandlers.get(method);
      handlers.add(handler);

      return () => {
        handlers.delete(handler);
        if (handlers.size === 0) {
          eventHandlers.delete(method);
        }
      };
    },
    close() {
      for (const request of pendingRequests.values()) {
        clearTimeout(request.timeout);
      }
      pendingRequests.clear();
      eventHandlers.clear();
      websocket.close();
    },
  };
}

export async function findPageTargets({
  host = '127.0.0.1',
  port,
  errorFactory,
} = {}) {
  const response = await fetch(`http://${host}:${port}/json/list`);

  if (!response.ok) {
    throw buildError(
      `CDP target discovery failed with ${response.status}.`,
      errorFactory,
    );
  }

  const targets = await response.json();
  return targets.filter(
    (target) => target.type === 'page' && !target.url.startsWith('chrome'),
  );
}

export async function evaluate(
  cdp,
  expression,
  { timeoutMs = 30_000, errorFactory } = {},
) {
  const result = await cdp.send(
    'Runtime.evaluate',
    {
      expression,
      returnByValue: true,
      awaitPromise: true,
    },
    timeoutMs,
  );

  if (result.exceptionDetails) {
    throw buildError(
      `Evaluation failed: ${result.exceptionDetails.text}`,
      errorFactory,
    );
  }

  return result.result.value;
}

export async function waitForBrowserDebugEndpoint({
  host = '127.0.0.1',
  port,
  timeoutMs = 15_000,
  pollMs = 500,
  errorFactory,
} = {}) {
  const deadline = Date.now() + timeoutMs;

  while (Date.now() < deadline) {
    try {
      const response = await fetch(`http://${host}:${port}/json/version`);
      if (response.ok) {
        return;
      }
    } catch {
      // Keep polling until the deadline.
    }

    await sleep(pollMs);
  }

  throw buildError('Browser failed to start in debug mode.', errorFactory);
}

export async function waitForBrowserPageTarget({
  host = '127.0.0.1',
  port,
  timeoutMs = 15_000,
  pollMs = 500,
  errorFactory,
} = {}) {
  const deadline = Date.now() + timeoutMs;

  while (Date.now() < deadline) {
    try {
      const pages = await findPageTargets({ host, port, errorFactory });
      if (pages.length > 0) {
        return pages;
      }
    } catch {
      // Keep polling until the deadline.
    }

    await sleep(pollMs);
  }

  throw buildError(
    'Browser launched in debug mode but no page window became available.',
    errorFactory,
  );
}

export async function isBrowserDebugEndpointAvailable({
  host = '127.0.0.1',
  port,
} = {}) {
  try {
    const response = await fetch(`http://${host}:${port}/json/version`);
    return response.ok;
  } catch {
    return false;
  }
}

export async function waitForBrowserDebugEndpointToClose({
  host = '127.0.0.1',
  port,
  timeoutMs = 8_000,
  pollMs = 250,
} = {}) {
  const deadline = Date.now() + timeoutMs;

  while (Date.now() < deadline) {
    const isAvailable = await isBrowserDebugEndpointAvailable({ host, port });
    if (!isAvailable) {
      return true;
    }

    await sleep(pollMs);
  }

  return !(await isBrowserDebugEndpointAvailable({ host, port }));
}

export async function getBrowserDebuggerWebSocketUrl({
  host = '127.0.0.1',
  port,
} = {}) {
  try {
    const response = await fetch(`http://${host}:${port}/json/version`);
    if (!response.ok) {
      return '';
    }

    const payload = await response.json();
    return String(payload.webSocketDebuggerUrl || '');
  } catch {
    return '';
  }
}

export async function navigatePageAndWait(
  cdp,
  url,
  {
    timeoutMs = 45_000,
    settleMs = 1_000,
    errorFactory,
  } = {},
) {
  await cdp.send('Page.enable');

  await new Promise((resolveNavigation, rejectNavigation) => {
    let settled = false;
    let timeout = null;
    let stopLoadListener = null;

    function finish(callback) {
      if (settled) {
        return;
      }

      settled = true;
      if (timeout) {
        clearTimeout(timeout);
      }
      if (stopLoadListener) {
        stopLoadListener();
      }
      callback();
    }

    stopLoadListener = cdp.on('Page.loadEventFired', () => {
      finish(resolveNavigation);
    });

    timeout = setTimeout(() => {
      finish(() =>
        rejectNavigation(
          buildError(`Navigation to "${url}" timed out.`, errorFactory),
        ),
      );
    }, timeoutMs);

    cdp
      .send('Page.navigate', { url }, timeoutMs)
      .then((result) => {
        if (result?.errorText) {
          finish(() =>
            rejectNavigation(
              buildError(
                `Navigation to "${url}" failed: ${result.errorText}`,
                errorFactory,
              ),
            ),
          );
        }
      })
      .catch((error) => {
        finish(() => rejectNavigation(error));
      });
  });

  await sleep(settleMs);
}

export async function reloadPageAndWait(
  cdp,
  {
    timeoutMs = 30_000,
    settleMs = 1_000,
  } = {},
) {
  await cdp.send('Page.enable');

  await new Promise((resolveReload, rejectReload) => {
    let settled = false;
    let timeout = null;
    let stopLoadListener = null;

    function finish(callback) {
      if (settled) {
        return;
      }

      settled = true;
      if (timeout) {
        clearTimeout(timeout);
      }
      if (stopLoadListener) {
        stopLoadListener();
      }
      callback();
    }

    stopLoadListener = cdp.on('Page.loadEventFired', () => {
      finish(resolveReload);
    });

    timeout = setTimeout(() => {
      finish(resolveReload);
    }, timeoutMs);

    cdp.send('Page.reload', { ignoreCache: true }).catch((error) => {
      finish(() => rejectReload(error));
    });
  });

  await sleep(settleMs);
}
