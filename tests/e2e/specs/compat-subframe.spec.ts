import { expect, test, type Page, type Route } from "playwright/test";

const expectedFacadeSteps = [
  "fixed slots and iframe URLs",
  "strict source origin instance session",
  "structured clone input wire",
  "typed resource loading state",
  "facade actions and callbacks",
  "callback error and timeout",
  "destroy rejects pending RPC",
];

const fakeChildHtml = String.raw`<!doctype html>
<meta charset="utf-8">
<script>
(() => {
  const SOURCE = "onlyoffice-embed-sdk/compat-subframe/v1";
  const VERSION = 1;
  const query = new URLSearchParams(location.search);
  const instanceId = query.get("instance");
  const sessionToken = query.get("session");
  const parentOrigin = query.get("parentOrigin");
  const readyDelay = Number(query.get("readyDelay") || 0);
  const pendingCallbacks = new Map();
  let lastOpen = null;
  let state = null;
  let identity = {
    packageVersion: "test",
    hostBuildId: "fake-child",
    assetManifestDigest: "a".repeat(64),
  };

  const envelope = () => ({
    source: SOURCE,
    version: VERSION,
    instanceId,
    sessionToken,
  });
  const send = (message) => parent.postMessage(message, parentOrigin);
  const respond = (request, ok, result, error) => send(ok
    ? { ...envelope(), type: "response", requestId: request.requestId, ok: true, result }
    : {
        ...envelope(),
        type: "response",
        requestId: request.requestId,
        ok: false,
        error: { name: error?.name || "Error", message: error?.message || String(error) },
      });
  const postState = () => send({
    ...envelope(),
    type: "event",
    event: "state-change",
    payload: state,
  });
  const postLoading = (payload) => send({
    ...envelope(),
    type: "event",
    event: "loading-change",
    payload,
  });
  const bytesOf = async (value) => {
    if (value instanceof Blob) return Array.from(new Uint8Array(await value.arrayBuffer()));
    if (value instanceof ArrayBuffer) return Array.from(new Uint8Array(value));
    if (ArrayBuffer.isView(value)) {
      return Array.from(new Uint8Array(value.buffer, value.byteOffset, value.byteLength));
    }
    return [];
  };
  const inspectDocument = async (document) => {
    if (document.kind === "empty") {
      return { kind: "empty", fileName: document.fileName, emptyType: document.emptyType };
    }
    const value = document.kind === "file" ? document.file : document.buffer;
    return {
      kind: document.kind,
      fileName: document.fileName,
      sourceKind: document.sourceKind,
      bytes: await bytesOf(value),
      constructorName: value?.constructor?.name,
    };
  };
  const callback = (name, file) => new Promise((resolve, reject) => {
    const callbackRequestId = "fake-callback-" + crypto.randomUUID();
    pendingCallbacks.set(callbackRequestId, { resolve, reject });
    send({
      ...envelope(),
      type: "callback-request",
      callbackRequestId,
      callback: name,
      payload: { file, fileName: file.name },
    });
  });
  const output = (targetExt) => new File(
    [Uint8Array.from([91, 92, 93])],
    "output." + (targetExt || "docx"),
    { type: "application/octet-stream" },
  );

  addEventListener("message", (event) => {
    if (event.source !== parent || event.origin !== parentOrigin) return;
    const message = event.data;
    if (
      !message ||
      message.source !== SOURCE ||
      message.version !== VERSION ||
      message.instanceId !== instanceId ||
      message.sessionToken !== sessionToken
    ) return;

    if (message.type === "callback-response") {
      const pending = pendingCallbacks.get(message.callbackRequestId);
      if (!pending) return;
      pendingCallbacks.delete(message.callbackRequestId);
      if (message.ok) pending.resolve(message.result);
      else pending.reject(Object.assign(new Error(message.error?.message || "callback failed"), {
        name: message.error?.name || "Error",
      }));
      return;
    }
    if (message.type !== "request") return;

    void (async () => {
      const payload = message.payload || {};
      if (message.action === "open") {
        postLoading({
          loading: true,
          phase: "static-resources",
          resourceStatus: "downloading",
          resourceDownload: true,
          transferredBytes: 4096,
          resourceCount: 1,
        });
        lastOpen = {
          ...(await inspectDocument(payload.document)),
          resourceOrigin: payload.resourceOrigin,
        };
        const fileName = payload.document.fileName;
        const fileType = fileName.split(".").pop() || "docx";
        state = {
          id: instanceId,
          fileName,
          fileType,
          mode: payload.mode,
          readonly: payload.readonly,
          dirty: false,
          sourceKind: payload.document.kind === "file"
            ? "local-file"
            : payload.document.kind === "empty"
              ? "new-document"
              : payload.document.sourceKind || "buffer",
          status: "ready",
          destroyed: false,
        };
        respond(message, true, { state, hostIdentity: identity });
        postLoading({
          loading: false,
          phase: "ready",
          resourceStatus: "downloaded",
          resourceDownload: false,
          transferredBytes: 4096,
          resourceCount: 1,
        });
        return;
      }
      if (message.action === "invoke-plugin") {
        if (payload.pluginGuid === "__pending__") return;
        if (payload.pluginGuid === "__mark_dirty__") {
          state = { ...state, dirty: true };
          postState();
          respond(message, true, null);
          return;
        }
        if (payload.pluginGuid === "__cache_hit__") {
          postLoading({
            loading: false,
            phase: "static-resources",
            resourceStatus: "cache-hit",
            resourceDownload: false,
            transferredBytes: 0,
            resourceCount: 0,
          });
          respond(message, true, null);
          return;
        }
        respond(message, true, lastOpen);
        return;
      }
      if (message.action === "get-state") {
        respond(message, true, state);
        return;
      }
      if (message.action === "get-host-identity") {
        respond(message, true, identity);
        return;
      }
      if (message.action === "set-readonly") {
        state = { ...state, readonly: payload.readonly, mode: payload.readonly ? "readonly" : "edit" };
        postState();
        respond(message, true, state);
        return;
      }
      if (message.action === "set-theme" || message.action === "set-language") {
        postState();
        respond(message, true, state);
        return;
      }
      if (message.action === "confirm-save-to-new-format") {
        respond(message, true, true);
        return;
      }
      if (["save", "save-as", "download"].includes(message.action)) {
        const file = output(payload.targetExt);
        const callbackName = message.action;
        try {
          await callback(callbackName, file);
          respond(message, true, file);
        } catch (error) {
          respond(message, false, undefined, error);
        }
        return;
      }
      if (message.action === "print") {
        respond(message, true, new File(
          [new TextEncoder().encode("%PDF-1.7\n% compat subframe regression\n")],
          "output.pdf",
          { type: "application/pdf" },
        ));
        return;
      }
      if (message.action === "destroy") {
        state = state ? { ...state, status: "destroyed", destroyed: true } : state;
        respond(message, true, null);
        return;
      }
      respond(message, false, undefined, new Error("unsupported action"));
    })().catch((error) => respond(message, false, undefined, error));
  });

  setTimeout(() => send({ ...envelope(), type: "ready" }), readyDelay);
})();
</script>`;

async function serveFakeChild(route: Route) {
  await route.fulfill({
    status: 200,
    contentType: "text/html; charset=utf-8",
    headers: { "Origin-Agent-Cluster": "?1" },
    body: fakeChildHtml,
  });
}

async function installFakeChild(page: Page) {
  await page.route(
    /^http:\/\/[a-z]+\.onlyoffice\.localhost:\d+\/subframe(?:\?|$)/,
    serveFakeChild,
  );
  await page.route("**/compat-fixture.docx", (route) =>
    route.fulfill({
      status: 200,
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      body: Buffer.from([21, 22, 23]),
    }),
  );
}

test("package facade enforces the compatibility transport contract", async ({ page }) => {
  await installFakeChild(page);
  await page.goto("/e2e/runtime-regressions?harness=compat-subframe", {
    waitUntil: "domcontentloaded",
  });

  await expect(page.getByTestId("compat-facade-status")).toHaveText(
    /passed|failed/,
    { timeout: 30_000 },
  );
  const status = await page.getByTestId("compat-facade-status").innerText();
  const steps = JSON.parse(
    await page.getByTestId("compat-facade-result").innerText(),
  ) as Array<{ name: string; status: string; detail?: string }>;

  expect(steps, JSON.stringify(steps, null, 2)).toEqual(
    expectedFacadeSteps.map((name) => ({ name, status: "passed" })),
  );
  expect(status).toBe("passed");
});

test("real compatibility subframe fails closed without an exact parent origin", async ({ page }) => {
  const response = await page.goto(
    "/subframe?runtime=compat&instance=missing-parent&session=missing-parent",
    { waitUntil: "domcontentloaded" },
  );
  expect(response?.headers()["origin-agent-cluster"]).toBe("?1");
  await expect(
    page.locator('[data-onlyoffice-subframe-runtime="invalid"]'),
  ).toBeVisible();
});

test("hosted cache policy separates documents from immutable resources", async ({
  request,
}) => {
  const subframe = await request.get("/subframe?runtime=compat");
  const version = await request.get("/api/version");
  const immutableAsset = await request.get(
    "/packages/onlyoffice/9.4.0-develop/web-apps/apps/api/documents/api.js",
  );

  expect(subframe.headers()["cache-control"]).toBe("no-store");
  expect(version.headers()["cache-control"]).toBe("no-store");
  expect(immutableAsset.headers()["cache-control"]).toContain("immutable");
  expect(immutableAsset.headers()["timing-allow-origin"]).toBe("*");
  expect(immutableAsset.headers()["vary"] ?? "").not.toMatch(/origin/i);
});

test("real compatibility subframe performs a strict ready handshake", async ({ page }) => {
  await page.goto("/", { waitUntil: "domcontentloaded" });
  const parentOrigin = new URL(page.url()).origin;
  const port = new URL(page.url()).port;
  const childOrigin = `http://rat.onlyoffice.localhost:${port}`;
  const responsePromise = page.waitForResponse((response) => {
    const url = new URL(response.url());
    return url.origin === childOrigin && url.pathname === "/subframe";
  });

  const ready = await page.evaluate(
    async ({ childOrigin, parentOrigin }) => {
      const instanceId = "real-handshake";
      const sessionToken = "real-handshake-session";
      const frame = document.createElement("iframe");
      frame.name = "real-compat-child";
      const url = new URL("/subframe", childOrigin);
      url.searchParams.set("runtime", "compat");
      url.searchParams.set("instance", instanceId);
      url.searchParams.set("session", sessionToken);
      url.searchParams.set("parentOrigin", parentOrigin);
      frame.src = url.href;
      document.body.appendChild(frame);

      return await new Promise<Record<string, unknown>>((resolve, reject) => {
        const timer = window.setTimeout(() => reject(new Error("real subframe ready timed out")), 10_000);
        const listener = (event: MessageEvent) => {
          const message = event.data as Record<string, unknown> | null;
          if (
            event.source !== frame.contentWindow ||
            event.origin !== childOrigin ||
            message?.source !== "onlyoffice-embed-sdk/compat-subframe/v1" ||
            message.type !== "ready" ||
            message.instanceId !== instanceId ||
            message.sessionToken !== sessionToken
          ) return;
          window.clearTimeout(timer);
          window.removeEventListener("message", listener);
          resolve(message);
        };
        window.addEventListener("message", listener);
      });
    },
    { childOrigin, parentOrigin },
  );
  const response = await responsePromise;

  expect(response.headers()["origin-agent-cluster"]).toBe("?1");
  expect(ready.version).toBe(1);
  const child = page.frames().find((frame) => frame.name() === "real-compat-child");
  expect(child).toBeDefined();
  await expect(
    child!.locator('[data-onlyoffice-subframe-runtime="compat"]'),
  ).toHaveAttribute("data-onlyoffice-subframe-protocol", "1");

  const strictResult = await page.evaluate(
    async ({ childOrigin }) => {
      const frame = document.querySelector<HTMLIFrameElement>('iframe[name="real-compat-child"]')!;
      const base = {
        source: "onlyoffice-embed-sdk/compat-subframe/v1",
        version: 1,
        type: "request",
        instanceId: "real-handshake",
        sessionToken: "real-handshake-session",
        action: "get-state",
      };
      let wrongSessionResponse = false;
      const listener = (event: MessageEvent) => {
        if (event.data?.requestId === "wrong-session") wrongSessionResponse = true;
      };
      window.addEventListener("message", listener);
      frame.contentWindow!.postMessage(
        { ...base, requestId: "wrong-session", sessionToken: "wrong" },
        childOrigin,
      );
      await new Promise((resolve) => window.setTimeout(resolve, 100));
      window.removeEventListener("message", listener);
      return { wrongSessionResponse };
    },
    { childOrigin },
  );
  expect(strictResult.wrongSessionResponse).toBe(false);
});

test("all 12 real fixed origins expose the exact compatibility handshake", async ({ page }) => {
  test.setTimeout(120_000);
  const slots = [
    "rat", "ox", "tiger", "rabbit", "dragon", "snake",
    "horse", "goat", "monkey", "rooster", "dog", "pig",
  ];
  await page.goto("/", { waitUntil: "domcontentloaded" });
  const parentOrigin = new URL(page.url()).origin;
  const port = new URL(page.url()).port;
  const origins = slots.map(
    (slot) => `http://${slot}.onlyoffice.localhost:${port}`,
  );
  const responsePromises = origins.map((origin) =>
    page.waitForResponse(
      (response) => {
        const url = new URL(response.url());
        return url.origin === origin && url.pathname === "/subframe";
      },
      { timeout: 60_000 },
    ),
  );

  const readyMessages = await page.evaluate(
    async ({ origins, parentOrigin, slots }) =>
      await Promise.all(
        origins.map(
          (origin, index) =>
            new Promise<Record<string, unknown>>((resolve, reject) => {
              const instanceId = `real-slot-${slots[index]}`;
              const sessionToken = `real-session-${slots[index]}`;
              const frame = document.createElement("iframe");
              frame.name = instanceId;
              const url = new URL("/subframe", origin);
              url.searchParams.set("runtime", "compat");
              url.searchParams.set("instance", instanceId);
              url.searchParams.set("session", sessionToken);
              url.searchParams.set("parentOrigin", parentOrigin);
              const timer = window.setTimeout(
                () => reject(new Error(`${slots[index]} ready timed out`)),
                60_000,
              );
              const listener = (event: MessageEvent) => {
                const message = event.data as Record<string, unknown> | null;
                if (
                  event.source !== frame.contentWindow ||
                  event.origin !== origin ||
                  message?.source !== "onlyoffice-embed-sdk/compat-subframe/v1" ||
                  message.version !== 1 ||
                  message.type !== "ready" ||
                  message.instanceId !== instanceId ||
                  message.sessionToken !== sessionToken
                ) return;
                window.clearTimeout(timer);
                window.removeEventListener("message", listener);
                resolve(message);
              };
              window.addEventListener("message", listener);
              frame.src = url.href;
              document.body.appendChild(frame);
            }),
        ),
      ),
    { origins, parentOrigin, slots },
  );
  const responses = await Promise.all(responsePromises);

  expect(readyMessages).toHaveLength(12);
  for (let index = 0; index < slots.length; index += 1) {
    expect(readyMessages[index]?.instanceId).toBe(`real-slot-${slots[index]}`);
    expect(readyMessages[index]?.sessionToken).toBe(`real-session-${slots[index]}`);
    expect(responses[index]?.headers()["origin-agent-cluster"]).toBe("?1");
    const child = page.frames().find((frame) => frame.name() === `real-slot-${slots[index]}`);
    expect(child, `${slots[index]} child frame missing`).toBeDefined();
    await expect(
      child!.locator('[data-onlyoffice-subframe-runtime="compat"]'),
    ).toHaveAttribute("data-onlyoffice-subframe-protocol", "1");
  }

  await page.evaluate(() => {
    document.querySelectorAll('iframe[name^="real-slot-"]').forEach((frame) => frame.remove());
  });
});

test("real package facade activates and controls the hosted child", async ({ page }) => {
  test.setTimeout(90_000);
  const cdp = await page.context().newCDPSession(page);
  await cdp.send("Network.enable");
  await cdp.send("Network.clearBrowserCache");
  const docsApiOrigins = new Set<string>();
  const x2tOrigins = new Set<string>();
  const editorFrameOrigins = new Set<string>();
  const hostedAssetOrigins = new Set<string>();
  const zodiacAssetStatuses: number[] = [];
  let canonicalAssetResponses = 0;
  let docsApiRequests = 0;
  let preloadRequests = 0;
  page.on("request", (request) => {
    const url = new URL(request.url());
    if (url.pathname.endsWith("/web-apps/apps/api/documents/api.js")) {
      docsApiRequests += 1;
      docsApiOrigins.add(url.origin);
    }
    if (url.pathname.endsWith("/web-apps/apps/api/documents/preload.html")) {
      preloadRequests += 1;
    }
    if (url.pathname.endsWith("/x2t.js") || url.pathname.endsWith("/x2t.wasm")) {
      x2tOrigins.add(url.origin);
    }
    if (
      url.pathname.endsWith("/web-apps/apps/documenteditor/main/index.html") &&
      url.searchParams.get("frameEditorId") === "onlyoffice-compat-subframe-editor"
    ) {
      editorFrameOrigins.add(url.origin);
    }
    if (
      url.pathname.includes("/onlyoffice/runtime/") &&
      !url.pathname.endsWith("/web-apps/apps/api/documents/api.js") &&
      !url.pathname.endsWith("/web-apps/apps/documenteditor/main/index.html")
    ) {
      hostedAssetOrigins.add(url.origin);
    }
  });
  page.on("response", (response) => {
    const url = new URL(response.url());
    if (
      !url.pathname.includes("/onlyoffice/runtime/") ||
      url.pathname.endsWith("/web-apps/apps/api/documents/api.js") ||
      url.pathname.endsWith("/web-apps/apps/documenteditor/main/index.html")
    ) {
      return;
    }
    if (url.hostname === "onlyoffice.localhost") {
      canonicalAssetResponses += 1;
    } else if (url.hostname.endsWith(".onlyoffice.localhost")) {
      zodiacAssetStatuses.push(response.status());
    }
  });
  await page.goto(
    "/e2e/runtime-regressions?harness=real-compat-subframe&cacheProbe=1",
    { waitUntil: "domcontentloaded" },
  );
  const appPort = new URL(page.url()).port;
  const canonicalOrigin = `${new URL(page.url()).protocol}//onlyoffice.localhost:${appPort}`;
  const readCanonicalResourceTimings = async (slot: "rat" | "ox") => {
    await expect
      .poll(
        () =>
          page.evaluate(
            () => document.documentElement.dataset.onlyofficeCacheProbe ?? "",
          ),
        { timeout: 75_000 },
      )
      .toBe(slot);
    await expect
      .poll(
        () =>
          page.frames().filter((frame) => {
            const url = new URL(frame.url());
            return (
              url.hostname === `${slot}.onlyoffice.localhost` &&
              url.pathname.endsWith(
                "/web-apps/apps/documenteditor/main/index.html",
              )
            );
          }).length,
        { timeout: 10_000 },
      )
      .toBe(1);
    const editorFrame = page.frames().find((frame) => {
      const url = new URL(frame.url());
      return (
        url.hostname === `${slot}.onlyoffice.localhost` &&
        url.pathname.endsWith(
          "/web-apps/apps/documenteditor/main/index.html",
        )
      );
    });
    expect(editorFrame, `${slot} editor frame is missing`).toBeTruthy();
    const timings = await editorFrame!.evaluate((origin) =>
      performance
        .getEntriesByType("resource")
        .filter((entry) => entry.name.startsWith(origin))
        .map((entry) => {
          const resource = entry as PerformanceResourceTiming;
          return {
            name: resource.name,
            decodedBodySize: resource.decodedBodySize,
            transferSize: resource.transferSize,
          };
        }), canonicalOrigin);
    await page.evaluate((stage) => {
      document.documentElement.dataset.onlyofficeCacheProbeResume = stage;
    }, slot);
    return timings;
  };

  const ratTimings = await readCanonicalResourceTimings("rat");
  expect(
    ratTimings.length > 0,
    "rat editor did not expose canonical resource timing",
  ).toBeTruthy();
  const waitForThemeProbe = async (stage: "initial-dark" | "switched-light") => {
    await expect
      .poll(
        () =>
          page.evaluate(
            () => document.documentElement.dataset.onlyofficeThemeProbe ?? "",
          ),
        { timeout: 30_000 },
      )
      .toBe(stage);
    const editorFrame = page.frames().find((frame) => {
      const url = new URL(frame.url());
      return (
        url.hostname === "rat.onlyoffice.localhost" &&
        url.pathname.endsWith("/web-apps/apps/documenteditor/main/index.html")
      );
    });
    expect(editorFrame, `rat editor frame is missing during ${stage}`).toBeTruthy();
    return editorFrame!;
  };
  const darkFrame = await waitForThemeProbe("initial-dark");
  await expect
    .poll(() => darkFrame.locator("body").getAttribute("class"), {
      timeout: 10_000,
    })
    .toContain("theme-night");
  const initialThemeNavigation = await darkFrame.evaluate(() => ({
    timeOrigin: performance.timeOrigin,
    href: location.href,
  }));
  await page.evaluate(() => {
    document.documentElement.dataset.onlyofficeThemeProbeResume = "initial-dark";
  });
  const lightFrame = await waitForThemeProbe("switched-light");
  await expect
    .poll(() => lightFrame.locator("body").getAttribute("class"), {
      timeout: 10_000,
    })
    .toContain("theme-white");
  expect(await lightFrame.evaluate(() => performance.timeOrigin)).toBe(
    initialThemeNavigation.timeOrigin,
  );
  expect(lightFrame.url()).toBe(initialThemeNavigation.href);
  await page.evaluate(() => {
    document.documentElement.dataset.onlyofficeThemeProbeResume = "switched-light";
  });
  const oxTimings = await readCanonicalResourceTimings("ox");
  expect(
    oxTimings.length > 0,
    `ox editor did not expose canonical resource timing: ${JSON.stringify(oxTimings)}`,
  ).toBeTruthy();
  const assertSharedCacheHit = (suffix: string) => {
    const matches = oxTimings.filter(({ name, decodedBodySize }) => {
      const path = new URL(name).pathname;
      return path.endsWith(suffix) && decodedBodySize > 0;
    });
    expect(
      matches.some(({ transferSize }) => transferSize === 0),
      `${suffix} was not reused from the canonical browser cache: ${JSON.stringify(matches)}`,
    ).toBeTruthy();
  };
  assertSharedCacheHit("/app.js");
  assertSharedCacheHit("/fonts.wasm");
  await expect(page.getByTestId("real-compat-facade-status")).toHaveText(
    /passed|failed/,
    { timeout: 80_000 },
  );
  const status = await page.getByTestId("real-compat-facade-status").innerText();
  const steps = JSON.parse(
    await page.getByTestId("real-compat-facade-result").innerText(),
  ) as Array<{ name: string; status: string; detail?: string }>;
  expect(steps, JSON.stringify(steps, null, 2)).toEqual([
    { name: "real facade activation and identity", status: "passed" },
    { name: "real facade live interface theme", status: "passed" },
    { name: "real facade PDF print", status: "passed" },
    { name: "real facade readonly language destroy", status: "passed" },
    { name: "real facade shared canonical cache", status: "passed" },
    { name: "real facade legacy DOC activation", status: "passed" },
  ]);
  expect(status).toBe("passed");
  expect(Array.from(docsApiOrigins).sort()).toEqual([
    `http://ox.onlyoffice.localhost:${new URL(page.url()).port}`,
    `http://rat.onlyoffice.localhost:${new URL(page.url()).port}`,
    `http://tiger.onlyoffice.localhost:${new URL(page.url()).port}`,
  ]);
  expect(Array.from(editorFrameOrigins).sort()).toEqual([
    `http://ox.onlyoffice.localhost:${new URL(page.url()).port}`,
    `http://rat.onlyoffice.localhost:${new URL(page.url()).port}`,
    `http://tiger.onlyoffice.localhost:${new URL(page.url()).port}`,
  ]);
  expect(Array.from(hostedAssetOrigins).sort()).toEqual([
    canonicalOrigin,
    `http://ox.onlyoffice.localhost:${new URL(page.url()).port}`,
    `http://rat.onlyoffice.localhost:${new URL(page.url()).port}`,
    `http://tiger.onlyoffice.localhost:${new URL(page.url()).port}`,
  ]);
  expect(canonicalAssetResponses).toBeGreaterThan(0);
  expect(zodiacAssetStatuses.length).toBeGreaterThan(0);
  expect(new Set(zodiacAssetStatuses)).toEqual(new Set([308]));
  expect(Array.from(x2tOrigins)).toEqual([canonicalOrigin]);
  expect(docsApiRequests).toBeGreaterThan(0);
  expect(preloadRequests).toBe(0);
  expect(
    page.frames().filter((frame) =>
      frame.url().includes("/web-apps/apps/api/documents/preload"),
    ),
  ).toHaveLength(0);
});
