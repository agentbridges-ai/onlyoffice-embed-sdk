"use client";

import { useEffect, useState } from "react";
import {
  callCrossOriginEditor,
  CROSS_ORIGIN_BRIDGE_MESSAGE,
  guardPendingPrintFrameLoad,
  registerCrossOriginBridge,
  subscribeCrossOriginEditorEvent,
  unregisterCrossOriginBridge,
} from "@/components/onlyoffice-embed-sdk/internal/editor/runtime-bridge";
import io from "@/components/onlyoffice-embed-sdk/internal/editor/runtime-bridge";
import {
  EditorManager,
  editorManagerFactory,
} from "@/components/onlyoffice-embed-sdk/core/editor-manager";
import {
  AscSaveTypes,
  AvsFileType,
  type X2tConvertParams,
  type X2tConvertResult,
} from "@/components/onlyoffice-embed-sdk/internal/editor/types";
import {
  detectEditorBinFileType,
  getX2tExportFormats,
} from "@/components/onlyoffice-embed-sdk/internal/editor/utils";
import {
  converter,
  X2tConversionError,
  X2tConverter,
} from "@/components/onlyoffice-embed-sdk/internal/editor/x2t";
import { createEditorView } from "@/components/onlyoffice-embed-sdk/util/x2t";
import {
  OfficeHostIdentityMismatchError,
  OfficeHostIsolationError,
  ONLYOFFICE_EMBED_SDK_VERSION,
  createOfficeRuntimeResourceManager,
  mountOfficeEditor,
  registerOnlyOfficeStaticResource,
  resetOnlyOfficeStaticResource,
  resolveOfficeEmbedHostIdentity,
} from "@/components/onlyoffice-embed-sdk/compat";
import { OfficePluginBridge } from "@/components/onlyoffice-embed-sdk/compat/plugin-bridge";
import { OfficeRuntimeResourceCompatibilityError } from "@/components/onlyoffice-embed-sdk/compat/runtime-resources";
import {
  runCompatSubframeFacadeTests,
  runRealCompatSubframeActivationTests,
} from "./compat-subframe-facade.page";

type RegressionStep = {
  name: string;
  status: "passed" | "failed";
  detail?: string;
};

type WorkerRequest = {
  message: {
    id: number;
    type: string;
    payload: X2tConvertParams;
  };
  transfer: Transferable[];
};

class FakeWorker {
  onmessage: ((event: MessageEvent) => unknown) | null = null;
  onerror: ((event: ErrorEvent) => unknown) | null = null;
  onmessageerror: ((event: MessageEvent) => unknown) | null = null;
  readonly requests: WorkerRequest[] = [];
  terminated = false;

  postMessage(message: WorkerRequest["message"], transfer?: Transferable[]) {
    this.requests.push({ message, transfer: transfer ?? [] });
  }

  terminate() {
    this.terminated = true;
  }

  emit(data: Record<string, unknown>) {
    this.onmessage?.({ data } as MessageEvent);
  }
}

function assert(condition: unknown, message: string): asserts condition {
  if (!condition) throw new Error(message);
}

function delay(timeout: number) {
  return new Promise<void>((resolve) => window.setTimeout(resolve, timeout));
}

async function testNativePrintFrameLoadGate() {
  const frame = document.createElement("iframe");
  frame.id = "id-print-frame";
  let printHandlerCalls = 0;
  frame.onload = () => {
    printHandlerCalls += 1;
  };
  const allowPrintNavigation = guardPendingPrintFrameLoad(frame);

  frame.dispatchEvent(new Event("load"));
  assert(
    printHandlerCalls === 0,
    "the initial about:blank print frame load was not suppressed",
  );
  allowPrintNavigation();
  frame.dispatchEvent(new Event("load"));
  assert(
    Number(printHandlerCalls) === 1,
    "the final PDF print frame load did not run exactly once",
  );
}

async function waitFor<T>(read: () => T | undefined, timeout = 1_000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeout) {
    const value = read();
    if (value !== undefined) return value;
    await delay(0);
  }
  throw new Error("Timed out waiting for regression harness state");
}

function createConvertParams(seed = 1): X2tConvertParams {
  return {
    data: Uint8Array.from([seed, seed + 1]).buffer,
    fileFrom: "doc.docx",
    fileTo: "Editor.bin",
    media: { "media/image.png": Uint8Array.from([seed + 2]) },
    pdfBin: Uint8Array.from([seed + 3]),
    fonts: { "Test.ttf": Uint8Array.from([seed + 4]) },
    themes: { "themes/theme1.xml": Uint8Array.from([seed + 5]) },
  };
}

function createConvertResult(seed = 1): X2tConvertResult {
  return {
    output: Uint8Array.from([seed, seed + 1]),
    media: {},
    themes: {},
  };
}

async function testX2tWorkerLifecycle() {
  const workers: FakeWorker[] = [];
  const converter = new X2tConverter({
    workerFactory: () => {
      const worker = new FakeWorker();
      workers.push(worker);
      return worker as unknown as Worker;
    },
    isolateEachConversion: false,
    workerReadyTimeoutMs: 100,
    requestTimeoutMs: 100,
  });

  let ready = false;
  const init = converter.init().then(() => {
    ready = true;
  });
  await Promise.resolve();
  assert(!ready, "init resolved before the Worker ready handshake");
  workers[0]!.emit({ type: "ready" });
  await init;
  assert(converter.isInitialized, "converter did not enter ready state");

  const params = createConvertParams();
  const originalData = params.data;
  const conversion = converter.convert(params);
  const firstRequest = await waitFor(() => workers[0]?.requests[0]);
  assert(
    firstRequest.transfer.length === 5,
    `expected five input transferables, got ${firstRequest.transfer.length}`,
  );
  assert(
    firstRequest.message.payload.data !== originalData,
    "converter transferred the caller-owned ArrayBuffer",
  );
  assert(
    originalData.byteLength === 2,
    "caller-owned ArrayBuffer was detached",
  );
  workers[0]!.emit({
    id: firstRequest.message.id,
    type: "convert:done",
    payload: createConvertResult(),
  });
  assert((await conversion).output?.byteLength === 2, "conversion failed");

  const failedConversion = converter.convert(createConvertParams(10));
  const failedRequest = await waitFor(() => workers[0]?.requests[1]);
  workers[0]!.emit({
    id: failedRequest.message.id,
    type: "error",
    error: "conversion rejected",
    errorName: "X2tWorkerError",
    errorStack: "worker-stack",
    errorDetails: { stage: "execute", exitCode: 7 },
  });
  const failure = await failedConversion.catch((error) => error);
  assert(
    failure instanceof X2tConversionError &&
      failure.code === "worker-response" &&
      failure.details?.stage === "execute",
    "structured worker error details were lost",
  );
  assert(workers[0]!.terminated, "failed Worker was not retired");

  const recovered = converter.convert(createConvertParams(20));
  const recoveredWorker = await waitFor(() => workers[1]);
  recoveredWorker.emit({ type: "ready" });
  const recoveredRequest = await waitFor(() => recoveredWorker.requests[0]);
  recoveredWorker.emit({
    id: recoveredRequest.message.id,
    type: "convert:done",
    payload: createConvertResult(20),
  });
  assert((await recovered).output?.[0] === 20, "Worker recovery failed");
  converter.terminate();
}

async function testX2tQueueAndTimeout() {
  const workers: FakeWorker[] = [];
  const converter = new X2tConverter({
    workerFactory: () => {
      const worker = new FakeWorker();
      workers.push(worker);
      return worker as unknown as Worker;
    },
    isolateEachConversion: false,
    workerReadyTimeoutMs: 100,
    requestTimeoutMs: 100,
  });

  const first = converter.convert(createConvertParams(30));
  const worker = await waitFor(() => workers[0]);
  worker.emit({ type: "ready" });
  const firstRequest = await waitFor(() => worker.requests[0]);
  const second = converter.convert(createConvertParams(40));
  await delay(70);
  assert(
    worker.requests.length === 1,
    "second request entered the Worker before the first completed",
  );
  worker.emit({
    id: firstRequest.message.id,
    type: "convert:done",
    payload: createConvertResult(30),
  });
  await first;

  const secondRequest = await waitFor(() => worker.requests[1]);
  await delay(60);
  worker.emit({
    id: secondRequest.message.id,
    type: "convert:done",
    payload: createConvertResult(40),
  });
  assert((await second).output?.[0] === 40, "queued request timed out early");
  converter.terminate();

  const timeoutWorkers: FakeWorker[] = [];
  const timeoutConverter = new X2tConverter({
    workerFactory: () => {
      const next = new FakeWorker();
      timeoutWorkers.push(next);
      return next as unknown as Worker;
    },
    isolateEachConversion: false,
    workerReadyTimeoutMs: 100,
    requestTimeoutMs: 20,
  });
  const timedOut = timeoutConverter.convert(createConvertParams(50));
  const timeoutWorker = await waitFor(() => timeoutWorkers[0]);
  timeoutWorker.emit({ type: "ready" });
  await waitFor(() => timeoutWorker.requests[0]);
  const timeoutError = await timedOut.catch((error) => error);
  assert(
    timeoutError instanceof X2tConversionError &&
      timeoutError.code === "worker-request-timeout",
    "request timeout did not return a typed error",
  );
  assert(timeoutWorker.terminated, "timed-out Worker was not retired");

  const recovered = timeoutConverter.convert(createConvertParams(60));
  const recoveredWorker = await waitFor(() => timeoutWorkers[1]);
  recoveredWorker.emit({ type: "ready" });
  const recoveredRequest = await waitFor(() => recoveredWorker.requests[0]);
  recoveredWorker.emit({
    id: recoveredRequest.message.id,
    type: "convert:done",
    payload: createConvertResult(60),
  });
  assert(
    (await recovered).output?.[0] === 60,
    "converter did not recover after a timed-out Worker",
  );
  timeoutConverter.terminate();
}

async function testX2tIsolatedWorkers() {
  const workers: FakeWorker[] = [];
  const converter = new X2tConverter({
    workerFactory: () => {
      const worker = new FakeWorker();
      workers.push(worker);
      return worker as unknown as Worker;
    },
    workerReadyTimeoutMs: 100,
    requestTimeoutMs: 100,
  });

  const convert = async (seed: number, workerIndex: number) => {
    const conversion = converter.convert(createConvertParams(seed));
    const worker = await waitFor(() => workers[workerIndex]);
    worker.emit({ type: "ready" });
    const request = await waitFor(() => worker.requests[0]);
    worker.emit({
      id: request.message.id,
      type: "convert:done",
      payload: createConvertResult(seed),
    });
    assert(
      (await conversion).output?.[0] === seed,
      "isolated conversion returned the wrong output",
    );
    assert(worker.terminated, "completed conversion kept its Worker alive");
  };

  await convert(70, 0);
  assert(!converter.isInitialized, "completed conversion retained WASM state");
  await convert(80, 1);
  assert(workers.length === 2, "conversions reused a process-global WASM module");
}

async function testEditorServerLatestWins() {
  const manager = new EditorManager("server-generation-regression");
  const server = (manager as any).server;
  let resolveSlow!: (value: ArrayBuffer) => void;
  const slowBuffer = new Promise<ArrayBuffer>((resolve) => {
    resolveSlow = resolve;
  });

  await server.openUrl("https://files.example/slow.pdf", {
    fileType: "pdf",
    fileName: "Slow.pdf",
    loader: () => slowBuffer,
  });
  server.openNew("docx", "Latest.docx");
  resolveSlow(new TextEncoder().encode("%PDF-1.7\n%%EOF").buffer);
  await delay(0);
  await delay(0);

  const snapshot = server.getDocumentSnapshot();
  assert(snapshot.fileName === "Latest.docx", "stale title replaced latest");
  assert(
    detectEditorBinFileType(snapshot.binData!) === "docx",
    "stale load replaced the latest Editor.bin",
  );

  let resolveDestroyed!: (value: ArrayBuffer) => void;
  const destroyedBuffer = new Promise<ArrayBuffer>((resolve) => {
    resolveDestroyed = resolve;
  });
  await server.openUrl("https://files.example/destroyed.pdf", {
    fileType: "pdf",
    loader: () => destroyedBuffer,
  });
  server.reset();
  resolveDestroyed(new TextEncoder().encode("%PDF-1.7\n%%EOF").buffer);
  await delay(0);
  assert(
    server.getDocumentSnapshot().binData === undefined,
    "destroyed load repopulated Editor.bin",
  );
  manager.destroy();
}

async function testBridgeIsolation() {
  const frameEditorId = `bridge-regression-${Date.now()}`;
  const trustedOrigin = "https://trusted-office.example";
  const trusted = document.createElement("iframe");
  trusted.name = "frameEditor";
  trusted.hidden = true;
  trusted.src = `${trustedOrigin}/editor?frameEditorId=${frameEditorId}`;
  document.body.appendChild(trusted);

  const sibling = document.createElement("iframe");
  sibling.name = "frameEditor";
  sibling.hidden = true;
  sibling.src = `${trustedOrigin}/editor?frameEditorId=attacker`;
  document.body.appendChild(sibling);

  const manager = new EditorManager(frameEditorId);
  const server = (manager as any).server;
  const registered = registerCrossOriginBridge(
    frameEditorId,
    trusted,
    server,
    () => io(undefined, { deferConnect: true }),
  );
  assert(registered, "trusted bridge registration failed");

  let eventCount = 0;
  let bridgeInstanceId = "bridge-instance-one";
  const unsubscribe = subscribeCrossOriginEditorEvent(
    frameEditorId,
    "regression:event",
    () => {
      eventCount += 1;
    },
  );
  const eventData = {
    source: "onlyoffice-bridge",
    bridgeInstanceId,
    type: CROSS_ORIGIN_BRIDGE_MESSAGE.EDITOR_EVENT,
    frameEditorId,
    event: "regression:event",
    args: [],
  };

  window.dispatchEvent(
    new MessageEvent("message", {
      data: eventData,
      origin: trustedOrigin,
      source: sibling.contentWindow,
    }),
  );
  window.dispatchEvent(
    new MessageEvent("message", {
      data: eventData,
      origin: "https://evil.example",
      source: trusted.contentWindow,
    }),
  );
  assert(eventCount === 0, "spoofed bridge event reached subscribers");

  window.dispatchEvent(
    new MessageEvent("message", {
      data: {
        source: "onlyoffice-bridge",
        type: "hello",
        frameEditorId,
        bridgeInstanceId,
      },
      origin: trustedOrigin,
      source: trusted.contentWindow,
    }),
  );
  window.dispatchEvent(
    new MessageEvent("message", {
      data: eventData,
      origin: trustedOrigin,
      source: trusted.contentWindow,
    }),
  );
  assert(Number(eventCount) === 1, "trusted bridge event was rejected");

  window.dispatchEvent(
    new MessageEvent("message", {
      data: { ...eventData, bridgeInstanceId: "stale-bridge-instance" },
      origin: trustedOrigin,
      source: trusted.contentWindow,
    }),
  );
  window.dispatchEvent(
    new MessageEvent("message", {
      data: { ...eventData, bridgeInstanceId: undefined },
      origin: trustedOrigin,
      source: trusted.contentWindow,
    }),
  );
  assert(
    Number(eventCount) === 1,
    "bridge accepted a missing or stale instance id",
  );

  unregisterCrossOriginBridge(frameEditorId, {});
  window.dispatchEvent(
    new MessageEvent("message", {
      data: eventData,
      origin: trustedOrigin,
      source: trusted.contentWindow,
    }),
  );
  assert(
    Number(eventCount) === 2,
    "non-owner teardown removed the active bridge",
  );

  const nativeFetch = window.fetch;
  const nativeHandleRequest = server.handleRequest.bind(server);
  let nativeFetchCount = 0;
  let handledProxyPath = "";
  window.fetch = ((...args: Parameters<typeof fetch>) => {
    nativeFetchCount += 1;
    return nativeFetch(...args);
  }) as typeof fetch;
  server.handleRequest = async (request: Request) => {
    handledProxyPath = new URL(request.url).pathname;
    return new Response(Uint8Array.from([1, 2, 3]));
  };
  try {
    window.dispatchEvent(
      new MessageEvent("message", {
        data: {
          source: "onlyoffice-bridge",
          type: "http",
          frameEditorId,
          bridgeInstanceId,
          requestId: "prefixed-cache-fetch",
          method: "GET",
          url: `${trustedOrigin}/onlyoffice/9.4.0-develop/cache/files/data/key/Editor.bin`,
          responseType: "arraybuffer",
        },
        origin: trustedOrigin,
        source: trusted.contentWindow,
      }),
    );
    await waitFor(() => handledProxyPath || undefined);
    assert(
      handledProxyPath.endsWith("/cache/files/data/key/Editor.bin"),
      "bridge rejected a valid prefixed CDN cache path",
    );
    assert(
      nativeFetchCount === 0,
      "cache request escaped the in-memory server",
    );
    const handledValidProxyPath = handledProxyPath;

    window.dispatchEvent(
      new MessageEvent("message", {
        data: {
          source: "onlyoffice-bridge",
          type: "http",
          frameEditorId,
          bridgeInstanceId,
          requestId: "private-fetch",
          method: "GET",
          url: `${trustedOrigin}/private/account`,
        },
        origin: trustedOrigin,
        source: trusted.contentWindow,
      }),
    );
    await delay(0);
    assert(nativeFetchCount === 0, "bridge allowed an arbitrary native fetch");
    assert(
      handledProxyPath === handledValidProxyPath,
      "bridge sent an arbitrary path to the in-memory server",
    );

    const replacementPending = callCrossOriginEditor(
      frameEditorId,
      "regression:replacement",
      {},
      1_000,
    );
    await delay(0);
    trusted.src = `${trustedOrigin}/editor-v2?frameEditorId=${frameEditorId}`;
    assert(
      registerCrossOriginBridge(frameEditorId, trusted, server, () =>
        io(undefined, { deferConnect: true }),
      ),
      "same-iframe navigation was not registered",
    );
    const replacementResult = await Promise.allSettled([replacementPending]);
    assert(
      replacementResult[0]?.status === "rejected" &&
        replacementResult[0].reason instanceof DOMException &&
        replacementResult[0].reason.name === "AbortError",
      "iframe navigation did not abort the prior bridge request",
    );

    bridgeInstanceId = "bridge-instance-two";
    eventData.bridgeInstanceId = bridgeInstanceId;
    window.dispatchEvent(
      new MessageEvent("message", {
        data: {
          source: "onlyoffice-bridge",
          type: "hello",
          frameEditorId,
          bridgeInstanceId,
        },
        origin: trustedOrigin,
        source: trusted.contentWindow,
      }),
    );
    window.dispatchEvent(
      new MessageEvent("message", {
        data: eventData,
        origin: trustedOrigin,
        source: trusted.contentWindow,
      }),
    );
    assert(Number(eventCount) === 3, "navigated bridge did not re-handshake");

    const reloadPending = callCrossOriginEditor(
      frameEditorId,
      "regression:reload",
      {},
      1_000,
    );
    await delay(0);
    bridgeInstanceId = "bridge-instance-three";
    eventData.bridgeInstanceId = bridgeInstanceId;
    window.dispatchEvent(
      new MessageEvent("message", {
        data: {
          source: "onlyoffice-bridge",
          type: "hello",
          frameEditorId,
          bridgeInstanceId,
        },
        origin: trustedOrigin,
        source: trusted.contentWindow,
      }),
    );
    const reloadResult = await Promise.allSettled([reloadPending]);
    assert(
      reloadResult[0]?.status === "rejected" &&
        reloadResult[0].reason instanceof DOMException &&
        reloadResult[0].reason.name === "AbortError",
      "same-URL iframe reload did not abort the prior bridge request",
    );
    window.dispatchEvent(
      new MessageEvent("message", {
        data: eventData,
        origin: trustedOrigin,
        source: trusted.contentWindow,
      }),
    );
    assert(Number(eventCount) === 4, "reloaded bridge did not re-handshake");

    const unregisterStarted = performance.now();
    const unregisterPending = callCrossOriginEditor(
      frameEditorId,
      "regression:unregister",
      {},
      1_000,
    );
    unregisterCrossOriginBridge(frameEditorId);
    const unregisterResult = await Promise.allSettled([unregisterPending]);
    assert(
      unregisterResult[0]?.status === "rejected" &&
        unregisterResult[0].reason instanceof DOMException &&
        unregisterResult[0].reason.name === "AbortError" &&
        performance.now() - unregisterStarted < 500,
      "unregistered bridge request did not abort immediately",
    );
  } finally {
    window.fetch = nativeFetch;
    server.handleRequest = nativeHandleRequest;
    unsubscribe();
    unregisterCrossOriginBridge(frameEditorId);
    trusted.remove();
    sibling.remove();
    manager.destroy();
  }
}

async function testBridgeLifecycleRaces() {
  const frameEditorId = `bridge-lifecycle-${Date.now()}`;
  const staleManager = new EditorManager(frameEditorId);
  (
    staleManager as unknown as {
      syncCrossOriginReadOnly: (readOnly: boolean, retries: number) => boolean;
    }
  ).syncCrossOriginReadOnly(true, 2);
  staleManager.destroy();

  const iframe = document.createElement("iframe");
  iframe.name = "frameEditor";
  iframe.hidden = true;
  iframe.src = `${window.location.origin}/bridge-harness?frameEditorId=${frameEditorId}`;
  iframe.srcdoc = "<!doctype html><title>bridge lifecycle harness</title>";
  const loaded = new Promise<void>((resolve) => {
    iframe.addEventListener("load", () => resolve(), { once: true });
  });
  document.body.appendChild(iframe);
  await loaded;

  const childWindow = iframe.contentWindow;
  assert(childWindow, "bridge lifecycle iframe has no contentWindow");
  const postedMessages: Array<Record<string, unknown>> = [];
  const captureMessage = (event: MessageEvent) => {
    if (event.data && typeof event.data === "object") {
      postedMessages.push(event.data as Record<string, unknown>);
    }
  };
  childWindow.addEventListener("message", captureMessage);

  const replacementManager = new EditorManager(frameEditorId);
  const server = (replacementManager as any).server;
  const nativeHandleRequest = server.handleRequest.bind(server);
  let bridgeInstanceId = "lifecycle-instance-one";
  const dispatchHello = () => {
    window.dispatchEvent(
      new MessageEvent("message", {
        data: {
          source: "onlyoffice-bridge",
          type: "hello",
          frameEditorId,
          bridgeInstanceId,
        },
        origin: window.location.origin,
        source: childWindow,
      }),
    );
  };

  try {
    assert(
      registerCrossOriginBridge(frameEditorId, iframe, server, () =>
        io(undefined, { deferConnect: true }),
      ),
      "replacement bridge registration failed",
    );
    dispatchHello();
    await delay(75);
    assert(
      !postedMessages.some(
        (message) =>
          message.type === CROSS_ORIGIN_BRIDGE_MESSAGE.EDITOR_SET_READONLY,
      ),
      "stale read-only retry changed the replacement bridge",
    );

    let markRequestStarted!: () => void;
    let resolveResponse!: (response: Response) => void;
    const requestStarted = new Promise<void>((resolve) => {
      markRequestStarted = resolve;
    });
    const deferredResponse = new Promise<Response>((resolve) => {
      resolveResponse = resolve;
    });
    server.handleRequest = async () => {
      markRequestStarted();
      return deferredResponse;
    };

    const requestId = "stale-http-response";
    window.dispatchEvent(
      new MessageEvent("message", {
        data: {
          source: "onlyoffice-bridge",
          type: "http",
          frameEditorId,
          bridgeInstanceId,
          requestId,
          method: "GET",
          url: `${window.location.origin}/cache/files/data/key/Editor.bin`,
          responseType: "arraybuffer",
        },
        origin: window.location.origin,
        source: childWindow,
      }),
    );
    await requestStarted;
    postedMessages.length = 0;

    bridgeInstanceId = "lifecycle-instance-two";
    dispatchHello();
    resolveResponse(new Response(Uint8Array.from([1, 2, 3])));
    await delay(20);
    assert(
      !postedMessages.some(
        (message) =>
          message.type === "http:response" && message.requestId === requestId,
      ),
      "stale HTTP response was delivered to a reloaded bridge",
    );
  } finally {
    server.handleRequest = nativeHandleRequest;
    childWindow.removeEventListener("message", captureMessage);
    unregisterCrossOriginBridge(frameEditorId);
    replacementManager.destroy();
    iframe.remove();
  }
}

async function testPluginConfigProxyAllowlist() {
  const frameEditorId = `plugin-config-proxy-${Date.now()}`;
  const iframe = document.createElement("iframe");
  iframe.name = "frameEditor";
  iframe.hidden = true;
  iframe.src = `${window.location.origin}/plugin-proxy-harness?frameEditorId=${frameEditorId}`;
  iframe.srcdoc = "<!doctype html><title>plugin proxy harness</title>";
  const loaded = new Promise<void>((resolve) => {
    iframe.addEventListener("load", () => resolve(), { once: true });
  });
  document.body.appendChild(iframe);
  await loaded;

  const childWindow = iframe.contentWindow;
  assert(childWindow, "plugin proxy iframe has no contentWindow");
  const postedMessages: Array<Record<string, unknown>> = [];
  const captureMessage = (event: MessageEvent) => {
    if (event.data && typeof event.data === "object") {
      postedMessages.push(event.data as Record<string, unknown>);
    }
  };
  childWindow.addEventListener("message", captureMessage);

  const sourceUrl =
    "https://app.example.test/plugin-bundles/demo/config.json?build=1";
  const nativeFetch = window.fetch;
  const fetchCalls: Array<{ url: string; init?: RequestInit }> = [];
  window.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
    const url =
      typeof input === "string"
        ? input
        : input instanceof URL
          ? input.href
          : input.url;
    fetchCalls.push({ url, init });
    if (url !== sourceUrl) {
      throw new Error(`unexpected plugin config fetch: ${url}`);
    }
    return new Response(
      JSON.stringify({
        name: "Regression plugin",
        guid: "asc.{PLUGIN-PROXY-REGRESSION}",
        baseUrl: "./",
        icons: ["icons/plugin.svg"],
        variations: [
          {
            url: "ui/index.html",
            icons: [
              "icons/light.svg",
              { dark: "../shared/dark.svg" },
            ],
          },
        ],
      }),
      { headers: { "Content-Type": "application/json" } },
    );
  }) as typeof fetch;

  const manager = new EditorManager(frameEditorId);
  const server = (manager as any).server;
  const bridgeInstanceId = "plugin-proxy-instance";
  const dispatchBridgeMessage = (data: Record<string, unknown>) => {
    window.dispatchEvent(
      new MessageEvent("message", {
        data: {
          source: "onlyoffice-bridge",
          frameEditorId,
          bridgeInstanceId,
          ...data,
        },
        origin: window.location.origin,
        source: childWindow,
      }),
    );
  };

  try {
    assert(
      registerCrossOriginBridge(
        frameEditorId,
        iframe,
        server,
        () => io(undefined, { deferConnect: true }),
        iframe,
        [sourceUrl],
      ),
      "plugin config bridge registration failed",
    );
    dispatchBridgeMessage({ type: "hello" });
    await waitFor(() =>
      postedMessages.find((message) => message.type === "hello:ack"),
    );

    const allowedRequestId = "allowed-plugin-config";
    dispatchBridgeMessage({
      type: "http",
      requestId: allowedRequestId,
      method: "GET",
      url: sourceUrl,
      responseType: "text",
    });
    const allowedResponse = await waitFor(() =>
      postedMessages.find(
        (message) =>
          message.type === "http:response" &&
          message.requestId === allowedRequestId,
      ),
    );
    assert(allowedResponse.status === 200, "allowed plugin config was rejected");
    assert(fetchCalls.length === 1, "plugin config was not fetched exactly once");
    assert(fetchCalls[0]!.url === sourceUrl, "bridge fetched a different URL");
    assert(
      fetchCalls[0]!.init?.credentials === "same-origin" &&
        fetchCalls[0]!.init?.redirect === "error",
      "plugin config fetch did not use strict parent fetch options",
    );

    const config = JSON.parse(String(allowedResponse.responseBody)) as {
      baseUrl: string;
      icons: string[];
      variations: Array<{
        url: string;
        icons: Array<string | { dark: string }>;
      }>;
    };
    assert(
      config.baseUrl === "./" &&
        config.icons[0] === "icons/plugin.svg" &&
        config.variations[0]?.url === "ui/index.html" &&
        config.variations[0]?.icons[0] === "icons/light.svg" &&
        (config.variations[0]?.icons[1] as { dark: string }).dark ===
          "../shared/dark.svg",
      "bridge changed relative plugin config URLs",
    );

    const forgedUrl = new URL(sourceUrl);
    forgedUrl.searchParams.set("build", "forged");
    const rejectedRequestId = "forged-plugin-config";
    dispatchBridgeMessage({
      type: "http",
      requestId: rejectedRequestId,
      method: "GET",
      url: forgedUrl.href,
      responseType: "text",
    });
    const rejectedResponse = await waitFor(() =>
      postedMessages.find(
        (message) =>
          message.type === "http:response" &&
          message.requestId === rejectedRequestId,
      ),
    );
    assert(
      rejectedResponse.status === 403 && fetchCalls.length === 1,
      "forged plugin config URL escaped the exact allowlist",
    );
  } finally {
    window.fetch = nativeFetch;
    childWindow.removeEventListener("message", captureMessage);
    unregisterCrossOriginBridge(frameEditorId, iframe);
    manager.destroy();
    iframe.remove();
  }
}

async function testEditorBinDetection() {
  for (const [signature, fileType] of [
    ["DOCY;v5", "docx"],
    ["XLSY;v4", "xlsx"],
    ["PPTY;v3", "pptx"],
  ] as const) {
    assert(
      detectEditorBinFileType(new TextEncoder().encode(signature)) === fileType,
      `${signature.slice(0, 4)} detection failed`,
    );
    assert(
      detectEditorBinFileType(
        new TextEncoder().encode(btoa(signature + ";payload")),
      ) === fileType,
      `base64 ${signature.slice(0, 4)} detection failed`,
    );
  }
  const formats = getX2tExportFormats("docx", "xlsx");
  assert(
    formats.formatFrom === AvsFileType.AVS_FILE_CANVAS_SPREADSHEET,
    "cross-format export selected the target document as its source",
  );
}

async function testCompatibilityFacadeContracts() {
  const containerId = `compat-container-${Date.now()}`;
  const firstContainer = document.createElement("div");
  const duplicateContainer = document.createElement("div");
  firstContainer.id = containerId;
  duplicateContainer.id = containerId;
  document.body.append(firstContainer, duplicateContainer);

  const options = {
    hostUrl: "https://host.example/office-host.html",
    emptyType: "docx" as const,
  };
  const firstMount = mountOfficeEditor(firstContainer, options);
  let duplicateError: unknown;
  try {
    mountOfficeEditor(duplicateContainer, options);
  } catch (error) {
    duplicateError = error;
  }
  assert(
    duplicateError instanceof OfficeHostIsolationError,
    "compat facade allowed two active editors to share a container id",
  );
  await firstMount.destroy();

  const replacementMount = mountOfficeEditor(duplicateContainer, options);
  await replacementMount.destroy();
  firstContainer.remove();
  duplicateContainer.remove();

  const identityContainer = document.createElement("div");
  document.body.appendChild(identityContainer);
  let identityError: unknown;
  const identityMount = mountOfficeEditor(identityContainer, {
    ...options,
    expectedHostIdentity: {
      packageVersion: "0.0.0",
      hostBuildId: "wrong-host",
      assetManifestDigest: "wrong-assets",
    },
    onError() {
      throw new Error("consumer onError failed");
    },
  });
  try {
    await identityMount.activate();
  } catch (error) {
    identityError = error;
  }
  assert(
    identityError instanceof OfficeHostIdentityMismatchError,
    "compat facade replaced the startup error with a consumer callback error",
  );
  const postFailureMount = mountOfficeEditor(identityContainer, options);
  await postFailureMount.destroy();
  await identityMount.destroy();
  identityContainer.remove();

  const resources = await createOfficeRuntimeResourceManager({
    fetch: async () => new Response(new Uint8Array([1])),
  });
  assert(
    resources.getSnapshot().packageVersion === ONLYOFFICE_EMBED_SDK_VERSION,
    "resource facade package identity drifted from the SDK version",
  );
  let resourceError: unknown;
  try {
    resources.remainingBytes();
  } catch (error) {
    resourceError = error;
  }
  assert(
    resourceError instanceof OfficeRuntimeResourceCompatibilityError,
    "resource facade silently claimed an offline installation size",
  );

  const trustedManifestDigest = "ab".repeat(32);
  registerOnlyOfficeStaticResource({
    assetManifestDigest: trustedManifestDigest,
  });
  try {
    assert(
      (await resolveOfficeEmbedHostIdentity()).assetManifestDigest ===
        trustedManifestDigest,
      "compat identity ignored the deployment-provided manifest digest",
    );
  } finally {
    resetOnlyOfficeStaticResource();
  }

  const editorFrame = document.createElement("iframe");
  const siblingFrame = document.createElement("iframe");
  document.body.append(editorFrame, siblingFrame);
  const pluginFrame = editorFrame.contentDocument!.createElement("iframe");
  const siblingPluginFrame =
    siblingFrame.contentDocument!.createElement("iframe");
  editorFrame.contentDocument!.body.appendChild(pluginFrame);
  siblingFrame.contentDocument!.body.appendChild(siblingPluginFrame);
  const pluginWindow = pluginFrame.contentWindow!;
  const siblingPluginWindow = siblingPluginFrame.contentWindow!;
  const pluginGuid = "asc.{COMPAT-REGRESSION}";
  let readyCount = 0;
  const pluginBridge = new OfficePluginBridge(window, {
    // A configUrls-discovered plugin is trusted through its manifest origin;
    // it does not need to appear in plugins.autostart.
    pluginGuids: [],
    pluginOrigins: new Map([[pluginGuid, [window.location.origin]]]),
    isAllowedSource: (source) => source.parent === editorFrame.contentWindow,
    onReady: () => {
      readyCount += 1;
    },
  });
  const readyMessage = {
    protocol: "onlyoffice-browser-plugin/v1",
    type: "READY",
    pluginGuid,
    pluginInstanceId: "plugin-instance",
    editorType: "word",
  } as const;
  try {
    window.dispatchEvent(
      new MessageEvent("message", {
        data: readyMessage,
        origin: "https://forged-plugin.example",
        source: pluginWindow,
      }),
    );
    assert(readyCount === 0, "wrong-origin plugin claimed the bridge");

    window.dispatchEvent(
      new MessageEvent("message", {
        data: readyMessage,
        origin: window.location.origin,
        source: siblingPluginWindow,
      }),
    );
    assert(readyCount === 0, "same-origin sibling claimed the plugin bridge");

    window.dispatchEvent(
      new MessageEvent("message", {
        data: readyMessage,
        origin: window.location.origin,
        source: pluginWindow,
      }),
    );
    assert(Number(readyCount) === 1, "owned plugin READY was rejected");

    let invokeMessage: Record<string, unknown> | undefined;
    pluginWindow.addEventListener(
      "message",
      (event) => {
        invokeMessage = event.data as Record<string, unknown>;
      },
      { once: true },
    );
    const invocation = pluginBridge.invoke(pluginGuid, { operation: "ping" });
    let invocationSettled = false;
    void invocation.then(
      () => {
        invocationSettled = true;
      },
      () => {
        invocationSettled = true;
      },
    );
    await waitFor(() => invokeMessage);
    assert(
      invokeMessage?.type === "INVOKE" &&
        typeof invokeMessage.requestId === "string",
      "plugin bridge did not send a bound invocation",
    );
    window.dispatchEvent(
      new MessageEvent("message", {
        data: {
          protocol: "onlyoffice-browser-plugin/v1",
          type: "RESULT",
          pluginGuid,
          pluginInstanceId: "forged-plugin-instance",
          requestId: invokeMessage.requestId,
          ok: true,
          result: "forged",
        },
        origin: window.location.origin,
        source: pluginWindow,
      }),
    );
    window.dispatchEvent(
      new MessageEvent("message", {
        data: {
          protocol: "onlyoffice-browser-plugin/v1",
          type: "RESULT",
          pluginGuid,
          pluginInstanceId: "plugin-instance",
          requestId: invokeMessage.requestId,
          ok: true,
          result: "forged",
        },
        origin: "https://forged-plugin.example",
        source: pluginWindow,
      }),
    );
    await delay(0);
    assert(
      !invocationSettled,
      "wrong-origin or wrong-instance plugin completed an invocation",
    );
    window.dispatchEvent(
      new MessageEvent("message", {
        data: {
          protocol: "onlyoffice-browser-plugin/v1",
          type: "RESULT",
          pluginGuid,
          pluginInstanceId: "plugin-instance",
          requestId: invokeMessage.requestId,
          ok: true,
          result: "pong",
        },
        origin: window.location.origin,
        source: pluginWindow,
      }),
    );
    assert((await invocation) === "pong", "plugin result routing failed");
  } finally {
    pluginBridge.destroy();
    editorFrame.remove();
    siblingFrame.remove();
  }
}

async function testNativeEditorConfiguration() {
  const popup = window.open(
    "about:blank",
    `onlyoffice-native-config-${Date.now()}`,
    "popup,width=640,height=480",
  );
  assert(popup, "native-config popup Window was blocked");

  const container = popup.document.createElement("div");
  container.id = `native-config-host-${Date.now()}`;
  popup.document.body.appendChild(container);

  type CapturedConfig = {
    type?: string;
    document?: { permissions?: Record<string, boolean> };
    editorConfig?: {
      mode?: string;
      embedded?: { autostart?: string; toolbarDocked?: string };
      customization?: Record<string, unknown>;
    };
    events?: { onDocumentReady?: () => void };
  };
  const capturedConfigs: CapturedConfig[] = [];

  class FakeDocEditor {
    static version() {
      return "native-config-fake";
    }

    private readonly iframe: HTMLIFrameElement;

    constructor(id: string | undefined, config: CapturedConfig) {
      const host = popup.document.getElementById(id ?? "");
      if (!host) throw new Error("native-config DocsAPI received the wrong container");
      capturedConfigs.push(config);
      this.iframe = popup.document.createElement("iframe");
      this.iframe.name = "frameEditor";
      this.iframe.srcdoc = "<!doctype html><title>fake embedded viewer</title>";
      host.appendChild(this.iframe);
      popup.setTimeout(() => config.events?.onDocumentReady?.(), 0);
    }

    downloadAs() {}

    destroyEditor() {
      this.iframe.remove();
    }
  }

  popup.DocsAPI = {
    DocEditor: FakeDocEditor as unknown as DocEditorConstructor,
  };
  const mount = mountOfficeEditor(container, {
    hostUrl: "https://host.example/office-host.html",
    emptyType: "docx",
    mode: "preview",
    interfaceTheme: "dark",
    spellcheck: true,
  });

  try {
    const instance = await mount.activate();
    const config = capturedConfigs[0];
    assert(config, "native DocsAPI configuration was not captured");
    const permissions = config.document?.permissions;
    const customization = config.editorConfig?.customization;
    assert(
      config.type === "embedded" &&
        config.editorConfig?.mode === "view" &&
        config.editorConfig.embedded?.autostart === "document" &&
        config.editorConfig.embedded.toolbarDocked === "top",
      "preview did not use the upstream embedded/view shell",
    );
    assert(
      permissions?.edit === false &&
        permissions.download === false &&
        permissions.print === true,
      "preview permissions did not preserve native printing or fail closed",
    );
    assert(
      customization?.uiTheme === "theme-night" &&
        customization.spellcheck === true,
      "native theme or spellcheck customization was not forwarded",
    );
    assert(
      !Object.prototype.hasOwnProperty.call(customization, "logo"),
      "native ONLYOFFICE logo was replaced by a custom runtime URL",
    );

    instance.setReadonly(false);
    const editConfig = await waitFor(() => capturedConfigs[1]);
    assert(
      editConfig.type === "desktop" &&
        editConfig.editorConfig?.mode === "edit" &&
        editConfig.document?.permissions?.edit === true,
      "preview-to-edit did not remount the upstream desktop editor shell",
    );
    instance.setReadonly(true);
    const returnedPreviewConfig = await waitFor(() => capturedConfigs[2]);
    assert(
      returnedPreviewConfig.type === "embedded" &&
        returnedPreviewConfig.editorConfig?.mode === "view",
      "readonly did not return the compatibility instance to native preview",
    );
  } finally {
    await mount.destroy();
    popup.close();
  }
}

async function testDocsApiInitializationOverlapsSourceRead() {
  const popup = window.open(
    "about:blank",
    `onlyoffice-startup-overlap-${Date.now()}`,
    "popup,width=640,height=480",
  );
  assert(popup, "startup-overlap popup Window was blocked");

  const container = popup.document.createElement("div");
  container.id = `startup-overlap-host-${Date.now()}`;
  popup.document.body.appendChild(container);
  const manager = new EditorManager(container);
  const server = (manager as any).server as {
    open: (file: File) => Promise<unknown>;
  };

  let releaseSourceRead!: () => void;
  const sourceReadGate = new Promise<void>((resolve) => {
    releaseSourceRead = resolve;
  });
  let sourceReadStarted = false;
  let sourceReadFinished = false;
  const sourceFile = new File(["startup overlap"], "Overlap.docx");
  Object.defineProperty(sourceFile, "arrayBuffer", {
    configurable: true,
    value: async () => {
      sourceReadStarted = true;
      await sourceReadGate;
      sourceReadFinished = true;
      return new TextEncoder().encode("startup overlap").buffer;
    },
  });
  server.open = async (file) => {
    await file.arrayBuffer();
    return {};
  };

  type StartupConfig = { events?: { onDocumentReady?: () => void } };
  class FakeDocEditor {
    static version() {
      return "startup-overlap-fake";
    }

    private readonly iframe: HTMLIFrameElement;

    constructor(id: string | undefined, config: StartupConfig) {
      const host = popup.document.getElementById(id ?? "");
      if (!host) throw new Error("startup DocsAPI received the wrong container");
      this.iframe = popup.document.createElement("iframe");
      this.iframe.name = "frameEditor";
      this.iframe.srcdoc = "<!doctype html><title>startup overlap</title>";
      host.appendChild(this.iframe);
      popup.setTimeout(() => config.events?.onDocumentReady?.(), 0);
    }

    destroyEditor() {
      this.iframe.remove();
    }
  }

  const originalAppendChild = popup.document.head.appendChild.bind(
    popup.document.head,
  );
  let docsApiScript: HTMLScriptElement | undefined;
  popup.document.head.appendChild = ((node: Node) => {
    if (
      node.nodeName === "SCRIPT" &&
      (node as HTMLScriptElement).src.includes(
        "/web-apps/apps/api/documents/api.js",
      )
    ) {
      docsApiScript = node as HTMLScriptElement;
      return node;
    }
    return originalAppendChild(node);
  }) as typeof popup.document.head.appendChild;

  try {
    const activation = manager.create({
      container,
      file: sourceFile,
      fileName: sourceFile.name,
      fileType: "docx",
      isNew: false,
    });
    await waitFor(() => (sourceReadStarted && docsApiScript ? true : undefined));
    assert(
      !sourceReadFinished,
      "source bytes finished before the startup overlap could be observed",
    );
    assert(
      docsApiScript,
      "DocsAPI initialization did not start while File.arrayBuffer was pending",
    );

    popup.DocsAPI = {
      DocEditor: FakeDocEditor as unknown as DocEditorConstructor,
    };
    docsApiScript.dispatchEvent(new Event("load"));
    releaseSourceRead();
    await activation;

    const phases = manager
      .getLogger()
      .getEntries()
      .filter((entry) => entry.message === "startup-phase")
      .map((entry) => (entry.details[0] as { phase?: string }).phase);
    assert(
      phases.includes("document-source-ready") &&
        phases.includes("docs-api-ready") &&
        phases.includes("editor-mounted"),
      `startup diagnostics omitted a phase: ${JSON.stringify(phases)}`,
    );
  } finally {
    popup.document.head.appendChild = originalAppendChild;
    manager.destroy();
    popup.close();
  }
}

async function testCompatibilityNativeOutputCallbacks() {
  const popup = window.open(
    "about:blank",
    `onlyoffice-native-output-${Date.now()}`,
    "popup,width=640,height=480",
  );
  assert(popup, "native-output popup Window was blocked");

  const container = popup.document.createElement("div");
  container.id = `native-output-host-${Date.now()}`;
  popup.document.body.appendChild(container);

  type FakeEditorConfig = {
    events?: {
      onDocumentReady?: () => void;
      onDocumentStateChange?: (event: { data: boolean }) => void;
    };
  };
  let dispatchDocumentStateChange:
    | ((event: { data: boolean }) => void)
    | undefined;
  let handleDownloadAs: ((format: string) => void) | undefined;

  class FakeDocEditor {
    static version() {
      return "native-output-fake";
    }

    private readonly iframe: HTMLIFrameElement;

    constructor(id: string | undefined, config: FakeEditorConfig) {
      const host = popup.document.getElementById(id ?? "");
      if (!host) throw new Error("fake DocsAPI received the wrong container");
      this.iframe = popup.document.createElement("iframe");
      this.iframe.name = "frameEditor";
      this.iframe.srcdoc = "<!doctype html><title>fake editor</title>";
      host.appendChild(this.iframe);
      const editorWindow = this.iframe.contentWindow as Window & {
        Asc?: { editor?: { asc_nativeGetPDF?: () => Uint8Array } };
      };
      editorWindow.Asc = {
        editor: {
          asc_nativeGetPDF: () =>
            new TextEncoder().encode(
              "%PDF-1.7\n% compatibility print regression\n",
            ),
        },
      };
      dispatchDocumentStateChange = config.events?.onDocumentStateChange;
      popup.setTimeout(() => config.events?.onDocumentReady?.(), 0);
    }

    downloadAs(format: string) {
      handleDownloadAs?.(format);
    }

    destroyEditor() {
      this.iframe.remove();
    }
  }

  popup.DocsAPI = {
    DocEditor: FakeDocEditor as unknown as DocEditorConstructor,
  };

  const savedFiles: File[] = [];
  const saveAsFiles: File[] = [];
  const downloadedFiles: File[] = [];
  const callbackErrors: Error[] = [];
  const dirtyChanges: boolean[] = [];
  let callbackInstance: unknown;
  let rejectProgrammaticSave = false;
  let blockProgrammaticSave = false;
  let resolveBlockedSave: (() => void) | undefined;
  const mount = mountOfficeEditor(container, {
    hostUrl: "https://host.example/office-host.html",
    emptyType: "docx",
    saveBehavior: "callback",
    async onSave(file, instance) {
      savedFiles.push(file);
      callbackInstance = instance;
      if (rejectProgrammaticSave) {
        throw new Error("programmatic persistence failed");
      }
      if (blockProgrammaticSave) {
        await new Promise<void>((resolve) => {
          resolveBlockedSave = resolve;
        });
      }
      return true;
    },
    onSaveAs(file, instance) {
      saveAsFiles.push(file);
      callbackInstance = instance;
      throw new Error("save-as persistence failed");
    },
    onDownload(file, instance) {
      downloadedFiles.push(file);
      callbackInstance = instance;
    },
    onError(error) {
      callbackErrors.push(error);
    },
    onDirtyChange(dirty) {
      dirtyChanges.push(dirty);
    },
  });

  const originalConverterConvert = converter.convert;
  let converterPatched = false;
  let instance: Awaited<ReturnType<typeof mount.activate>> | undefined;
  try {
    instance = await mount.activate();
    const manager = editorManagerFactory.get(container);
    type CompatSnapshot = {
      fileName: string;
      fileType: string;
      binData?: Uint8Array;
      media: Record<string, Uint8Array>;
      themes: Record<string, Uint8Array>;
      capturedDirtyRevision?: number;
    };
    type CompatServer = {
      getDocument(): { key: string };
      handleRequest(request: Request): Promise<Response>;
      captureCurrentDocument(
        trigger: () => void,
        timeout?: number,
      ): Promise<CompatSnapshot>;
      getDocumentSnapshot(): CompatSnapshot;
      registerSocketTransport(socket: ReturnType<typeof io>): void;
    };
    const server = (
      manager as unknown as { server: CompatServer }
    ).server;
    const documentKey = server.getDocument().key;
    const socket = io(undefined, { deferConnect: true });
    const serverMessages: Array<Record<string, unknown>> = [];
    socket.on("message", (message: Record<string, unknown>) => {
      serverMessages.push(message);
    });
    server.registerSocketTransport(socket);
    socket.connect();

    const postDownloadAs = (
      cmd: Record<string, unknown>,
      bytes: Uint8Array,
    ) => {
      const url = new URL(
        `/downloadas/${documentKey}`,
        "https://runtime.test",
      );
      url.searchParams.set("cmd", JSON.stringify(cmd));
      return server.handleRequest(
        new Request(url, {
          method: "POST",
          body: bytes.slice().buffer,
        }),
      );
    };
    const waitForSaveMessage = (after: number) =>
      waitFor(() =>
        serverMessages.slice(after).find((message) => {
          const data = message.data as Record<string, unknown> | undefined;
          return message.type === "documentOpen" && data?.type === "save";
        }),
      );
    const assertAck = (message: Record<string, unknown>, label: string) => {
      const data = message.data as Record<string, unknown>;
      assert(
        data.status === "ok" && data.data === "onlyoffice://export/ack",
        `${label} did not suppress the native browser download`,
      );
    };

    const PopupFile = (popup as Window & { File: typeof File }).File;
    const programmaticFile = new PopupFile(
      [Uint8Array.from([1, 2, 3])],
      "programmatic.docx",
    );
    await (
      instance as unknown as {
        persistSavedFile(file: File): Promise<void>;
      }
    ).persistSavedFile(programmaticFile);
    assert(
      savedFiles.length === 1 &&
        saveAsFiles.length === 0 &&
        downloadedFiles.length === 0 &&
        callbackInstance === instance,
      "programmatic save persistence escaped the onSave channel",
    );

    converter.convert = async (params) => ({
      output: new Uint8Array(params.data.slice(0)),
      media: {},
    });
    converterPatched = true;

    dispatchDocumentStateChange({ data: true });
    await waitFor(() => instance?.getState().dirty === true);
    blockProgrammaticSave = true;
    const beforeNativeSave = serverMessages.length;
    const nativeSaveBytes = new TextEncoder().encode(
      "DOCY;v5;native-toolbar-save-latest-ZZSAVETEST",
    );
    let nativeCaptureRequest: Promise<Response> | undefined;
    handleDownloadAs = (format) => {
      assert(format === "bin", "native toolbar Save did not request Editor.bin");
      nativeCaptureRequest = postDownloadAs(
        {
          c: "save",
          savetype: AscSaveTypes.CompleteAll,
          outputformat: AvsFileType.AVS_FILE_DOCUMENT_DOCX,
          title: "New_Document.docx",
          nobase64: true,
          isSaveAs: true,
          saveAsPath: null,
        },
        nativeSaveBytes,
      );
    };
    socket.emit("message", {
      type: "saveChanges",
      changes: [JSON.stringify({ marker: "ZZSAVETEST" })],
      startSaveChanges: true,
      endSaveChanges: true,
    });
    await waitFor(() => resolveBlockedSave);
    assert(
      !serverMessages.slice(beforeNativeSave).some(
        (message) => message.type === "unSaveLock",
      ),
      "saveChanges was acknowledged before external persistence settled",
    );
    blockProgrammaticSave = false;
    resolveBlockedSave?.();
    resolveBlockedSave = undefined;
    await waitFor(() => nativeCaptureRequest);
    const nativeSaveResponse = await nativeCaptureRequest!;
    const nativeSaveResult = (await nativeSaveResponse.json()) as { status?: unknown };
    const nativeSaveAck = await waitFor(() =>
      serverMessages
        .slice(beforeNativeSave)
        .find((message) => message.type === "unSaveLock"),
    );
    await waitFor(() => instance?.getState().dirty === false);
    const persistedNativeBytes = new Uint8Array(
      await savedFiles[1]!.arrayBuffer(),
    );
    assert(
      nativeSaveResult.status === "ok" &&
        nativeSaveAck.type === "unSaveLock" &&
        Number(savedFiles.length) === 2 &&
        savedFiles[1]!.name.toLowerCase().endsWith(".docx") &&
        new TextDecoder().decode(persistedNativeBytes).includes("ZZSAVETEST") &&
        callbackInstance === instance,
      "saveChanges did not persist the latest Editor.bin through onSave before ACK",
    );
    handleDownloadAs = undefined;
    savedFiles.pop();

    const saveAsBytes = Uint8Array.from([11, 12, 13]);
    const beforeSaveAs = serverMessages.length;
    const exportSnapshotPromise = server.captureCurrentDocument(
      () => {},
      1_000,
    );
    const saveAsResponse = await postDownloadAs(
      {
        savetype: AscSaveTypes.CompleteAll,
        outputformat: AvsFileType.AVS_FILE_DOCUMENT_ODT,
        title: "copy.odt",
        isSaveAs: true,
      },
      saveAsBytes,
    );
    assert(saveAsResponse.ok, "Save Copy As request failed after callback error");
    assertAck(await waitForSaveMessage(beforeSaveAs), "Save Copy As");
    assert(
      Number(saveAsFiles.length) === 1 &&
        saveAsFiles[0]!.name === "copy.odt" &&
        saveAsFiles[0] instanceof PopupFile &&
        downloadedFiles.length === 0 &&
        savedFiles.length === 1,
      "Save Copy As did not route exclusively to onSaveAs",
    );
    assert(
      callbackErrors.length === 1 &&
        callbackErrors[0]!.message === "save-as persistence failed" &&
        instance.getState().status === "ready",
      "Save Copy As callback failure broke the editor lifecycle",
    );

    const downloadBytes = Uint8Array.from([21, 22, 23]);
    const beforeDownload = serverMessages.length;
    const downloadResponse = await postDownloadAs(
      {
        savetype: AscSaveTypes.CompleteAll,
        outputformat: AvsFileType.AVS_FILE_DOCUMENT_RTF,
        title: "download.rtf",
      },
      downloadBytes,
    );
    assert(downloadResponse.ok, "Download As request failed");
    assertAck(await waitForSaveMessage(beforeDownload), "Download As");
    assert(
      Number(downloadedFiles.length) === 1 &&
        downloadedFiles[0]!.name === "download.rtf" &&
        downloadedFiles[0] instanceof PopupFile &&
        Number(saveAsFiles.length) === 1 &&
        savedFiles.length === 1 &&
        callbackInstance === instance,
      "Download As did not route exclusively to onDownload",
    );
    assert(
      Array.from(new Uint8Array(await downloadedFiles[0]!.arrayBuffer())).join(
        ",",
      ) === Array.from(downloadBytes).join(","),
      "Download As callback received corrupted output bytes",
    );

    const sameFormatDownloadBytes = Uint8Array.from([31, 32, 33]);
    const beforeSameFormatDownload = serverMessages.length;
    const sameFormatDownloadResponse = await postDownloadAs(
      {
        savetype: AscSaveTypes.CompleteAll,
        outputformat: AvsFileType.AVS_FILE_DOCUMENT_DOCX,
        title: "download-copy.docx",
      },
      sameFormatDownloadBytes,
    );
    assert(sameFormatDownloadResponse.ok, "same-format Download As failed");
    assertAck(
      await waitForSaveMessage(beforeSameFormatDownload),
      "same-format Download As",
    );
    assert(
      Number(downloadedFiles.length) === 2 &&
        downloadedFiles[1]!.name === "download-copy.docx" &&
        Array.from(
          new Uint8Array(await downloadedFiles[1]!.arrayBuffer()),
        ).join(",") === Array.from(sameFormatDownloadBytes).join(","),
      "same-format non-Editor.bin output consumed the pending export",
    );

    const beforeExport = serverMessages.length;
    const exportRequest = postDownloadAs(
      {
        c: "save",
        savetype: AscSaveTypes.CompleteAll,
        // DocsAPI 9.4 sends the open document format here even when the
        // caller requested downloadAs("bin").
        outputformat: AvsFileType.AVS_FILE_DOCUMENT_DOCX,
        title: "New_Document.docx",
        nobase64: true,
        isSaveAs: true,
        saveAsPath: null,
      },
      new TextEncoder().encode("DOCY;v5;compat-export"),
    );
    const [exportSnapshot] = await Promise.all([
      exportSnapshotPromise,
      exportRequest,
    ]);
    assert(
      exportSnapshot.binData?.byteLength &&
        Number(saveAsFiles.length) === 1 &&
        Number(downloadedFiles.length) === 2,
      "native output consumed the pending programmatic export",
    );
    assertAck(
      await waitForSaveMessage(beforeExport),
      "real-shape programmatic export after native output",
    );

    converter.convert = async (params) => ({
      output: params.fileTo.toLowerCase().endsWith(".pdf")
        ? new TextEncoder().encode("%PDF-1.7\n% compatibility print regression\n")
        : new Uint8Array(params.data.slice(0)),
      media: {},
    });
    converterPatched = true;

    assert(
      dispatchDocumentStateChange,
      "fake editor did not expose document dirty events",
    );

    dispatchDocumentStateChange({ data: true });
    await waitFor(() => instance?.getState().dirty === true);

    let signalCopyCaptureStarted: (() => void) | undefined;
    const copyCaptureStarted = new Promise<void>((resolve) => {
      signalCopyCaptureStarted = resolve;
    });
    const copyDirtyChangeStart = dirtyChanges.length;
    handleDownloadAs = (format) => {
      assert(format === "bin", "exportCopy did not request an Editor.bin snapshot");
      signalCopyCaptureStarted?.();
    };
    const copyPromise = instance.exportCopy("odt");
    await copyCaptureStarted;
    const copyResponse = await postDownloadAs(
      {
        c: "save",
        savetype: AscSaveTypes.CompleteAll,
        outputformat: AvsFileType.AVS_FILE_DOCUMENT_DOCX,
        title: "New_Document.docx",
        nobase64: true,
        isSaveAs: true,
        saveAsPath: null,
      },
      new TextEncoder().encode("DOCY;v5;copy-export"),
    );
    assert(copyResponse.ok, "exportCopy snapshot request failed");
    const copiedFile = await copyPromise;
    handleDownloadAs = undefined;
    await delay(150);
    assert(
      copiedFile.name.endsWith(".odt") &&
        instance.getState().dirty &&
        !dirtyChanges.slice(copyDirtyChangeStart).includes(false),
      "exportCopy acknowledged persistence or cleared the dirty revision",
    );

    const originalPopupOpen = popup.open;
    let popupOpenCalls = 0;
    let printCalls = 0;
    let printedPdf = false;
    let printFrameCount = 0;
    const observedPrintFrames: HTMLIFrameElement[] = [];
    const printObserver = new MutationObserver((records) => {
      for (const record of records) {
        for (const node of record.addedNodes) {
          if (
            node.nodeType !== Node.ELEMENT_NODE ||
            (node as Element).getAttribute("data-onlyoffice-print-target") !==
              "pdf"
          ) {
            continue;
          }
          const frame = node as HTMLIFrameElement;
          observedPrintFrames.push(frame);
          printFrameCount += 1;
          const markPrintableFrame = () => {
            const frameWindow = frame.contentWindow;
            if (!frameWindow) return;
            printedPdf = frame.src.startsWith("blob:");
            Object.defineProperty(frameWindow, "print", {
              configurable: true,
              value: () => {
                printCalls += 1;
              },
            });
          };
          markPrintableFrame();
          frame.addEventListener("load", markPrintableFrame);
        }
      }
    });
    printObserver.observe(popup.document.body, { childList: true });
    popup.open = function () {
      popupOpenCalls += 1;
      return null;
    } as typeof popup.open;
    try {
      const [printedFile, duplicatePrintFile] = await Promise.all([
        instance.print(),
        instance.print(),
      ]);
      const printedHeader = new TextDecoder("ascii").decode(
        (await printedFile.arrayBuffer()).slice(0, 5),
      );
      assert(
        printedFile.name.endsWith(".pdf") &&
          duplicatePrintFile.name === printedFile.name &&
          duplicatePrintFile.size === printedFile.size &&
          printedHeader === "%PDF-" &&
          printedPdf &&
          printFrameCount === 1 &&
          printCalls === 1 &&
          popupOpenCalls === 0 &&
          observedPrintFrames.every((frame) => !frame.isConnected),
        "first print did not use exactly one direct PDF frame",
      );
      assert(
        instance.getState().dirty,
        "printing a copy incorrectly persisted the edited document",
      );

      printedPdf = false;
      const secondFile = await instance.print();
      assert(
        secondFile.name.endsWith(".pdf") &&
          printedPdf &&
          Number(printFrameCount) === 2 &&
          Number(printCalls) === 2 &&
          popupOpenCalls === 0 &&
          observedPrintFrames.every((frame) => !frame.isConnected),
        "second print did not create exactly one independent PDF frame",
      );
    } finally {
      printObserver.disconnect();
      popup.open = originalPopupOpen;
      for (const frame of observedPrintFrames) frame.remove();
      handleDownloadAs = undefined;
    }

    let signalCaptureStarted: (() => void) | undefined;
    const captureStarted = new Promise<void>((resolve) => {
      signalCaptureStarted = resolve;
    });
    const requestedDownloadFormats: string[] = [];
    handleDownloadAs = (format) => {
      requestedDownloadFormats.push(format);
      signalCaptureStarted?.();
    };
    const originalCaptureCurrentDocument =
      server.captureCurrentDocument.bind(server);
    let signalSnapshotAccepted: (() => void) | undefined;
    const snapshotAccepted = new Promise<void>((resolve) => {
      signalSnapshotAccepted = resolve;
    });
    let releaseSnapshotToManager: (() => void) | undefined;
    const snapshotReturnGate = new Promise<void>((resolve) => {
      releaseSnapshotToManager = resolve;
    });
    server.captureCurrentDocument = async (trigger, timeout) => {
      const snapshot = await originalCaptureCurrentDocument(trigger, timeout);
      signalSnapshotAccepted?.();
      await snapshotReturnGate;
      return snapshot;
    };
    const captureDirtyChangeStart = dirtyChanges.length;
    const captureWindowSave = instance.save("odt");
    await captureStarted;
    assert(
      requestedDownloadFormats.length === 1 &&
        requestedDownloadFormats[0] === "bin",
      "manager export did not request an Editor.bin capture",
    );

    const beforeCaptureResolution = serverMessages.length;
    const captureResolutionPromise = postDownloadAs(
      {
        c: "save",
        savetype: AscSaveTypes.CompleteAll,
        outputformat: AvsFileType.AVS_FILE_DOCUMENT_DOCX,
        title: "New_Document.docx",
        nobase64: true,
        isSaveAs: true,
        saveAsPath: null,
      },
      new TextEncoder().encode("DOCY;v5;capture-window-export"),
    );
    await snapshotAccepted;

    // The server has accepted the snapshot bytes and sampled their revision,
    // but manager.export has not received the snapshot yet. This edit belongs
    // to a newer revision and must survive both manager settlement and the
    // subsequent successful external onSave callback.
    dispatchDocumentStateChange({ data: true });
    await delay(150);
    assert(
      instance.getState().dirty &&
        !dirtyChanges.slice(captureDirtyChangeStart).includes(false),
      "capture-pending edit emitted a transient clean state before resolution",
    );

    releaseSnapshotToManager?.();
    const captureResolution = await captureResolutionPromise;
    assert(captureResolution.ok, "capture-window export request failed");
    assertAck(
      await waitForSaveMessage(beforeCaptureResolution),
      "capture-window programmatic export",
    );
    await captureWindowSave;
    handleDownloadAs = undefined;
    server.captureCurrentDocument = originalCaptureCurrentDocument;
    await delay(250);
    assert(
      Number(savedFiles.length) === 2 &&
        instance.getState().dirty &&
        !dirtyChanges.slice(captureDirtyChangeStart).includes(false),
      "a newer edit was cleared after capture and successful onSave",
    );

    (manager as unknown as { export: () => Promise<unknown> }).export =
      async () => {
        (manager as unknown as { dirty: boolean }).dirty = false;
        const snapshot = server.getDocumentSnapshot();
        if (!snapshot.binData) {
          throw new Error("compat regression snapshot is missing Editor.bin");
        }
        return {
          ...snapshot,
          binData: snapshot.binData,
          instanceId: manager.getInstanceId(),
        };
      };
    rejectProgrammaticSave = true;
    const failedSave = await instance.save().catch((error) => error);
    assert(
      failedSave instanceof Error &&
        failedSave.message === "programmatic persistence failed",
      "programmatic onSave rejection was not returned to the caller",
    );
    await delay(250);
    assert(
      instance.getState().dirty,
      "failed programmatic persistence lost its dirty latch after polling",
    );

    rejectProgrammaticSave = false;
    const originalState = instance.getState();
    blockProgrammaticSave = true;
    const dirtyChangeStart = dirtyChanges.length;
    const pendingSave = instance.save("odt");
    await waitFor(() => resolveBlockedSave);
    const duplicateSave = await instance.save().catch((error) => error);
    assert(
      duplicateSave instanceof Error &&
        duplicateSave.message ===
          "A save request is already in progress for this editor" &&
        Number(savedFiles.length) === 4,
      "concurrent save did not reject with the onlyoffice-browser contract",
    );

    dispatchDocumentStateChange?.({ data: true });
    await delay(150);
    blockProgrammaticSave = false;
    resolveBlockedSave?.();
    await pendingSave;
    resolveBlockedSave = undefined;
    await delay(250);
    const stateAfterTargetSave = instance.getState();
    assert(
      stateAfterTargetSave.fileName === originalState.fileName &&
        stateAfterTargetSave.fileType === originalState.fileType,
      "save(targetExt) mutated the open document identity",
    );
    assert(
      stateAfterTargetSave.dirty &&
        !dirtyChanges.slice(dirtyChangeStart).includes(false),
      "an edit made during onSave emitted a transient clean state",
    );

    await instance.save();
    await delay(250);
    assert(
      !instance.getState().dirty &&
        Number(savedFiles.length) === 5 &&
        Number(callbackErrors.length) === 2 &&
        callbackErrors[1]!.message === "programmatic persistence failed",
      "successful programmatic persistence did not release the dirty latch",
    );

    dispatchDocumentStateChange({ data: true });
    await waitFor(() => instance?.getState().dirty === true);
    rejectProgrammaticSave = true;
    const beforeRejectedNativeSave = serverMessages.length;
    const failedCaptureRequests: Promise<Response>[] = [];
    handleDownloadAs = (format) => {
      assert(format === "bin", "failed native Save did not request Editor.bin");
      failedCaptureRequests.push(
        postDownloadAs(
          {
            c: "save",
            savetype: AscSaveTypes.CompleteAll,
            outputformat: AvsFileType.AVS_FILE_DOCUMENT_DOCX,
            title: "New_Document.docx",
            nobase64: true,
            isSaveAs: true,
            saveAsPath: null,
          },
          new TextEncoder().encode("DOCY;v5;native-toolbar-save-rejected"),
        ),
      );
    };
    socket.emit("message", {
      type: "saveChanges",
      changes: [JSON.stringify({ marker: "rejected-save" })],
      startSaveChanges: true,
      endSaveChanges: true,
    });
    await waitFor(() =>
      callbackErrors.at(-1)?.message === "programmatic persistence failed"
        ? true
        : undefined,
    );
    await delay(150);
    assert(
      failedCaptureRequests.length === 1 &&
        instance.getState().dirty &&
        !serverMessages.slice(beforeRejectedNativeSave).some(
          (message) => message.type === "unSaveLock",
        ) &&
        callbackErrors.at(-1)?.message === "programmatic persistence failed",
      "failed saveChanges persistence was acknowledged or cleared dirty state",
    );

    rejectProgrammaticSave = false;
    const beforeNativeRetry = serverMessages.length;
    socket.emit("message", { type: "unSaveLock" });
    await waitFor(() => failedCaptureRequests.length === 2 ? true : undefined);
    await Promise.all(failedCaptureRequests);
    await waitFor(() =>
      serverMessages
        .slice(beforeNativeRetry)
        .find((message) => message.type === "unSaveLock"),
    );
    await waitFor(() => instance?.getState().dirty === false);
    handleDownloadAs = undefined;
  } finally {
    if (converterPatched) converter.convert = originalConverterConvert;
    await mount.destroy();
    popup.close();
  }
}

async function testCrossDocumentCompatMount() {
  // Use a same-origin child browsing context for the owner-Document contract.
  // A second script-opened popup can be left with a suspended event loop by
  // headless Chromium after the native-output regression closes its popup.
  // The iframe still provides an independent Window and Document without
  // relying on browser popup lifecycle policy.
  const documentHost = document.createElement("iframe");
  // `display:none` browsing contexts can have their event loop frozen on
  // shared headless Chromium runners. Keep the secondary Window rendered but
  // move its 1px surface off-screen so its DocsAPI ready timer remains active.
  Object.assign(documentHost.style, {
    border: "0",
    height: "1px",
    left: "-10000px",
    opacity: "0",
    pointerEvents: "none",
    position: "fixed",
    top: "0",
    width: "1px",
  });
  documentHost.src = "about:blank";
  document.body.appendChild(documentHost);
  const popup = documentHost.contentWindow;
  assert(popup, "same-origin child Window was not created");

  const containerId = `cross-document-host-${Date.now()}`;
  const rootContainer = document.createElement("div");
  rootContainer.id = containerId;
  document.body.appendChild(rootContainer);

  popup.document.title = "OnlyOffice cross-document host";
  const popupContainer = popup.document.createElement("div");
  popupContainer.id = containerId;
  popup.document.body.appendChild(popupContainer);

  type FakeEditorConfig = {
    editorConfig?: {
      plugins?: { pluginsData?: string[] };
    };
    events?: {
      onDocumentReady?: () => void;
    };
  };
  let popupPluginsData: string[] | undefined;

  class PopupDocEditor {
    static version() {
      return "cross-document-fake";
    }

    private readonly iframe: HTMLIFrameElement;

    constructor(id: string | undefined, config: FakeEditorConfig) {
      const host = popup.document.getElementById(id ?? "");
      if (!host) throw new Error("fake DocsAPI received the wrong Document");
      this.iframe = popup.document.createElement("iframe");
      this.iframe.name = "frameEditor";
      this.iframe.srcdoc = "<!doctype html><title>fake editor</title>";
      host.appendChild(this.iframe);
      popupPluginsData = config.editorConfig?.plugins?.pluginsData;
      popup.setTimeout(() => config.events?.onDocumentReady?.(), 0);
    }

    destroyEditor() {
      this.iframe.remove();
    }
  }

  popup.DocsAPI = {
    DocEditor: PopupDocEditor as unknown as DocEditorConstructor,
  };

  const pluginGuid = "asc.{E2E4D0B6-6F1E-4B80-9A4D-8F6B1C2D3E40}";
  const pluginConfigUrl =
    `${window.location.origin}/e2e/nexolyra-plugin/config.json`;
  const readyPlugins: Array<{ pluginGuid: string; editorType: string }> = [];
  const pluginCdnOrigin = "https://plugins-cdn.example.test";
  registerOnlyOfficeStaticResource({ cdnOrigin: pluginCdnOrigin });
  const options = {
    hostUrl: "https://host.example/office-host.html",
    emptyType: "xlsx" as const,
    plugins: {
      configUrls: [pluginConfigUrl],
    },
    onPluginReady(readyPluginGuid: string, editorType: string) {
      readyPlugins.push({ pluginGuid: readyPluginGuid, editorType });
    },
  };
  const rootMount = mountOfficeEditor(rootContainer, options);
  let popupMount: ReturnType<typeof mountOfficeEditor> | undefined;
  try {
    popupMount = mountOfficeEditor(popupContainer, options);
    const rootManager = editorManagerFactory.get(rootContainer);
    const popupManager = editorManagerFactory.get(popupContainer);
    assert(
      rootManager !== popupManager,
      "factory reused a manager across owner Documents",
    );

    const rootCreate = rootManager.create;
    const popupCreate = popupManager.create;
    try {
      // Keep this assertion focused on createEditorView's manager routing. The
      // real popup activation below continues to cover owner-Window DocsAPI,
      // DOM mounting and bridge registration.
      (rootManager as any).create = async () => rootManager;
      (popupManager as any).create = async () => popupManager;
      const rootViewManager = await createEditorView({
        container: rootContainer,
        containerId,
        fileName: "root.docx",
        isNew: true,
      });
      const popupViewManager = await createEditorView({
        container: popupContainer,
        containerId,
        fileName: "popup.docx",
        isNew: true,
      });
      assert(
        rootViewManager === rootManager && popupViewManager === popupManager,
        "createEditorView ignored the container owner Document",
      );
    } finally {
      (rootManager as any).create = rootCreate;
      (popupManager as any).create = popupCreate;
    }

    const instance = await popupMount.activate();
    assert(
      instance.getState().status === "ready",
      "popup editor did not reach ready",
    );
    const editorFrame = popupContainer.querySelector<HTMLIFrameElement>(
      'iframe[name="frameEditor"]',
    );
    assert(
      editorFrame,
      "DocsAPI mounted outside the popup container",
    );
    assert(
      !rootContainer.querySelector('iframe[name="frameEditor"]'),
      "popup DocsAPI leaked editor DOM into the root Document",
    );
    assert(
      Boolean(
        (
          popup as Window & {
            __ONLYOFFICE_SCOPED_IO__?: Record<string, unknown>;
          }
        ).__ONLYOFFICE_SCOPED_IO__?.[containerId],
      ),
      "runtime bridge registry was installed on the wrong Window",
    );
    const exposedPluginUrl = new URL(popupPluginsData?.[0] ?? "");
    assert(
      exposedPluginUrl.href === pluginConfigUrl &&
        exposedPluginUrl.origin !== pluginCdnOrigin &&
        exposedPluginUrl.pathname.endsWith("/config.json"),
      "EditorManager did not preserve the source config.json URL in CDN mode",
    );
    const configuredBridge = (
      instance as unknown as {
        pluginBridge: {
          configuredPluginGuids: Set<string>;
          configuredPluginOrigins: Map<string, Set<string>>;
        };
      }
    ).pluginBridge;
    assert(
      configuredBridge.configuredPluginGuids.has(pluginGuid) &&
        configuredBridge.configuredPluginOrigins
          .get(pluginGuid)
          ?.has(window.location.origin) &&
        readyPlugins.length === 0,
      "configUrls manifest GUID/origin was not added without autostart",
    );
  } finally {
    await popupMount?.destroy();
    await rootMount.destroy();
    editorManagerFactory.destroy(rootContainer);
    rootContainer.remove();
    documentHost.remove();
    resetOnlyOfficeStaticResource();
  }
}

const regressionTests = [
  ["x2t worker lifecycle", testX2tWorkerLifecycle],
  ["x2t queue and timeout", testX2tQueueAndTimeout],
  ["x2t isolated workers", testX2tIsolatedWorkers],
  ["editor server latest wins", testEditorServerLatestWins],
  ["cross-origin bridge isolation", testBridgeIsolation],
  ["bridge lifecycle races", testBridgeLifecycleRaces],
  ["plugin config proxy allowlist", testPluginConfigProxyAllowlist],
  ["Editor.bin source detection", testEditorBinDetection],
  ["compatibility facade contracts", testCompatibilityFacadeContracts],
  ["native print frame load gate", testNativePrintFrameLoadGate],
  ["native preview print logo configuration", testNativeEditorConfiguration],
  [
    "DocsAPI initialization overlaps source read",
    testDocsApiInitializationOverlapsSourceRead,
  ],
  [
    "compatibility native output callbacks",
    testCompatibilityNativeOutputCallbacks,
  ],
  ["cross-document compat mount", testCrossDocumentCompatMount],
] as const;

async function runRegressionTests(onChange: (steps: RegressionStep[]) => void) {
  const steps: RegressionStep[] = [];
  for (const [name, test] of regressionTests) {
    try {
      await test();
      steps.push({ name, status: "passed" });
    } catch (error) {
      steps.push({
        name,
        status: "failed",
        detail: error instanceof Error ? error.message : String(error),
      });
      onChange([...steps]);
      throw error;
    }
    onChange([...steps]);
  }
  return steps;
}

export function RuntimeRegressionsE2EPage() {
  const [suite, setSuite] = useState<"regressions" | "compat" | "real-compat">("regressions");
  const [result, setResult] = useState<{
    status: "running" | "passed" | "failed";
    steps: RegressionStep[];
  }>({ status: "running", steps: [] });

  useEffect(() => {
    let disposed = false;
    const harness = new URLSearchParams(window.location.search).get("harness");
    const nextSuite = harness === "compat-subframe"
      ? "compat"
      : harness === "real-compat-subframe"
        ? "real-compat"
        : "regressions";
    setSuite(nextSuite);
    const runner = nextSuite === "compat"
      ? runCompatSubframeFacadeTests
      : nextSuite === "real-compat"
        ? runRealCompatSubframeActivationTests
        : runRegressionTests;
    runner((next) => {
      if (!disposed) setResult({ status: "running", steps: next });
    })
      .then((finalSteps) => {
        if (!disposed) {
          setResult({ status: "passed", steps: finalSteps });
        }
      })
      .catch(() => {
        if (!disposed) {
          setResult((current) => ({ ...current, status: "failed" }));
        }
      });
    return () => {
      disposed = true;
    };
  }, []);

  return (
    <main className="p-4">
      <h1>{suite === "regressions" ? "Runtime regressions" : "Compatibility subframe facade"}</h1>
      <p data-testid={suite === "compat"
        ? "compat-facade-status"
        : suite === "real-compat"
          ? "real-compat-facade-status"
          : "regression-status"}>
        {result.status}
      </p>
      <pre data-testid={suite === "compat"
        ? "compat-facade-result"
        : suite === "real-compat"
          ? "real-compat-facade-result"
          : "regression-result"}>
        {JSON.stringify(result.steps, null, 2)}
      </pre>
    </main>
  );
}
