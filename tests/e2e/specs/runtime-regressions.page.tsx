"use client";

import { useEffect, useState } from "react";
import {
  callCrossOriginEditor,
  CROSS_ORIGIN_BRIDGE_MESSAGE,
  registerCrossOriginBridge,
  subscribeCrossOriginEditorEvent,
  unregisterCrossOriginBridge,
} from "@/components/onlyoffice-web-comp/internal/editor/runtime-bridge";
import io from "@/components/onlyoffice-web-comp/internal/editor/runtime-bridge";
import { EditorManager } from "@/components/onlyoffice-web-comp/core/editor-manager";
import {
  AvsFileType,
  type X2tConvertParams,
  type X2tConvertResult,
} from "@/components/onlyoffice-web-comp/internal/editor/types";
import {
  detectEditorBinFileType,
  getX2tExportFormats,
} from "@/components/onlyoffice-web-comp/internal/editor/utils";
import {
  X2tConversionError,
  X2tConverter,
} from "@/components/onlyoffice-web-comp/internal/editor/x2t";

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

const regressionTests = [
  ["x2t worker lifecycle", testX2tWorkerLifecycle],
  ["x2t queue and timeout", testX2tQueueAndTimeout],
  ["editor server latest wins", testEditorServerLatestWins],
  ["cross-origin bridge isolation", testBridgeIsolation],
  ["bridge lifecycle races", testBridgeLifecycleRaces],
  ["Editor.bin source detection", testEditorBinDetection],
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
  const [status, setStatus] = useState<"running" | "passed" | "failed">(
    "running",
  );
  const [steps, setSteps] = useState<RegressionStep[]>([]);

  useEffect(() => {
    let disposed = false;
    runRegressionTests((next) => {
      if (!disposed) setSteps(next);
    })
      .then(() => {
        if (!disposed) setStatus("passed");
      })
      .catch(() => {
        if (!disposed) setStatus("failed");
      });
    return () => {
      disposed = true;
    };
  }, []);

  return (
    <main className="p-4">
      <h1>Runtime regressions</h1>
      <p data-testid="regression-status">{status}</p>
      <pre data-testid="regression-result">
        {JSON.stringify(steps, null, 2)}
      </pre>
    </main>
  );
}
