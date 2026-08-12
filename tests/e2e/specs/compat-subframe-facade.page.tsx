"use client";

import type {
  CreateOfficeEditorOptions,
  OfficeEditorInstance,
  OfficeEditorMount,
  OfficeSubframeSlot,
} from "@/components/onlyoffice-embed-sdk/compat/subframe";

type FacadeModule = typeof import(
  "@/components/onlyoffice-embed-sdk/compat/subframe"
);

let loadedFacade: FacadeModule | undefined;

async function loadFacade() {
  loadedFacade ??= await import(
    "@/components/onlyoffice-embed-sdk/compat/subframe"
  );
  return loadedFacade;
}

function facade() {
  if (!loadedFacade) throw new Error("compatibility facade was not loaded");
  return loadedFacade;
}

export type HarnessStep = {
  name: string;
  status: "passed" | "failed";
  detail?: string;
};

type InspectResult = {
  kind: "file" | "buffer" | "empty";
  fileName: string;
  sourceKind?: "buffer" | "url";
  emptyType?: string;
  bytes?: number[];
  constructorName?: string;
  resourceOrigin?: string;
};

type InternalMount = OfficeEditorMount & {
  request(action: string, payload?: unknown): Promise<unknown>;
};

function assert(condition: unknown, message: string): asserts condition {
  if (!condition) throw new Error(message);
}

function delay(timeout: number) {
  return new Promise<void>((resolve) => window.setTimeout(resolve, timeout));
}

async function waitFor(check: () => boolean, timeout = 1_000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeout) {
    if (check()) return;
    await delay(0);
  }
  throw new Error("Timed out waiting for compatibility facade state");
}

async function pauseForBrowserCacheProbe(stage: "rat" | "ox") {
  if (new URLSearchParams(window.location.search).get("cacheProbe") !== "1") {
    return;
  }
  document.documentElement.dataset.onlyofficeCacheProbe = stage;
  await waitFor(
    () => document.documentElement.dataset.onlyofficeCacheProbeResume === stage,
    15_000,
  );
  delete document.documentElement.dataset.onlyofficeCacheProbe;
}

async function pauseForThemeProbe(stage: "initial-dark" | "switched-light") {
  if (new URLSearchParams(window.location.search).get("cacheProbe") !== "1") {
    return;
  }
  document.documentElement.dataset.onlyofficeThemeProbe = stage;
  await waitFor(
    () => document.documentElement.dataset.onlyofficeThemeProbeResume === stage,
    30_000,
  );
  delete document.documentElement.dataset.onlyofficeThemeProbe;
}

function localBase() {
  return `${window.location.protocol}//onlyoffice.localhost:${window.location.port}`;
}

function hostUrl(slot: OfficeSubframeSlot = "rat") {
  return facade().getOfficeSubframeOrigin(slot, localBase());
}

function makeContainer() {
  const container = document.createElement("div");
  document.body.appendChild(container);
  return container;
}

function makeOptions(
  source: Pick<CreateOfficeEditorOptions, "file" | "buffer" | "url" | "emptyType" | "fileName">,
  overrides: Partial<CreateOfficeEditorOptions> = {},
): CreateOfficeEditorOptions {
  return {
    hostUrl: hostUrl(),
    requestTimeoutMs: 1_000,
    startupTimeoutMs: 2_000,
    destroyTimeoutMs: 500,
    callbackTimeoutMs: 100,
    ...source,
    ...overrides,
  } as CreateOfficeEditorOptions;
}

async function withEditor<T>(
  options: CreateOfficeEditorOptions,
  run: (instance: OfficeEditorInstance, mount: OfficeEditorMount) => Promise<T>,
) {
  const container = makeContainer();
  const mount = facade().mountOfficeEditor(container, options);
  try {
    const instance = await mount.activate();
    return await run(instance, mount);
  } finally {
    await mount.destroy();
    container.remove();
  }
}

async function inspectInput(options: CreateOfficeEditorOptions) {
  return withEditor(options, async (instance) =>
    instance.invokePlugin("__inspect_open__", null) as Promise<InspectResult>,
  );
}

async function testFixedSlotsAndIframeUrls() {
  const { OFFICE_SUBFRAME_SLOTS } = facade();
  assert(OFFICE_SUBFRAME_SLOTS.length === 12, "expected exactly 12 fixed slots");
  const expected = [
    "rat", "ox", "tiger", "rabbit", "dragon", "snake",
    "horse", "goat", "monkey", "rooster", "dog", "pig",
  ];
  assert(
    JSON.stringify(OFFICE_SUBFRAME_SLOTS) === JSON.stringify(expected),
    "fixed slot ordering changed",
  );

  for (const slot of OFFICE_SUBFRAME_SLOTS) {
    const container = makeContainer();
    const origin = hostUrl(slot);
    const mount = facade().mountOfficeEditor(
      container,
      makeOptions(
        { emptyType: "docx", fileName: `${slot}.docx` },
        { hostUrl: origin },
      ),
    );
    try {
      const frame = container.querySelector("iframe");
      assert(frame, `${slot} did not mount an iframe`);
      const url = new URL(frame.src);
      assert(url.origin === origin, `${slot} iframe used the wrong origin`);
      assert(url.searchParams.get("runtime") === "compat", "compat runtime query missing");
      assert(url.searchParams.get("instance") === mount.id, "instance query mismatch");
      assert(Boolean(url.searchParams.get("session")), "session query missing");
      assert(
        url.searchParams.get("parentOrigin") === window.location.origin,
        "parentOrigin query mismatch",
      );
      assert(frame.referrerPolicy === "strict-origin", "strict referrer policy missing");
    } finally {
      await mount.destroy();
      container.remove();
    }
  }

  assert(
    facade().getOfficeSubframeOrigin("dog") === "https://dog.onlyoffice.agent-bridges.com",
    "production slot origin has the wrong shape",
  );
  for (const invalidBase of [
    "http://onlyoffice.agent-bridges.com",
    "https://onlyoffice.agent-bridges.com:8443",
    "https://user:password@onlyoffice.agent-bridges.com",
  ]) {
    let rejected = false;
    try {
      facade().getOfficeSubframeOrigin("rat", invalidBase);
    } catch {
      rejected = true;
    }
    assert(rejected, `unsafe production base was accepted: ${invalidBase}`);
  }
  for (const invalidResourceOrigin of [
    window.location.origin,
    "https://rat.onlyoffice.agent-bridges.com",
    "https://user:password@onlyoffice.agent-bridges.com",
  ]) {
    let rejected = false;
    const container = makeContainer();
    let unsafeMount: OfficeEditorMount | undefined;
    try {
      unsafeMount = facade().mountOfficeEditor(
        container,
        makeOptions(
          { emptyType: "docx" },
          { resourceOrigin: invalidResourceOrigin },
        ),
      );
    } catch {
      rejected = true;
    } finally {
      await unsafeMount?.destroy();
      container.remove();
    }
    assert(
      rejected,
      `unsafe canonical resource origin was accepted: ${invalidResourceOrigin}`,
    );
  }
  {
    const container = makeContainer();
    let rejected = false;
    try {
      facade().mountOfficeEditor(
        container,
        makeOptions(
          { emptyType: "docx" },
          {
            hostUrl: "https://rat.onlyoffice.agent-bridges.com",
            resourceOrigin: "http://onlyoffice.localhost",
          },
        ),
      );
    } catch {
      rejected = true;
    } finally {
      container.remove();
    }
    assert(rejected, "production subframe accepted a local resource origin");
  }

  const firstContainer = makeContainer();
  const duplicateContainer = makeContainer();
  const firstMount = facade().mountOfficeEditor(
    firstContainer,
    makeOptions({ emptyType: "docx" }),
  );
  let duplicateError: unknown;
  try {
    facade().mountOfficeEditor(duplicateContainer, makeOptions({ emptyType: "docx" }));
  } catch (error) {
    duplicateError = error;
  }
  assert(
    duplicateError instanceof facade().OfficeHostIsolationError,
    "the same rat origin accepted a concurrent second mount",
  );
  await firstMount.destroy();
  const reusedMount = facade().mountOfficeEditor(
    duplicateContainer,
    makeOptions({ emptyType: "docx" }),
  );
  await reusedMount.destroy();
  firstContainer.remove();
  duplicateContainer.remove();
  assert(facade().getActiveOfficeEditorCount() === 0, "slot mounts leaked active instances");
}

async function testStrictEnvelopeChecks() {
  const container = makeContainer();
  const mount = facade().mountOfficeEditor(
    container,
    makeOptions(
      { emptyType: "docx" },
      { subframePath: "/subframe?readyDelay=250" },
    ),
  );
  const frame = container.querySelector("iframe");
  assert(frame?.contentWindow, "strict envelope test iframe is missing");
  const url = new URL(frame.src);
  const instanceId = url.searchParams.get("instance") || "";
  const sessionToken = url.searchParams.get("session") || "";
  const ready = {
    source: facade().COMPAT_SUBFRAME_PROTOCOL_SOURCE,
    version: facade().COMPAT_SUBFRAME_PROTOCOL_VERSION,
    type: "ready",
    instanceId,
    sessionToken,
  } as const;
  const activation = mount.activate();
  const dispatch = (origin: string, source: MessageEventSource | null, data: unknown) =>
    window.dispatchEvent(new MessageEvent("message", { origin, source, data }));

  dispatch(url.origin, window, ready);
  dispatch("http://ox.onlyoffice.localhost:1", frame.contentWindow, ready);
  dispatch(url.origin, frame.contentWindow, { ...ready, instanceId: "wrong-instance" });
  dispatch(url.origin, frame.contentWindow, { ...ready, sessionToken: "wrong-session" });

  const acceptedForgedReady = await Promise.race([
    activation.then(() => true),
    delay(100).then(() => false),
  ]);
  assert(!acceptedForgedReady, "host accepted a forged ready envelope");
  const instance = await activation;
  assert(instance.getState().status === "ready", "real ready envelope was not accepted");
  await mount.destroy();
  container.remove();
}

async function testStructuredCloneInputs() {
  const file = new File([Uint8Array.from([1, 2, 3])], "source.docx", {
    type: "application/test",
  });
  const fileResult = await inspectInput(makeOptions({ file }));
  assert(fileResult.kind === "file", "File input lost its wire kind");
  assert(fileResult.fileName === "source.docx", "File name was not preserved");
  assert(JSON.stringify(fileResult.bytes) === "[1,2,3]", "File bytes changed");
  assert(fileResult.constructorName === "File", "File did not remain a File");
  assert(
    fileResult.resourceOrigin === localBase(),
    "default resource origin did not use the shared local canonical host",
  );

  const blob = new Blob([Uint8Array.from([4, 5, 6])], { type: "application/test" });
  const blobResult = await inspectInput(
    makeOptions({ file: blob, fileName: "source.xlsx" }),
  );
  assert(blobResult.kind === "file", "Blob input lost its wire kind");
  assert(JSON.stringify(blobResult.bytes) === "[4,5,6]", "Blob bytes changed");
  assert(blobResult.constructorName === "Blob", "Blob did not remain a Blob");

  const arrayBuffer = Uint8Array.from([7, 8, 9]).buffer;
  const arrayResult = await inspectInput(
    makeOptions({ buffer: arrayBuffer, fileName: "source.pptx" }),
  );
  assert(JSON.stringify(arrayResult.bytes) === "[7,8,9]", "ArrayBuffer bytes changed");
  assert(arrayBuffer.byteLength === 3, "caller ArrayBuffer was detached");

  const uint8 = Uint8Array.from([10, 11, 12]);
  const uint8Result = await inspectInput(
    makeOptions({ buffer: uint8, fileName: "source.docx" }),
  );
  assert(JSON.stringify(uint8Result.bytes) === "[10,11,12]", "Uint8Array bytes changed");
  assert(uint8.byteLength === 3, "caller Uint8Array was detached");

  const urlResult = await inspectInput(
    makeOptions({ url: `${window.location.origin}/compat-fixture.docx` }),
  );
  assert(urlResult.kind === "buffer", "URL was not fetched in the parent");
  assert(urlResult.sourceKind === "url", "URL sourceKind was lost");
  assert(JSON.stringify(urlResult.bytes) === "[21,22,23]", "URL bytes changed");

  const emptyResult = await inspectInput(makeOptions({ emptyType: "csv" }));
  assert(emptyResult.kind === "empty", "empty input lost its wire kind");
  assert(emptyResult.emptyType === "csv", "empty document type changed");
}

async function testTypedResourceLoadingState() {
  const states: Array<ReturnType<OfficeEditorInstance["getLoadingState"]>> = [];
  await withEditor(
    makeOptions(
      { emptyType: "docx", fileName: "loading.docx" },
      {
        onLoadingChange: (state) => {
          states.push({ ...state });
        },
      },
    ),
    async (instance, mount) => {
      await waitFor(() => states.some((state) => state.resourceStatus === "downloaded"));
      assert(
        states.some(
          (state) =>
            state.phase === "static-resources" &&
            state.resourceStatus === "downloading" &&
            state.resourceDownload &&
            state.transferredBytes === 4096 &&
            state.resourceCount === 1,
        ),
        "verified resource download state was not delivered",
      );
      assert(
        JSON.stringify(instance.getLoadingState()) ===
          JSON.stringify(mount.getLoadingState()),
        "instance and mount loading snapshots diverged",
      );
      const downloadCount = states.filter((state) => state.resourceDownload).length;
      await instance.invokePlugin("__cache_hit__", null);
      await waitFor(() => instance.getLoadingState().resourceStatus === "cache-hit");
      assert(
        states.filter((state) => state.resourceDownload).length === downloadCount,
        "cache-hit state incorrectly reported a resource download",
      );
    },
  );
}

async function testFacadeActionsAndCallbacks() {
  const callbacks: string[] = [];
  const callbackBytes: number[][] = [];
  await withEditor(
    makeOptions(
      { emptyType: "docx", fileName: "actions.docx" },
      {
        mode: "preview",
        canReturnToPreview: true,
        onSave: async (file) => {
          callbacks.push("save");
          callbackBytes.push(Array.from(new Uint8Array(await file.arrayBuffer())));
          return true;
        },
        onSaveAs: async (file) => {
          callbacks.push("save-as");
          callbackBytes.push(Array.from(new Uint8Array(await file.arrayBuffer())));
          return true;
        },
        onDownload: async (file) => {
          callbacks.push("download");
          callbackBytes.push(Array.from(new Uint8Array(await file.arrayBuffer())));
        },
      },
    ),
    async (instance, mount) => {
      const identity = instance.getHostIdentity();
      assert(identity.packageVersion === "test", "activation identity was not cached");
      const remoteIdentity = await (mount as InternalMount).request("get-host-identity");
      assert(
        (remoteIdentity as { hostBuildId?: string }).hostBuildId === "fake-child",
        "get-host-identity RPC failed",
      );
      const remoteState = await (mount as InternalMount).request("get-state");
      assert((remoteState as { fileName?: string }).fileName === "actions.docx", "get-state RPC failed");

      instance.setReadonly(false);
      assert(instance.getState().mode === "edit", "preview did not enter edit mode");
      instance.setReadonly(true);
      assert(
        instance.getState().mode === "preview",
        "setReadonly did not synchronously return to preview mode",
      );
      instance.setInterfaceTheme("dark");
      await instance.setLanguage("zh-CN");
      await waitFor(() => instance.getState().readonly === true);

      await instance.invokePlugin("__mark_dirty__", null);
      await waitFor(() => instance.getState().dirty === true);
      const copied = await instance.exportCopy("pdf");
      assert(copied.name === "output.pdf", "exportCopy targetExt was not forwarded");
      assert(instance.getState().dirty, "exportCopy incorrectly cleared dirty state");

      const saved = await instance.save("pdf");
      const savedAs = await instance.saveAs("docx");
      const downloaded = await instance.download("xlsx");
      const originalOpen = window.open;
      let popupOpenCalls = 0;
      let printCalls = 0;
      let printedPdfNavigation = false;
      let navigatedPrintUrl = "";
      let printFrameCount = 0;
      let printFrame: HTMLIFrameElement | null = null;
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
            printFrame = frame;
            printFrameCount += 1;
            const markPrintableFrame = () => {
              navigatedPrintUrl = frame.src;
              printedPdfNavigation = frame.src.startsWith("blob:");
              const frameWindow = frame.contentWindow;
              if (!frameWindow) return;
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
      printObserver.observe(document.body, { childList: true });
      window.open = function () {
        popupOpenCalls += 1;
        return null;
      } as typeof window.open;
      let printed: File;
      try {
        const [primaryPrint, duplicatePrint] = await Promise.all([
          instance.print(),
          instance.print(),
        ]);
        printed = primaryPrint;
        assert(
          duplicatePrint.name === primaryPrint.name &&
            duplicatePrint.size === primaryPrint.size,
          "concurrent print did not share the active PDF export",
        );
      } finally {
        window.open = originalOpen;
        printObserver.disconnect();
      }
      assert(saved.name === "output.pdf", "save targetExt was not forwarded");
      assert(savedAs.name === "output.docx", "saveAs targetExt was not forwarded");
      assert(downloaded.name === "output.xlsx", "download targetExt was not forwarded");
      assert(
        printed.name === "output.pdf" &&
          printedPdfNavigation &&
          navigatedPrintUrl.startsWith("blob:") &&
          printFrameCount === 1 &&
          printCalls === 1 &&
          popupOpenCalls === 0 &&
          printFrame?.isConnected === false,
        "print PDF did not use exactly one direct PDF frame",
      );
      assert(await instance.confirmSaveToNewFormat({ dontshow: true }), "confirm RPC failed");
      assert(
        JSON.stringify(callbacks) === JSON.stringify(["save-as", "save", "save-as", "download"]),
        "output callbacks were not routed exactly once",
      );
      assert(
        callbackBytes.every((bytes) => JSON.stringify(bytes) === "[91,92,93]"),
        "callback file bytes changed across the boundary",
      );
    },
  );
}

async function testCallbackErrorAndTimeout() {
  await withEditor(
    makeOptions(
      { emptyType: "docx" },
      {
        callbackTimeoutMs: 30,
        onSave: () => {
          throw new Error("save callback rejected");
        },
        onSaveAs: () => new Promise(() => undefined),
      },
    ),
    async (instance) => {
      const saveError = await instance.save().then(
        () => new Error("save unexpectedly succeeded"),
        (error) => error instanceof Error ? error : new Error(String(error)),
      );
      assert(saveError.message.includes("save callback rejected"), "callback error was not returned");
      const timeoutError = await instance.saveAs().then(
        () => new Error("saveAs unexpectedly succeeded"),
        (error) => error instanceof Error ? error : new Error(String(error)),
      );
      assert(timeoutError.message.includes("save-as callback timed out"), "callback timeout was not returned");
    },
  );
}

async function testDestroyRejectsPending() {
  const container = makeContainer();
  const mount = facade().mountOfficeEditor(container, makeOptions({ emptyType: "docx" }));
  const instance = await mount.activate();
  const pending = instance.invokePlugin("__pending__", null).then(
    () => new Error("pending plugin unexpectedly succeeded"),
    (reason) => reason instanceof Error ? reason : new Error(String(reason)),
  );
  await delay(10);
  const destroying = instance.destroy();
  assert(
    instance.getState().destroyed && instance.getState().status === "destroyed",
    "destroy did not synchronously invalidate public state",
  );
  await destroying;
  const error = await pending;
  assert(error.name === "AbortError", "destroy did not reject pending RPC with AbortError");
  assert(instance.getState().destroyed, "destroyed state was not applied");
  assert(facade().getActiveOfficeEditorCount() === 0, "destroy leaked an active mount");
  container.remove();
}

const facadeTests = [
  ["fixed slots and iframe URLs", testFixedSlotsAndIframeUrls],
  ["strict source origin instance session", testStrictEnvelopeChecks],
  ["structured clone input wire", testStructuredCloneInputs],
  ["typed resource loading state", testTypedResourceLoadingState],
  ["facade actions and callbacks", testFacadeActionsAndCallbacks],
  ["callback error and timeout", testCallbackErrorAndTimeout],
  ["destroy rejects pending RPC", testDestroyRejectsPending],
] as const;

export async function runCompatSubframeFacadeTests(
  onChange: (steps: HarnessStep[]) => void,
) {
  await loadFacade();
  const steps: HarnessStep[] = [];
  for (const [name, run] of facadeTests) {
    try {
      await run();
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

export async function runRealCompatSubframeActivationTests(
  onChange: (steps: HarnessStep[]) => void,
) {
  const steps: HarnessStep[] = [];
  const run = async (name: string, test: () => Promise<void>) => {
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
  };

  let instance: OfficeEditorInstance | undefined;
  let container: HTMLElement | undefined;
  const ratLoadingStates: Array<ReturnType<OfficeEditorInstance["getLoadingState"]>> = [];
  const oxLoadingStates: Array<ReturnType<OfficeEditorInstance["getLoadingState"]>> = [];
  await run("real facade activation and identity", async () => {
    const facade = await loadFacade();
    container = makeContainer();
    instance = await facade.createOfficeEditor(container, {
      hostUrl: hostUrl("rat"),
      emptyType: "docx",
      fileName: "Real_Activation.docx",
      interfaceTheme: "dark",
      expectedHostIdentity: facade.HOSTED_COMPAT_SUBFRAME_IDENTITY,
      startupTimeoutMs: 75_000,
      requestTimeoutMs: 30_000,
      destroyTimeoutMs: 10_000,
      onLoadingChange: (state) => {
        ratLoadingStates.push({ ...state });
      },
    });
    assert(instance.getState().status === "ready", "real facade did not reach ready");
    assert(
      instance.getState().fileName === "Real_Activation.docx",
      "real facade changed the document identity",
    );
    assert(
      JSON.stringify(instance.getHostIdentity()) ===
        JSON.stringify(facade.HOSTED_COMPAT_SUBFRAME_IDENTITY),
      "real facade returned the wrong hosted identity",
    );
    await pauseForBrowserCacheProbe("rat");
    assert(
      ratLoadingStates.some(
        (state) => state.resourceStatus === "downloading" && state.resourceDownload,
      ),
      "cold hosted runtime did not report verified transferred resources",
    );
    await pauseForThemeProbe("initial-dark");
  });
  await run("real facade live interface theme", async () => {
    assert(instance, "real facade instance is missing");
    await instance.setInterfaceTheme("light");
    await pauseForThemeProbe("switched-light");
  });
  await run("real facade PDF print", async () => {
    assert(instance, "real facade instance is missing");
    const pdf = await instance.print();
    const header = new TextDecoder("ascii").decode(
      (await pdf.arrayBuffer()).slice(0, 5),
    );
    assert(
      pdf.name.toLowerCase().endsWith(".pdf") && header === "%PDF-",
      "real compat subframe did not return printable PDF bytes",
    );
  });
  await run("real facade readonly language destroy", async () => {
    assert(instance, "real facade instance is missing");
    instance.setReadonly(true);
    assert(instance.getState().readonly, "setReadonly did not update facade state");
    await instance.setLanguage("zh-CN");
    await instance.destroy();
    assert(instance.getState().destroyed, "real facade did not settle destroyed state");
    assert(facade().getActiveOfficeEditorCount() === 0, "real facade leaked its active mount");
    container?.remove();
  });
  await run("real facade shared canonical cache", async () => {
    const hotContainer = makeContainer();
    let hotInstance: OfficeEditorInstance | undefined;
    try {
      hotInstance = await facade().createOfficeEditor(hotContainer, {
        hostUrl: hostUrl("ox"),
        emptyType: "docx",
        fileName: "Real_Activation_Hot.docx",
        interfaceTheme: "light",
        lang: "en-US",
        expectedHostIdentity: facade().HOSTED_COMPAT_SUBFRAME_IDENTITY,
        startupTimeoutMs: 75_000,
        requestTimeoutMs: 30_000,
        destroyTimeoutMs: 10_000,
        onLoadingChange: (state) => {
          oxLoadingStates.push({ ...state });
        },
      });
      await pauseForBrowserCacheProbe("ox");
      assert(
        oxLoadingStates.some(
          (state) =>
            state.phase === "static-resources" &&
            state.resourceCount > 0 &&
            ["cache-hit", "downloading", "downloaded"].includes(
              state.resourceStatus,
            ),
        ),
        `hosted runtime did not report typed resource telemetry: ${JSON.stringify(oxLoadingStates)}`,
      );
    } finally {
      await hotInstance?.destroy();
      hotContainer.remove();
    }
  });
  await run("real facade legacy DOC activation", async () => {
    const facade = await loadFacade();
    const response = await fetch("/e2e/fixtures/example-title-ole.doc");
    assert(response.ok, "legacy DOC fixture request failed");
    const file = new File([await response.arrayBuffer()], "Example Title.doc", {
      type: "application/msword",
    });
    const legacyContainer = makeContainer();
    let legacyInstance: OfficeEditorInstance | undefined;
    try {
      legacyInstance = await facade.createOfficeEditor(legacyContainer, {
        hostUrl: hostUrl("tiger"),
        file,
        expectedHostIdentity: facade.HOSTED_COMPAT_SUBFRAME_IDENTITY,
        startupTimeoutMs: 75_000,
        requestTimeoutMs: 30_000,
        destroyTimeoutMs: 10_000,
      });
      assert(
        legacyInstance.getState().status === "ready" &&
          legacyInstance.getState().fileName === "Example Title.doc",
        "legacy DOC did not reach ready through the hosted facade",
      );
      const exported = await legacyInstance.saveAs("docx");
      assert(
        exported.name.toLowerCase().endsWith(".docx") && exported.size > 0,
        "legacy DOC did not export a non-empty DOCX through the hosted facade",
      );
    } finally {
      await legacyInstance?.destroy();
      legacyContainer.remove();
    }
    assert(
      facade.getActiveOfficeEditorCount() === 0,
      "legacy DOC hosted facade leaked its active mount",
    );
  });
  return steps;
}
