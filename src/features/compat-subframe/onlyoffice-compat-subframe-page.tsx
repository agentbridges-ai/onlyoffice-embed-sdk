"use client";

import { useEffect, useState } from "react";
import {
  mountOfficeEditor,
  registerOnlyOfficeStaticResource,
  type CreateOfficeEditorOptions,
  type OfficeEditorInstance,
  type OfficeEditorMount,
} from "@/components/onlyoffice-embed-sdk/compat";
import {
  COMPAT_SUBFRAME_PROTOCOL_SOURCE,
  COMPAT_SUBFRAME_PROTOCOL_VERSION,
  COMPAT_SUBFRAME_RUNTIME,
  isCompatSubframeCallbackResponse,
  isCompatSubframeRequest,
  serializeCompatSubframeError,
  type CompatSubframeAction,
  type CompatSubframeCallbackName,
  type CompatSubframeCallbackRequest,
  type CompatSubframeCallbackResponse,
  type CompatSubframeChildMessage,
  type CompatSubframeDocument,
  type CompatSubframeEventName,
  type CompatSubframeOpenPayload,
  type CompatSubframeParentMessage,
  type CompatSubframeRequest,
} from "@/components/onlyoffice-embed-sdk/compat/subframe-protocol";
import {
  ONLYOFFICE_EMBED_HOST_ASSET_DIGEST,
  ONLYOFFICE_EMBED_HOST_BUILD_ID,
} from "@/components/onlyoffice-embed-sdk/compat/version";

const ONLYOFFICE_RESOURCE_VERSION = "9.4.0-develop";
const CALLBACK_TIMEOUT_MS = 75_000;
const REQUEST_TIMEOUT_MS = 75_000;
const OPEN_TIMEOUT_MS = 5 * 60_000;
const DESTROY_TIMEOUT_MS = 10_000;
const MAX_SEEN_REQUESTS = 2_048;
const EDITOR_CONTAINER_ID = "onlyoffice-compat-subframe-editor";

type RuntimeConfiguration = {
  instanceId: string;
  sessionToken: string;
  parentOrigin: string;
};

type PendingCallback = {
  callback: CompatSubframeCallbackName;
  resolve: (value: unknown) => void;
  reject: (error: Error) => void;
  timer: number;
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object";
}

function isNonEmptyString(value: unknown, maximum = 4_096): value is string {
  return typeof value === "string" && value.length > 0 && value.length <= maximum;
}

function parseExactHttpOrigin(value: string | null) {
  if (!value) return null;
  try {
    const parsed = new URL(value);
    if (
      !/^https?:$/.test(parsed.protocol) ||
      parsed.username ||
      parsed.password ||
      parsed.pathname !== "/" ||
      parsed.search ||
      parsed.hash ||
      parsed.origin !== value.replace(/\/$/, "")
    ) {
      return null;
    }
    return parsed.origin;
  } catch {
    return null;
  }
}

function readRuntimeConfiguration(): RuntimeConfiguration | null {
  const query = new URLSearchParams(window.location.search);
  if (query.get("runtime") !== COMPAT_SUBFRAME_RUNTIME) return null;
  const instanceId = query.get("instance");
  const sessionToken = query.get("session");
  const parentOrigin = parseExactHttpOrigin(query.get("parentOrigin"));
  if (
    !isNonEmptyString(instanceId, 256) ||
    !isNonEmptyString(sessionToken, 256) ||
    !parentOrigin
  ) {
    return null;
  }
  if (document.referrer) {
    try {
      if (new URL(document.referrer).origin !== parentOrigin) return null;
    } catch {
      return null;
    }
  }
  return { instanceId, sessionToken, parentOrigin };
}

function isIdentity(value: unknown) {
  if (!isRecord(value)) return false;
  return (
    isNonEmptyString(value.packageVersion, 128) &&
    isNonEmptyString(value.hostBuildId, 256) &&
    typeof value.assetManifestDigest === "string" &&
    /^[a-f0-9]{64}$/i.test(value.assetManifestDigest)
  );
}

function isDocument(value: unknown): value is CompatSubframeDocument {
  if (!isRecord(value) || !isNonEmptyString(value.kind, 32)) return false;
  if (value.kind === "empty") {
    return (
      ["docx", "xlsx", "pptx", "csv"].includes(String(value.emptyType)) &&
      isNonEmptyString(value.fileName, 1_024)
    );
  }
  if (value.kind === "file") {
    return (
      value.file instanceof Blob && isNonEmptyString(value.fileName, 1_024)
    );
  }
  if (value.kind === "buffer") {
    const buffer = value.buffer;
    return (
      (buffer instanceof Blob ||
        buffer instanceof ArrayBuffer ||
        ArrayBuffer.isView(buffer)) &&
      isNonEmptyString(value.fileName, 1_024) &&
      (value.sourceKind === undefined ||
        value.sourceKind === "buffer" ||
        value.sourceKind === "url")
    );
  }
  return false;
}

function isOpenPayload(value: unknown): value is CompatSubframeOpenPayload {
  if (!isRecord(value) || !isRecord(value.callbacks)) return false;
  let hostUrl: URL;
  let resourceOrigin: URL;
  try {
    hostUrl = new URL(String(value.hostUrl));
    resourceOrigin = new URL(String(value.resourceOrigin));
  } catch {
    return false;
  }
  const productionSubframe =
    window.location.hostname.endsWith(".onlyoffice.agent-bridges.com") &&
    window.location.protocol === "https:" &&
    window.location.port === "";
  const canonicalResourceOrigin = productionSubframe
    ? resourceOrigin.hostname === "onlyoffice.agent-bridges.com" &&
      resourceOrigin.protocol === "https:" &&
      resourceOrigin.port === ""
    : resourceOrigin.hostname === "onlyoffice.localhost" &&
      /^https?:$/.test(resourceOrigin.protocol) &&
      resourceOrigin.port === window.location.port;
  const plugins = value.plugins;
  const pluginsValid =
    plugins === undefined ||
    (isRecord(plugins) &&
      Array.isArray(plugins.configUrls) &&
      plugins.configUrls.every((url) => isNonEmptyString(url, 4_096)) &&
      (plugins.autostart === undefined ||
        (Array.isArray(plugins.autostart) &&
          plugins.autostart.every((guid) => isNonEmptyString(guid, 256)))));
  return (
    hostUrl.origin === window.location.origin &&
    !hostUrl.username &&
    !hostUrl.password &&
    canonicalResourceOrigin &&
    !resourceOrigin.username &&
    !resourceOrigin.password &&
    resourceOrigin.pathname === "/" &&
    !resourceOrigin.search &&
    !resourceOrigin.hash &&
    isDocument(value.document) &&
    ["edit", "readonly", "preview"].includes(String(value.mode)) &&
    typeof value.readonly === "boolean" &&
    (value.canReturnToPreview === undefined ||
      typeof value.canReturnToPreview === "boolean") &&
    (value.spellcheck === undefined || typeof value.spellcheck === "boolean") &&
    (value.interfaceTheme === undefined ||
      ["light", "dark"].includes(String(value.interfaceTheme))) &&
    (value.lang === undefined || isNonEmptyString(value.lang, 32)) &&
    (value.saveBehavior === undefined ||
      ["auto", "callback", "download"].includes(String(value.saveBehavior))) &&
    pluginsValid &&
    (value.expectedHostIdentity === undefined || isIdentity(value.expectedHostIdentity)) &&
    ["onSave", "onSaveAs", "onDownload"].every(
      (name) => typeof value.callbacks[name] === "boolean",
    )
  );
}

function isTargetExtension(value: unknown) {
  return value === undefined ||
    (typeof value === "string" && /^[a-z0-9]{1,16}$/i.test(value));
}

function validateActionPayload(request: CompatSubframeRequest) {
  const payload = request.payload;
  switch (request.action) {
    case "open":
      return isOpenPayload(payload);
    case "set-readonly":
      return isRecord(payload) && typeof payload.readonly === "boolean";
    case "set-theme":
      return isRecord(payload) && ["light", "dark"].includes(String(payload.theme));
    case "set-language":
      return isRecord(payload) && isNonEmptyString(payload.lang, 32);
    case "invoke-plugin":
      return isRecord(payload) && isNonEmptyString(payload.pluginGuid, 256);
    case "save":
    case "save-as":
    case "download":
    case "print":
      return payload === undefined ||
        (request.action !== "print" &&
          isRecord(payload) && isTargetExtension(payload.targetExt));
    case "confirm-save-to-new-format":
      return payload === undefined ||
        (isRecord(payload) &&
          (payload.title === undefined || isNonEmptyString(payload.title, 1_024)) &&
          (payload.message === undefined || isNonEmptyString(payload.message, 8_192)) &&
          (payload.dontshow === undefined || typeof payload.dontshow === "boolean"));
    case "get-state":
    case "get-host-identity":
    case "destroy":
      return payload === undefined;
  }
}

function toError(error: unknown) {
  if (error instanceof Error) return error;
  if (isRecord(error) && typeof error.message === "string") {
    const normalized = new Error(error.message);
    if (typeof error.name === "string") normalized.name = error.name;
    return normalized;
  }
  return new Error(String(error));
}

function withTimeout<T>(
  promise: Promise<T>,
  timeoutMs: number,
  message: string,
) {
  return new Promise<T>((resolve, reject) => {
    const timer = window.setTimeout(() => reject(new Error(message)), timeoutMs);
    void promise.then(
      (value) => {
        window.clearTimeout(timer);
        resolve(value);
      },
      (error) => {
        window.clearTimeout(timer);
        reject(error);
      },
    );
  });
}

export function OnlyOfficeCompatSubframePage() {
  const [configuration, setConfiguration] = useState<
    RuntimeConfiguration | null | undefined
  >(undefined);

  useEffect(() => {
    setConfiguration(readRuntimeConfiguration());
  }, []);

  useEffect(() => {
    if (!configuration || window.parent === window) return;

    const { instanceId, sessionToken, parentOrigin } = configuration;
    let generation = 0;
    let mount: OfficeEditorMount | null = null;
    let instance: OfficeEditorInstance | null = null;
    let openInFlight = false;
    let explicitOutput: CompatSubframeCallbackName | null = null;
    let callbackCapabilities = {
      onSave: false,
      onSaveAs: false,
      onDownload: false,
    };
    let disposed = false;
    const seenRequests = new Set<string>();
    const pendingCallbacks = new Map<string, PendingCallback>();

    const envelope = () => ({
      source: COMPAT_SUBFRAME_PROTOCOL_SOURCE,
      version: COMPAT_SUBFRAME_PROTOCOL_VERSION,
      instanceId,
      sessionToken,
    }) as const;

    const sendRaw = (message: CompatSubframeChildMessage) => {
      window.parent.postMessage(message, parentOrigin);
    };

    const send = (message: CompatSubframeChildMessage) => {
      try {
        sendRaw(message);
        return true;
      } catch (error) {
        console.error("[onlyoffice-embed-sdk] compat subframe message failed", error);
        return false;
      }
    };

    const postEvent = (event: CompatSubframeEventName, payload?: unknown) => {
      send({ ...envelope(), type: "event", event, payload });
    };

    const callbackRequest = (
      callback: CompatSubframeCallbackName,
      file: File,
    ) => {
      const callbackRequestId = `callback-${crypto.randomUUID()}`;
      return new Promise<unknown>((resolve, reject) => {
        const timer = window.setTimeout(() => {
          pendingCallbacks.delete(callbackRequestId);
          reject(new Error(`Parent ${callback} callback timed out`));
        }, CALLBACK_TIMEOUT_MS);
        pendingCallbacks.set(callbackRequestId, {
          callback,
          resolve,
          reject,
          timer,
        });
        const sent = send({
          ...envelope(),
          type: "callback-request",
          callbackRequestId,
          callback,
          payload: { file, fileName: file.name },
        } satisfies CompatSubframeCallbackRequest);
        if (!sent) {
          pendingCallbacks.delete(callbackRequestId);
          window.clearTimeout(timer);
          reject(new DOMException("Callback payload is not cloneable", "DataCloneError"));
        }
      });
    };

    const destroyCurrent = async () => {
      generation += 1;
      const current = mount;
      mount = null;
      instance = null;
      explicitOutput = null;
      if (current) {
        await withTimeout(
          current.destroy(),
          DESTROY_TIMEOUT_MS,
          "Compatibility editor destroy timed out",
        ).catch(() => undefined);
      }
      document.getElementById(EDITOR_CONTAINER_ID)?.replaceChildren();
    };

    const requireInstance = () => {
      if (!instance) throw new Error("Compatibility editor is not open");
      return instance;
    };

    const openEditor = async (payload: CompatSubframeOpenPayload) => {
      if (openInFlight || instance) {
        throw new Error("Compatibility editor is already open");
      }
      openInFlight = true;
      const ownGeneration = ++generation;
      callbackCapabilities = { ...payload.callbacks };
      postEvent("loading-change", {
        loading: true,
        phase: "runtime-loading",
        resourceStatus: "checking",
        resourceDownload: false,
        transferredBytes: 0,
        resourceCount: 0,
      });
      try {
        registerOnlyOfficeStaticResource({
          cdnOrigin: payload.resourceOrigin,
          frameOrigin: window.location.origin,
          onlyofficeVersion: ONLYOFFICE_RESOURCE_VERSION,
          onlyofficePath: `/onlyoffice/runtime/${ONLYOFFICE_EMBED_HOST_BUILD_ID}`,
          assetManifestDigest: ONLYOFFICE_EMBED_HOST_ASSET_DIGEST,
        });
        const container = document.getElementById(EDITOR_CONTAINER_ID);
        if (!(container instanceof HTMLElement)) {
          throw new Error("Compatibility subframe editor container is missing");
        }
        const input: Pick<CreateOfficeEditorOptions, "file" | "buffer" | "emptyType"> =
          payload.document.kind === "file"
            ? { file: payload.document.file }
            : payload.document.kind === "buffer"
              ? { buffer: payload.document.buffer }
              : { emptyType: payload.document.emptyType };
        const options: CreateOfficeEditorOptions = {
          // The direct facade keeps hostUrl only as informational compatibility
          // state and requires it to be independent from its owner window. The
          // outer parent origin satisfies that invariant; the actual hosted
          // runtime remains this already-validated zodiac subframe.
          hostUrl: parentOrigin,
          expectedHostIdentity: payload.expectedHostIdentity,
          ...input,
          fileName: payload.document.fileName,
          mode: payload.mode,
          readonly: payload.readonly,
          canReturnToPreview: payload.canReturnToPreview,
          spellcheck: payload.spellcheck,
          interfaceTheme: payload.interfaceTheme,
          lang: payload.lang,
          plugins: payload.plugins,
          saveBehavior: payload.saveBehavior,
          onReady: (readyInstance) => {
            if (disposed || ownGeneration !== generation) return;
            instance = readyInstance;
            postEvent("document-ready", { state: readyInstance.getState() });
          },
          onPluginReady: (pluginGuid, editorType) => {
            postEvent("plugin-ready", { pluginGuid, editorType });
          },
          onDirtyChange: (dirty, editor) => {
            postEvent("dirty-change", { dirty, state: editor.getState() });
          },
          onStateChange: (state) => postEvent("state-change", state),
          onLoadingChange: (state) => postEvent("loading-change", state),
          onError: (error) => postEvent("error", serializeCompatSubframeError(error)),
          onSave: async (file) => {
            if (explicitOutput) return true;
            if (payload.callbacks.onSave) {
                const result = await callbackRequest("save", file);
                return result === true ||
                  (isRecord(result) && result.handled === true);
            }
            if (
              payload.saveBehavior === "download" ||
              (!payload.saveBehavior && payload.document.kind === "empty") ||
              (payload.saveBehavior === "auto" && payload.document.kind === "empty")
            ) {
              return false;
            }
            throw new Error(
              'A save callback is required for this document source. Provide onSave or use saveBehavior: "download".',
            );
          },
          onSaveAs: payload.callbacks.onSaveAs
            ? async (file) => {
                const result = await callbackRequest("save-as", file);
                return result === true ||
                  (isRecord(result) && result.handled === true);
              }
            : undefined,
          onDownload: payload.callbacks.onDownload
            ? async (file) => {
                await callbackRequest("download", file);
              }
            : undefined,
        };
        const nextMount = mountOfficeEditor(container, options);
        mount = nextMount;
        const nextInstance = await nextMount.activate();
        if (disposed || ownGeneration !== generation) {
          await nextMount.destroy().catch(() => undefined);
          throw new DOMException("Editor activation was superseded", "AbortError");
        }
        instance = nextInstance;
        return {
          state: nextInstance.getState(),
          hostIdentity: nextInstance.getHostIdentity(),
        };
      } finally {
        openInFlight = false;
        if (!instance) {
          postEvent("loading-change", {
            loading: false,
            phase: "error",
            resourceStatus: "not-observed",
            resourceDownload: false,
            transferredBytes: 0,
            resourceCount: 0,
          });
        }
      }
    };

    const runExplicitOutput = async (
      callback: "save-as" | "download",
      targetExt?: string,
    ) => {
      if (explicitOutput) throw new Error("An explicit output request is already in progress");
      explicitOutput = callback;
      try {
        const file = await requireInstance().exportCopy(targetExt);
        const enabled = callback === "save-as"
          ? callbackCapabilities.onSaveAs
          : callbackCapabilities.onDownload;
        if (enabled) await callbackRequest(callback, file);
        return file;
      } finally {
        explicitOutput = null;
      }
    };

    const runAction = async (request: CompatSubframeRequest) => {
      const payload = request.payload as Record<string, unknown> | undefined;
      switch (request.action) {
        case "open":
          return openEditor(request.payload as CompatSubframeOpenPayload);
        case "set-readonly":
          requireInstance().setReadonly(payload?.readonly as boolean);
          return requireInstance().getState();
        case "set-theme":
          await requireInstance().setInterfaceTheme(
            payload?.theme as CreateOfficeEditorOptions["interfaceTheme"],
          );
          return requireInstance().getState();
        case "set-language":
          await requireInstance().setLanguage(payload?.lang as string);
          return requireInstance().getState();
        case "invoke-plugin":
          return requireInstance().invokePlugin(
            payload?.pluginGuid as string,
            payload?.payload,
          );
        case "save":
          return requireInstance().save(payload?.targetExt as string | undefined);
        case "save-as":
          return runExplicitOutput("save-as", payload?.targetExt as string | undefined);
        case "download":
          return runExplicitOutput("download", payload?.targetExt as string | undefined);
        case "print":
          return (
            requireInstance() as OfficeEditorInstance & {
              exportPrintPdfFile(): Promise<File>;
            }
          ).exportPrintPdfFile();
        case "confirm-save-to-new-format":
          return requireInstance().confirmSaveToNewFormat(
            request.payload as Parameters<OfficeEditorInstance["confirmSaveToNewFormat"]>[0],
          );
        case "get-state":
          return requireInstance().getState();
        case "get-host-identity":
          return requireInstance().getHostIdentity();
        case "destroy":
          await destroyCurrent();
          return null;
      }
    };

    const respond = (
      request: CompatSubframeRequest,
      ok: boolean,
      result?: unknown,
      error?: unknown,
    ) => {
      const response: CompatSubframeChildMessage = ok
        ? { ...envelope(), type: "response", requestId: request.requestId, ok: true, result }
        : {
            ...envelope(),
            type: "response",
            requestId: request.requestId,
            ok: false,
            error: serializeCompatSubframeError(error),
          };
      if (send(response) || !ok) return;
      send({
        ...envelope(),
        type: "response",
        requestId: request.requestId,
        ok: false,
        error: serializeCompatSubframeError(
          new DOMException("Response is not structured-cloneable", "DataCloneError"),
        ),
      });
    };

    const processRequest = (request: CompatSubframeRequest) => {
      if (seenRequests.has(request.requestId)) {
        respond(request, false, undefined, new Error("Duplicate requestId"));
        return;
      }
      seenRequests.add(request.requestId);
      if (seenRequests.size > MAX_SEEN_REQUESTS) {
        const oldest = seenRequests.values().next().value;
        if (typeof oldest === "string") seenRequests.delete(oldest);
      }
      if (!validateActionPayload(request)) {
        respond(request, false, undefined, new TypeError(`Invalid ${request.action} payload`));
        return;
      }
      if (request.action === "destroy") {
        void runAction(request).then(
          (result) => respond(request, true, result),
          (error) => respond(request, false, undefined, error),
        );
        return;
      }
      const actionPromise = Promise.resolve().then(() => runAction(request));
      const mayHaveExternalSideEffects =
        request.action === "save" ||
        request.action === "save-as" ||
        request.action === "download" ||
        request.action === "print" ||
        request.action === "invoke-plugin";
      const execution = mayHaveExternalSideEffects
        ? actionPromise
        : withTimeout(
            actionPromise,
            request.action === "open" ? OPEN_TIMEOUT_MS : REQUEST_TIMEOUT_MS,
            `Compatibility subframe ${request.action} timed out`,
          );
      void execution.then(
        (result) => respond(request, true, result),
        (error) => respond(request, false, undefined, error),
      );
    };

    const onMessage = (event: MessageEvent) => {
      if (disposed || event.source !== window.parent || event.origin !== parentOrigin) return;
      const message = event.data as unknown;
      if (
        (!isCompatSubframeRequest(message) &&
          !isCompatSubframeCallbackResponse(message)) ||
        message.instanceId !== instanceId ||
        message.sessionToken !== sessionToken
      ) {
        return;
      }
      if (isCompatSubframeCallbackResponse(message)) {
        const pending = pendingCallbacks.get(message.callbackRequestId);
        if (!pending) return;
        pendingCallbacks.delete(message.callbackRequestId);
        window.clearTimeout(pending.timer);
        if (message.ok) pending.resolve(message.result);
        else pending.reject(toError(message.error || "Parent callback failed"));
        return;
      }
      processRequest(message);
    };

    window.addEventListener("message", onMessage);
    send({ ...envelope(), type: "ready" });

    return () => {
      disposed = true;
      generation += 1;
      window.removeEventListener("message", onMessage);
      pendingCallbacks.forEach((pending) => {
        window.clearTimeout(pending.timer);
        pending.reject(new DOMException("Compatibility subframe disposed", "AbortError"));
      });
      pendingCallbacks.clear();
      void mount?.destroy();
      mount = null;
      instance = null;
    };
  }, [configuration]);

  if (configuration === undefined) {
    return (
      <main
        className="h-screen w-screen bg-white"
        data-onlyoffice-subframe-runtime="loading"
      />
    );
  }

  if (configuration === null) {
    return (
      <main
        className="flex h-screen w-screen items-center justify-center bg-white p-6 text-center text-sm text-red-700"
        data-onlyoffice-subframe-runtime="invalid"
      >
        Invalid compatibility subframe configuration.
      </main>
    );
  }

  return (
    <main
      className="h-screen w-screen overflow-hidden bg-white"
      data-onlyoffice-subframe-runtime={COMPAT_SUBFRAME_RUNTIME}
      data-onlyoffice-subframe-protocol={COMPAT_SUBFRAME_PROTOCOL_VERSION}
    >
      <div id={EDITOR_CONTAINER_ID} className="h-full w-full" />
    </main>
  );
}
