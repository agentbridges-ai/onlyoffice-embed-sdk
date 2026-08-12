import {
  ONLYOFFICE_EVENT_KEYS,
  getStaticResource,
  getOnlyOfficeStaticResourceManifestDigest,
  type OfficeThemeId,
} from "../const";
import {
  editorManagerFactory,
  type EditorDocumentExport,
  type EditorManager,
} from "../core/editor-manager";
import {
  onlyofficeEventbus,
  type LoadingChangeData,
  type OfficeLoadingPhase,
  type OfficeResourceLoadStatus,
} from "../core/eventbus";
import type {
  EditorDownloadOutput,
  OnlyOfficePluginOptions,
} from "../internal/editor/types";
import { getOnlyOfficeMimeType } from "../util/document-file";
import { initializeOnlyOffice } from "../util/initialize";
import { convertBinToDocument } from "../util/x2t";
import type { OnlyOfficeLang } from "../store/lang";
import { OfficePluginBridge } from "./plugin-bridge";
import { printOfficePdfFile } from "./print";
import {
  ONLYOFFICE_EMBED_HOST_BUILD_ID,
  ONLYOFFICE_EMBED_SDK_VERSION,
} from "./version";

const SUPPORTED_EMPTY_TYPES = ["docx", "xlsx", "pptx", "csv"] as const;
const STARTUP_TIMEOUT_MS = 5 * 60_000;
const DIRTY_POLL_INTERVAL_MS = 100;

type OfficeEmptyType = (typeof SUPPORTED_EMPTY_TYPES)[number];
type OfficeEditorStatus = "opening" | "ready" | "destroyed" | "error";

export type OfficeEditorMode = "edit" | "readonly" | "preview";
export type OfficeEditorInput = Blob | ArrayBuffer | Uint8Array;
export type OfficeEditorSourceKind =
  | "local-file"
  | "new-document"
  | "buffer"
  | "url";
export type OfficeSaveBehavior = "auto" | "callback" | "download";
export type OfficeInterfaceTheme = "light" | "dark";
export type OfficePluginOptions = OnlyOfficePluginOptions;
export type OfficeSaveCallbackResult = void | boolean;
export type OfficeSaveAsCallbackResult = void | boolean;
export type OfficeDownloadCallbackResult = void;

export type { OfficeLoadingPhase, OfficeResourceLoadStatus };

export type OfficeLoadingState = Pick<
  LoadingChangeData,
  | "loading"
  | "phase"
  | "resourceStatus"
  | "resourceDownload"
  | "transferredBytes"
  | "resourceCount"
>;

export type OfficeSaveToNewFormatConfirmationOptions = {
  title?: string;
  message?: string;
  dontshow?: boolean;
};

export interface OfficeHostIdentity {
  packageVersion: string;
  hostBuildId: string;
  assetManifestDigest: string;
}

export type OfficeHostUrlContext = {
  sessionId: string;
  fileName: string;
  fileType: string;
  mode: OfficeEditorMode;
};

export type OfficeHostUrlResolver =
  | string
  | ((context: OfficeHostUrlContext) => string | URL);

export interface CreateOfficeEditorOptions {
  /**
   * Kept for onlyoffice-browser source compatibility. The direct-embed
   * facade calls the resolver and reports its origin in OfficeEditorMount,
   * but it does not navigate an outer office-host.html iframe.
   */
  hostUrl: OfficeHostUrlResolver;
  /** Compared with resolveOfficeEmbedHostIdentity(), not hostUrl content. */
  expectedHostIdentity?: OfficeHostIdentity;
  file?: File | Blob;
  buffer?: OfficeEditorInput;
  url?: string;
  emptyType?: OfficeEmptyType;
  fileName?: string;
  mode?: OfficeEditorMode;
  readonly?: boolean;
  canReturnToPreview?: boolean;
  /** Initial native spellcheck policy. */
  spellcheck?: boolean;
  interfaceTheme?: OfficeInterfaceTheme;
  lang?: string;
  plugins?: OfficePluginOptions;
  fetchOptions?: RequestInit;
  /** Accepted for source compatibility; direct embed never reloads the page. */
  hardResetOnLastDestroy?: boolean;
  /** Accepted for source compatibility; direct teardown is synchronous. */
  destroyTimeoutMs?: number;
  onReady?: (instance: OfficeEditorInstance) => void;
  onPluginReady?: (
    pluginGuid: string,
    editorType: string,
    instance: OfficeEditorInstance,
  ) => void;
  saveBehavior?: OfficeSaveBehavior;
  onSave?: (
    file: File,
    instance: OfficeEditorInstance,
  ) => OfficeSaveCallbackResult | Promise<OfficeSaveCallbackResult>;
  /** Receives files produced by the native File → Save Copy As action. */
  onSaveAs?: (
    file: File,
    instance: OfficeEditorInstance,
  ) => OfficeSaveAsCallbackResult | Promise<OfficeSaveAsCallbackResult>;
  /** Receives files produced by the native File → Download As action. */
  onDownload?: (
    file: File,
    instance: OfficeEditorInstance,
  ) => OfficeDownloadCallbackResult | Promise<OfficeDownloadCallbackResult>;
  onDirtyChange?: (
    dirty: boolean,
    instance: OfficeEditorInstance,
  ) => void | Promise<void>;
  onStateChange?: (
    state: OfficeEditorState,
    instance: OfficeEditorInstance,
  ) => void | Promise<void>;
  /**
   * Reports host/runtime activity separately from verified static-resource
   * downloads. Show a resource download prompt only while
   * `state.resourceDownload` is true.
   */
  onLoadingChange?: (
    state: OfficeLoadingState,
    instance: OfficeEditorInstance,
  ) => void | Promise<void>;
  onError?: (error: Error, instance?: OfficeEditorInstance) => void;
}

export interface OfficeEditorState {
  id: string;
  fileName: string;
  fileType: string;
  mode: OfficeEditorMode;
  readonly: boolean;
  dirty: boolean;
  sourceKind: OfficeEditorSourceKind;
  status: OfficeEditorStatus;
  destroyed: boolean;
}

export interface OfficeEditorInstance {
  readonly id: string;
  invokePlugin(pluginGuid: string, payload: unknown): Promise<unknown>;
  save(targetExt?: string): Promise<File>;
  /** Export a copy without marking the current document as persisted. */
  exportCopy(targetExt?: string): Promise<File>;
  /** Export the current document as PDF, invoke the browser print flow, and return that PDF. */
  print(): Promise<File>;
  confirmSaveToNewFormat(
    options?: OfficeSaveToNewFormatConfirmationOptions,
  ): Promise<boolean>;
  /**
   * Applies the interface theme in place when the hosted runtime supports its
   * native theme controller. Existing callers may continue to ignore the
   * returned Promise.
   */
  setInterfaceTheme(theme: OfficeInterfaceTheme): void | Promise<void>;
  setReadonly(readonly: boolean): void;
  /** Changes the editor language while preserving the current document. */
  setLanguage(lang: string): Promise<void>;
  destroy(): Promise<void>;
  getState(): OfficeEditorState;
  getLoadingState(): OfficeLoadingState;
  getHostIdentity(): OfficeHostIdentity;
}

export type OfficeEditorMountPhase =
  | "host-loading"
  | "waiting-for-activation"
  | "runtime-loading"
  | "ready"
  | "error"
  | "destroyed";

export interface OfficeEditorMountState {
  id: string;
  origin: string;
  phase: OfficeEditorMountPhase;
  error?: Error;
}

export interface OfficeEditorMount {
  readonly id: string;
  activate(): Promise<OfficeEditorInstance>;
  destroy(): Promise<void>;
  getState(): OfficeEditorMountState;
  getLoadingState(): OfficeLoadingState;
}

export class OfficeHostIdentityMismatchError extends Error {
  readonly expected: OfficeHostIdentity;
  readonly actual: OfficeHostIdentity;

  constructor(expected: OfficeHostIdentity, actual: OfficeHostIdentity) {
    super(
      `Office host identity mismatch (expected ${expected.packageVersion}/${expected.hostBuildId}/${expected.assetManifestDigest}, received ${actual.packageVersion}/${actual.hostBuildId}/${actual.assetManifestDigest})`,
    );
    this.name = "OfficeHostIdentityMismatchError";
    this.expected = { ...expected };
    this.actual = { ...actual };
  }
}

export class OfficeHostIsolationError extends Error {
  readonly origin: string;
  readonly existingSessionId: string;
  readonly requestedSessionId: string;

  constructor(
    origin: string,
    existingSessionId: string,
    requestedSessionId: string,
  ) {
    super(
      `Office editor container is already owned by active session ${existingSessionId}`,
    );
    this.name = "OfficeHostIsolationError";
    this.origin = origin;
    this.existingSessionId = existingSessionId;
    this.requestedSessionId = requestedSessionId;
  }
}

export class OfficeEditorStartupError extends Error {
  readonly phase: string;
  readonly retryable: boolean;

  constructor(phase: string, message: string, retryable = true) {
    super(message);
    this.name = "OfficeEditorStartupError";
    this.phase = phase;
    this.retryable = retryable;
  }
}

/**
 * Signals a behavior that onlyoffice-browser implemented in its independent
 * outer host, but the direct-embed compatibility facade cannot reproduce.
 */
export class OfficeDirectEmbedCompatibilityError extends Error {
  readonly feature: string;

  constructor(feature: string, message: string) {
    super(message);
    this.name = "OfficeDirectEmbedCompatibilityError";
    this.feature = feature;
  }
}

type InitialDescriptor = {
  context: OfficeHostUrlContext;
  hostUrl: URL;
  state: OfficeEditorState;
};

type ReadyWaiter = {
  promise: Promise<void>;
  cancel: (error?: Error) => void;
};

let nextEditorId = 1;
const activeInstances = new Map<string, DirectEmbedOfficeEditor>();
const activeContainers = new WeakMap<HTMLElement, DirectEmbedOfficeEditor>();
const activeContainerIds = new WeakMap<
  Document,
  Map<string, DirectEmbedOfficeEditor>
>();

function getActiveContainerIds(ownerDocument: Document) {
  let ids = activeContainerIds.get(ownerDocument);
  if (!ids) {
    ids = new Map();
    activeContainerIds.set(ownerDocument, ids);
  }
  return ids;
}

function isHTMLElementContainer(value: unknown): value is HTMLElement {
  if (!value || typeof value !== "object") return false;
  const element = value as Element;
  const HTMLElementConstructor = element.ownerDocument?.defaultView?.HTMLElement;
  return (
    typeof HTMLElementConstructor === "function" &&
    value instanceof HTMLElementConstructor
  );
}

function makeSessionId() {
  const uuid = globalThis.crypto?.randomUUID?.();
  return uuid
    ? `office-editor-${uuid}`
    : `office-editor-${nextEditorId++}-${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

function getFileExtension(fileName: string) {
  return fileName.split(".").pop()?.toLowerCase() || "";
}

function normalizeExtension(value: string | undefined, fallback: string) {
  return (value || fallback).replace(/^\./, "").toLowerCase();
}

type PluginManifest = {
  guid?: unknown;
  baseUrl?: unknown;
  variations?: unknown;
};

function addPluginOrigin(
  originsByGuid: Map<string, Set<string>>,
  pluginGuid: string,
  value: string | URL,
) {
  try {
    const url = value instanceof URL ? value : new URL(value);
    if (url.protocol !== "http:" && url.protocol !== "https:") return;
    let origins = originsByGuid.get(pluginGuid);
    if (!origins) {
      origins = new Set();
      originsByGuid.set(pluginGuid, origins);
    }
    origins.add(url.origin);
  } catch {
    // Invalid manifest URLs are never added to the postMessage allowlist.
  }
}

async function resolveConfiguredPluginOrigins(
  ownerWindow: Window,
  plugins: OfficePluginOptions | undefined,
) {
  const originsByGuid = new Map<string, Set<string>>();
  await Promise.all(
    (plugins?.configUrls ?? []).map(async (configUrl) => {
      let sourceUrl: URL;
      try {
        sourceUrl = new URL(configUrl, ownerWindow.location.href);
      } catch {
        return;
      }
      if (
        (sourceUrl.protocol !== "http:" && sourceUrl.protocol !== "https:") ||
        sourceUrl.username ||
        sourceUrl.password
      ) {
        return;
      }
      sourceUrl.hash = "";

      let manifest: PluginManifest;
      try {
        const response = await ownerWindow.fetch(sourceUrl.href, {
          method: "GET",
          credentials: "same-origin",
          redirect: "error",
          headers: { Accept: "application/json" },
        });
        if (!response.ok) return;
        manifest = (await response.json()) as PluginManifest;
      } catch {
        return;
      }

      const pluginGuid =
        typeof manifest.guid === "string" ? manifest.guid.trim() : "";
      if (!pluginGuid) return;

      addPluginOrigin(originsByGuid, pluginGuid, sourceUrl);
      let pluginBaseUrl = new URL("./", sourceUrl);
      if (typeof manifest.baseUrl === "string" && manifest.baseUrl.trim()) {
        try {
          pluginBaseUrl = new URL(manifest.baseUrl, sourceUrl);
        } catch {
          // Keep the config directory as the only trusted base.
        }
      }
      if (!Array.isArray(manifest.variations)) return;
      for (const variation of manifest.variations) {
        if (!variation || typeof variation !== "object") continue;
        const variationUrl = (variation as { url?: unknown }).url;
        if (typeof variationUrl !== "string" || !variationUrl.trim()) {
          continue;
        }
        try {
          addPluginOrigin(
            originsByGuid,
            pluginGuid,
            new URL(variationUrl, pluginBaseUrl),
          );
        } catch {
          // Invalid variation URLs are not trusted.
        }
      }
    }),
  );
  return originsByGuid;
}

function makeHostSessionLabel(sessionId: string) {
  return sessionId
    .toLowerCase()
    .replace(/[^a-z0-9-]/g, "-")
    .replace(/-+/g, "-")
    .slice(0, 48);
}

function isFile(value: unknown): value is File {
  return (
    typeof value === "object" &&
    value !== null &&
    typeof (value as File).name === "string" &&
    typeof (value as Blob).arrayBuffer === "function"
  );
}

function describeInput(
  options: CreateOfficeEditorOptions,
  sessionId: string,
  ownerWindow: Window,
): InitialDescriptor {
  const emptyType = options.emptyType
    ? (normalizeExtension(options.emptyType, "docx") as OfficeEmptyType)
    : undefined;
  if (emptyType && !SUPPORTED_EMPTY_TYPES.includes(emptyType)) {
    throw new Error(`Unsupported empty document type: ${options.emptyType}`);
  }

  const input = options.file || options.buffer;
  let urlFileName: string | undefined;
  if (options.url) {
    try {
      urlFileName =
        new URL(options.url, ownerWindow.location.href).pathname
          .split("/")
          .pop() || "document.docx";
    } catch {
      urlFileName = "document.docx";
    }
  }

  const fileName =
    options.fileName ||
    (emptyType ? `New_Document.${emptyType}` : undefined) ||
    (isFile(input) ? input.name : undefined) ||
    urlFileName ||
    "document.docx";
  const fileType = emptyType || getFileExtension(fileName) || "docx";
  const mode = options.mode || (options.readonly ? "readonly" : "edit");
  const sourceKind: OfficeEditorSourceKind = emptyType
    ? "new-document"
    : options.url
      ? "url"
      : options.file
        ? "local-file"
        : "buffer";

  if (!emptyType && !options.url && !input) {
    throw new Error(
      "createOfficeEditor requires file, buffer, url, or emptyType",
    );
  }

  const context = { sessionId, fileName, fileType, mode };
  const hostUrlValue =
    typeof options.hostUrl === "function"
      ? options.hostUrl(context)
      : options.hostUrl;
  const hostUrl = new URL(hostUrlValue, ownerWindow.location.href);
  if (
    hostUrl.hostname === "localhost" ||
    hostUrl.hostname.endsWith(".localhost")
  ) {
    hostUrl.hostname = `host-${makeHostSessionLabel(sessionId)}.office.localhost`;
  }
  if (hostUrl.origin === ownerWindow.location.origin) {
    throw new Error(
      "createOfficeEditor requires hostUrl to be an independent origin",
    );
  }

  return {
    context,
    hostUrl,
    state: {
      id: sessionId,
      fileName,
      fileType,
      mode,
      readonly: mode !== "edit",
      dirty: false,
      sourceKind,
      status: "opening",
      destroyed: false,
    },
  };
}

function bytesToHex(bytes: ArrayBuffer) {
  return Array.from(new Uint8Array(bytes), (byte) =>
    byte.toString(16).padStart(2, "0"),
  ).join("");
}

/**
 * Returns the identity of the direct-embed adapter and its currently
 * registered static-resource coordinates. This is not the identity of the
 * informational hostUrl because no outer host document is loaded.
 */
export async function resolveOfficeEmbedHostIdentity(): Promise<OfficeHostIdentity> {
  const registeredManifestDigest =
    getOnlyOfficeStaticResourceManifestDigest();
  if (registeredManifestDigest) {
    return {
      packageVersion: ONLYOFFICE_EMBED_SDK_VERSION,
      hostBuildId: ONLYOFFICE_EMBED_HOST_BUILD_ID,
      assetManifestDigest: registeredManifestDigest,
    };
  }

  const subtle = globalThis.crypto?.subtle;
  if (!subtle) {
    throw new OfficeDirectEmbedCompatibilityError(
      "host-identity",
      "Web Crypto is required to identify the direct-embed runtime",
    );
  }
  // Compatibility fallback only: without a deployment-provided manifest
  // digest, bind identity to the registered immutable resource coordinates.
  // Consumers using identity as a release gate should always register the
  // actual manifest SHA-256 through registerOnlyOfficeStaticResource().
  const material = JSON.stringify({
    packageVersion: ONLYOFFICE_EMBED_SDK_VERSION,
    hostBuildId: ONLYOFFICE_EMBED_HOST_BUILD_ID,
    staticResource: getStaticResource(),
  });
  const digest = await subtle.digest(
    "SHA-256",
    new TextEncoder().encode(material),
  );
  return {
    packageVersion: ONLYOFFICE_EMBED_SDK_VERSION,
    hostBuildId: ONLYOFFICE_EMBED_HOST_BUILD_ID,
    assetManifestDigest: bytesToHex(digest),
  };
}

function identitiesEqual(
  left: OfficeHostIdentity,
  right: OfficeHostIdentity,
) {
  return (
    left.packageVersion === right.packageVersion &&
    left.hostBuildId === right.hostBuildId &&
    left.assetManifestDigest === right.assetManifestDigest
  );
}

function toError(error: unknown) {
  return error instanceof Error ? error : new Error(String(error));
}

function notifyOfficeEditorError(
  callback: CreateOfficeEditorOptions["onError"],
  error: Error,
  instance?: OfficeEditorInstance,
) {
  try {
    callback?.(error, instance);
  } catch (callbackError) {
    // Consumer diagnostics must never interrupt editor teardown or replace the
    // operation error that triggered the callback.
    console.error(
      "[onlyoffice-embed-sdk] onError callback failed",
      callbackError,
    );
  }
}

function interfaceThemeToOfficeTheme(
  theme: OfficeInterfaceTheme | undefined,
): OfficeThemeId {
  return theme === "dark" ? "theme-night" : "theme-white";
}

function makeFile(
  ownerWindow: Window,
  parts: BlobPart[],
  name: string,
  type: string,
) {
  const FileConstructor =
    (ownerWindow as Window & { File?: typeof File }).File || File;
  return new FileConstructor(parts, name, { type });
}

function makeReadyWaiter(
  ownerWindow: Window,
  containerId: string,
  manager: EditorManager,
): ReadyWaiter {
  let rejectPromise: (error: Error) => void = () => {};
  let settled = false;
  let timer = 0;
  const handler = (event: { instanceId?: string; manager?: object }) => {
    if (
      event.instanceId !== containerId ||
      event.manager !== manager ||
      settled
    ) {
      return;
    }
    settled = true;
    ownerWindow.clearTimeout(timer);
    onlyofficeEventbus.off(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, handler);
    resolvePromise();
  };
  let resolvePromise: () => void = () => {};
  const promise = new Promise<void>((resolve, reject) => {
    resolvePromise = resolve;
    rejectPromise = reject;
    onlyofficeEventbus.on(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, handler);
    timer = ownerWindow.setTimeout(() => {
      if (settled) return;
      settled = true;
      onlyofficeEventbus.off(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, handler);
      reject(
        new OfficeEditorStartupError(
          "document-ready",
          "Office editor startup exceeded five minutes",
        ),
      );
    }, STARTUP_TIMEOUT_MS);
  });
  return {
    promise,
    cancel(error = new Error("Editor was destroyed before it became ready")) {
      if (settled) return;
      settled = true;
      ownerWindow.clearTimeout(timer);
      onlyofficeEventbus.off(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, handler);
      rejectPromise(error);
    },
  };
}

class DirectEmbedOfficeEditor implements OfficeEditorInstance {
  readonly id: string;
  private readonly container: HTMLElement;
  private readonly ownerWindow: Window;
  private readonly options: CreateOfficeEditorOptions;
  private readonly descriptor: InitialDescriptor;
  private readonly generatedContainerId: boolean;
  private readonly containerId: string;
  private manager: EditorManager | null = null;
  private pluginBridge: OfficePluginBridge | null = null;
  private hostIdentity: OfficeHostIdentity | null = null;
  private state: OfficeEditorState;
  private mountPhase: OfficeEditorMountPhase = "waiting-for-activation";
  private mountError: Error | undefined;
  private activationPromise: Promise<OfficeEditorInstance> | null = null;
  private destroyPromise: Promise<void> | null = null;
  private readyWaiter: ReadyWaiter | null = null;
  private dirtyPollTimer: number | null = null;
  /** Keeps unsaved state latched when external persistence rejects. */
  private persistenceDirty = false;
  private savePromise: Promise<File> | null = null;
  private printPromise: Promise<File> | null = null;
  private destroyed = false;
  private returnsToPreview: boolean;
  private loadingEventHandler: ((data: LoadingChangeData) => void) | null = null;
  private loadingState: OfficeLoadingState = {
    loading: true,
    phase: "host-loading",
    resourceStatus: "checking",
    resourceDownload: false,
    transferredBytes: 0,
    resourceCount: 0,
  };

  constructor(
    container: HTMLElement,
    options: CreateOfficeEditorOptions,
    descriptor: InitialDescriptor,
  ) {
    this.id = descriptor.context.sessionId;
    this.container = container;
    this.ownerWindow = container.ownerDocument.defaultView || window;
    this.options = { ...options };
    this.descriptor = descriptor;
    this.state = { ...descriptor.state };
    this.returnsToPreview =
      descriptor.state.mode === "preview" || options.canReturnToPreview === true;
    this.generatedContainerId = !container.id;
    this.containerId =
      container.id || `onlyoffice-embed-${this.id.replace(/[^a-z0-9-]/gi, "-")}`;
    container.id = this.containerId;
    container.replaceChildren();
    container.classList.add("office-editor-host");
    container.style.width ||= "100%";
    container.style.height ||= "100%";
    container.style.minWidth ||= "0";
    container.style.minHeight ||= "0";
    this.ownerWindow.queueMicrotask(() => {
      if (!this.destroyed) this.notifyLoadingChange(this.loadingState);
    });
  }

  activate(): Promise<OfficeEditorInstance> {
    if (this.activationPromise) return this.activationPromise;
    if (this.destroyed) {
      return Promise.reject(new Error("Editor was destroyed before activation"));
    }

    this.mountPhase = "runtime-loading";
    this.updateLoadingState({ loading: true, phase: "runtime-loading" });
    this.activationPromise = this.activateInternal().catch(async (error) => {
      const normalized = toError(error);
      this.mountPhase = "error";
      this.mountError = normalized;
      this.state = { ...this.state, status: "error" };
      try {
        await this.cleanup(false);
      } catch (cleanupError) {
        console.error(
          "[onlyoffice-embed-sdk] failed to clean up editor startup",
          cleanupError,
        );
      }
      notifyOfficeEditorError(this.options.onError, normalized, this);
      throw normalized;
    });
    return this.activationPromise;
  }

  private async activateInternal(): Promise<OfficeEditorInstance> {
    const actualIdentity = await resolveOfficeEmbedHostIdentity();
    if (
      this.options.expectedHostIdentity &&
      !identitiesEqual(this.options.expectedHostIdentity, actualIdentity)
    ) {
      throw new OfficeHostIdentityMismatchError(
        this.options.expectedHostIdentity,
        actualIdentity,
      );
    }
    this.hostIdentity = actualIdentity;

    if (this.destroyed) {
      throw new Error("Editor was destroyed before activation completed");
    }

    const pluginOrigins = await resolveConfiguredPluginOrigins(
      this.ownerWindow,
      this.options.plugins,
    );
    if (this.destroyed) {
      throw new Error("Editor was destroyed before activation completed");
    }
    this.pluginBridge = new OfficePluginBridge(this.ownerWindow, {
      pluginGuids: this.options.plugins?.autostart ?? [],
      pluginOrigins,
      isAllowedSource: (source) => this.isOwnedPluginSource(source),
      onReady: (pluginGuid, editorType) => {
        if (this.destroyed) return;
        try {
          this.options.onPluginReady?.(pluginGuid, editorType, this);
        } catch (error) {
          notifyOfficeEditorError(this.options.onError, toError(error), this);
        }
      },
    });

    const createOptions = await this.createEditorOptions();
    if (this.destroyed) {
      throw new Error("Editor was destroyed before activation completed");
    }

    this.manager = editorManagerFactory.get(this.container);
    this.loadingEventHandler = (data) => {
      if (data.manager !== this.manager || this.destroyed) return;
      this.updateLoadingState({
        loading: data.loading,
        phase: data.phase,
        resourceStatus: data.resourceStatus,
        resourceDownload: data.resourceDownload,
        transferredBytes: data.transferredBytes,
        resourceCount: data.resourceCount,
      });
    };
    onlyofficeEventbus.on(
      ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE,
      this.loadingEventHandler,
    );
    this.readyWaiter = makeReadyWaiter(
      this.ownerWindow,
      this.containerId,
      this.manager,
    );
    try {
      await Promise.all([
        this.manager.create({
          ...createOptions,
          containerId: this.containerId,
          container: this.container,
          editorManager: this.manager,
          readOnly: this.state.readonly,
          mode: this.state.mode,
          spellcheck: this.options.spellcheck,
          lang: this.options.lang,
          theme: interfaceThemeToOfficeTheme(this.options.interfaceTheme),
          plugins: this.options.plugins,
          onUserSave: (output) => this.persistNativeSavedFile(output),
          onDownloadOutput: (output) =>
            this.handleNativeDownloadOutput(output),
        }),
        this.readyWaiter.promise,
      ]);
    } finally {
      // A create failure must not leave the DOCUMENT_READY listener and its
      // five-minute timeout alive. Calling cancel after success is a no-op.
      this.readyWaiter?.cancel(
        new Error("Office editor startup ended before document ready"),
      );
      this.readyWaiter = null;
    }

    if (this.destroyed) {
      throw new Error("Editor was destroyed before activation completed");
    }

    this.mountPhase = "ready";
    this.updateState({ status: "ready", dirty: this.manager.isDirty() });
    this.updateLoadingState({ loading: false, phase: "ready", resourceDownload: false });
    this.startDirtyPolling();
    try {
      this.options.onReady?.(this);
    } catch (error) {
      notifyOfficeEditorError(this.options.onError, toError(error), this);
    }
    return this;
  }

  private async createEditorOptions() {
    const base = {
      fileName: this.state.fileName,
      fileType: this.state.fileType,
      isNew: false,
    };
    if (this.options.emptyType) {
      return { ...base, isNew: true };
    }
    if (this.options.url) {
      return {
        ...base,
        url: this.options.url,
        loader: async (url: string) => {
          const response = await this.ownerWindow.fetch(
            url,
            this.options.fetchOptions,
          );
          if (!response.ok) {
            throw new Error(
              `Failed to fetch document: ${response.status} ${response.statusText}`,
            );
          }
          return response.arrayBuffer();
        },
      };
    }

    const input = this.options.file || this.options.buffer;
    if (!input) {
      throw new Error(
        "createOfficeEditor requires file, buffer, url, or emptyType",
      );
    }
    let part: BlobPart;
    if (input instanceof Uint8Array) {
      part = input.slice().buffer;
    } else if (input instanceof ArrayBuffer) {
      part = input.slice(0);
    } else {
      part = input;
    }
    const file = isFile(input) && input.name === this.state.fileName
      ? input
      : makeFile(
          this.ownerWindow,
          [part],
          this.state.fileName,
          input instanceof Blob && input.type
            ? input.type
            : getOnlyOfficeMimeType(this.state.fileType),
        );
    return { ...base, file };
  }

  private isOwnedPluginSource(source: WindowProxy) {
    // DocsAPI replaces/wraps the supplied container during mount. Resolve the
    // live frame through the owning manager, which still performs an exact
    // WindowProxy parent comparison and rejects same-origin siblings.
    return this.manager?.ownsPluginSource(source) ?? false;
  }

  getMountState(): OfficeEditorMountState {
    return {
      id: this.id,
      origin: this.descriptor.hostUrl.origin,
      phase: this.mountPhase,
      ...(this.mountError ? { error: this.mountError } : {}),
    };
  }

  getLoadingState(): OfficeLoadingState {
    return { ...this.loadingState };
  }

  invokePlugin(pluginGuid: string, payload: unknown): Promise<unknown> {
    if (this.destroyed || !this.pluginBridge) {
      return Promise.reject(new Error("Editor is not open"));
    }
    return this.pluginBridge.invoke(pluginGuid, payload);
  }

  save(targetExt?: string): Promise<File> {
    if (this.destroyed || !this.manager || this.state.status !== "ready") {
      return Promise.reject(new Error("Editor is not open"));
    }
    if (this.state.readonly) {
      return Promise.reject(new Error("Current document is readonly"));
    }
    if (this.savePromise) {
      return Promise.reject(
        new Error("A save request is already in progress for this editor"),
      );
    }

    const manager = this.manager;
    let operation: Promise<File>;
    operation = this.saveInternal(manager, targetExt).finally(() => {
      if (this.savePromise === operation) this.savePromise = null;
    });
    this.savePromise = operation;
    return operation;
  }

  exportCopy(targetExt?: string): Promise<File> {
    if (this.destroyed || !this.manager || this.state.status !== "ready") {
      return Promise.reject(new Error("Editor is not open"));
    }
    if (this.savePromise) {
      return Promise.reject(
        new Error("A save request is already in progress for this editor"),
      );
    }
    const manager = this.manager;
    let operation: Promise<File>;
    operation = (async () => {
      const exported = await manager.exportCopy();
      const extension = normalizeExtension(targetExt, this.state.fileType);
      const converted = await convertBinToDocument(
        exported.binData,
        exported.fileName || this.state.fileName,
        extension,
        exported.media,
        exported.themes,
        manager.getLogger(),
      );
      return makeFile(
        this.ownerWindow,
        [converted.data as BlobPart],
        converted.fileName,
        getOnlyOfficeMimeType(extension),
      );
    })().finally(() => {
      if (this.savePromise === operation) this.savePromise = null;
    });
    this.savePromise = operation;
    return operation;
  }

  async print(): Promise<File> {
    if (this.printPromise) return this.printPromise;
    if (this.destroyed || !this.manager || this.state.status !== "ready") {
      throw new Error("Editor is not open");
    }
    if (this.savePromise) {
      throw new Error("A save request is already in progress for this editor");
    }
    let operation: Promise<File>;
    operation = (async () => {
      const file = await this.exportPrintPdfFile();
      await printOfficePdfFile({
        ownerWindow: this.ownerWindow,
        ownerDocument: this.container.ownerDocument,
        file,
      });
      return file;
    })().finally(() => {
      if (this.printPromise === operation) this.printPromise = null;
    });
    this.printPromise = operation;
    return operation;
  }

  async exportPrintPdfFile(): Promise<File> {
    if (this.destroyed || !this.manager || this.state.status !== "ready") {
      throw new Error("Editor is not open");
    }
    if (this.savePromise) {
      throw new Error("A save request is already in progress for this editor");
    }
    const manager = this.manager;
    let operation: Promise<File>;
    operation = (async () => {
      const bytes = await manager.exportPrintPdf();
      const baseName = this.state.fileName.replace(/\.[^.]*$/, "") || "document";
      return makeFile(
        this.ownerWindow,
        [bytes],
        `${baseName}.pdf`,
        "application/pdf",
      );
    })().finally(() => {
      if (this.savePromise === operation) this.savePromise = null;
    });
    this.savePromise = operation;
    return operation;
  }

  private async saveInternal(manager: EditorManager, targetExt?: string) {
    // export() clears EditorManager.dirty before the consumer persistence
    // callback settles. Preserve the pre-save dirty state so the 100ms poll
    // cannot announce a successful save while onSave is still pending.
    if (this.persistenceDirty || this.state.dirty || manager.isDirty()) {
      this.persistenceDirty = true;
    }
    try {
      const exported = await manager.export();
      const extension = normalizeExtension(targetExt, this.state.fileType);
      const converted = await convertBinToDocument(
        exported.binData,
        exported.fileName || this.state.fileName,
        extension,
        exported.media,
        exported.themes,
        manager.getLogger(),
      );
      const file = makeFile(
        this.ownerWindow,
        [converted.data as BlobPart],
        converted.fileName,
        getOnlyOfficeMimeType(extension),
      );
      await this.persistSavedFile(file);
      this.persistenceDirty = false;
      if (!this.destroyed) {
        // A new edit can arrive while onSave is awaiting external storage.
        // Settle only the captured generation; the manager remains dirty for
        // edits made after export and must keep the facade dirty as well.
        this.updateState({ dirty: manager.isDirty() });
      }
      return file;
    } catch (error) {
      const normalized = toError(error);
      this.persistenceDirty = true;
      if (!this.destroyed) this.updateState({ dirty: true });
      notifyOfficeEditorError(this.options.onError, normalized, this);
      throw normalized;
    }
  }

  private persistNativeSavedFile(output: EditorDocumentExport) {
    if (this.destroyed || !this.manager || this.state.status !== "ready") {
      return Promise.reject(new Error("Editor is not open"));
    }
    if (this.savePromise) {
      return Promise.reject(
        new Error("A save request is already in progress for this editor"),
      );
    }

    const manager = this.manager;
    let operation: Promise<File>;
    operation = (async () => {
      if (this.persistenceDirty || this.state.dirty || manager.isDirty()) {
        this.persistenceDirty = true;
      }
      const extension = normalizeExtension(undefined, this.state.fileType);
      const converted = await convertBinToDocument(
        output.binData,
        output.fileName || this.state.fileName,
        extension,
        output.media,
        output.themes,
        manager.getLogger(),
      );
      const file = makeFile(
        this.ownerWindow,
        [converted.data as BlobPart],
        converted.fileName,
        getOnlyOfficeMimeType(extension),
      );
      await this.persistSavedFile(file);
      this.persistenceDirty = false;
      return file;
    })()
      .catch((error) => {
        const normalized = toError(error);
        this.persistenceDirty = true;
        if (!this.destroyed) this.updateState({ dirty: true });
        notifyOfficeEditorError(this.options.onError, normalized, this);
        throw normalized;
      })
      .finally(() => {
        if (this.savePromise === operation) this.savePromise = null;
      });
    this.savePromise = operation;
    return operation.then(() => undefined);
  }

  private async persistSavedFile(file: File) {
    const behavior = this.options.saveBehavior || "auto";
    const shouldCallCallback =
      behavior === "callback" ||
      (behavior === "auto" &&
        (this.state.sourceKind !== "new-document" ||
          Boolean(this.options.onSave)));
    let handled = false;
    if (shouldCallCallback) {
      if (!this.options.onSave) {
        throw new Error(
          'A save callback is required for this document source. Provide onSave or use saveBehavior: "download".',
        );
      }
      handled = (await this.options.onSave(file, this)) === true;
    }

    if (
      behavior === "download" ||
      (behavior === "auto" &&
        this.state.sourceKind === "new-document" &&
        !handled)
    ) {
      this.downloadFile(file);
    }
  }

  private async handleNativeDownloadOutput(output: EditorDownloadOutput) {
    const callback =
      output.kind === "save-as"
        ? this.options.onSaveAs
        : this.options.onDownload;
    if (!callback) return false;

    try {
      const file = makeFile(
        this.ownerWindow,
        [output.data as BlobPart],
        output.fileName,
        getOnlyOfficeMimeType(output.fileType),
      );
      await callback(file, this);
    } catch (error) {
      notifyOfficeEditorError(this.options.onError, toError(error), this);
    }
    return true;
  }

  private downloadFile(file: File) {
    const URLConstructor =
      (this.ownerWindow as Window & { URL?: typeof URL }).URL || URL;
    const objectUrl = URLConstructor.createObjectURL(file);
    const anchor = this.container.ownerDocument.createElement("a");
    anchor.href = objectUrl;
    anchor.download = file.name || "document";
    anchor.style.display = "none";
    this.container.ownerDocument.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    this.ownerWindow.setTimeout(
      () => URLConstructor.revokeObjectURL(objectUrl),
      0,
    );
  }

  async confirmSaveToNewFormat(
    options?: OfficeSaveToNewFormatConfirmationOptions,
  ): Promise<boolean> {
    if (this.destroyed) throw new Error("Editor is not open");
    // The independent host used an ONLYOFFICE modal. Direct embed deliberately
    // falls back to the browser confirmation API; dontshow cannot be persisted.
    const message =
      options?.message ||
      "This document will be saved in a new Office format. Continue?";
    return this.ownerWindow.confirm(message);
  }

  setReadonly(readonly: boolean): void {
    const previous = {
      readonly: this.state.readonly,
      mode: this.state.mode,
    };
    const mode =
      this.returnsToPreview && readonly
        ? "preview"
        : readonly
          ? "readonly"
          : "edit";
    this.updateState({ readonly, mode });
    if (!this.destroyed && this.manager) {
      const changesShell = previous.mode === "preview" || mode === "preview";
      const operation = changesShell
        ? this.manager.setMode(mode)
        : this.manager.setReadOnly(readonly);
      void operation.catch((error) => {
        this.updateState(previous);
        notifyOfficeEditorError(this.options.onError, toError(error), this);
      });
    }
  }

  setInterfaceTheme(theme: OfficeInterfaceTheme): Promise<void> {
    this.options.interfaceTheme = theme;
    if (this.destroyed || !this.manager) {
      return Promise.resolve();
    }
    const operation = this.manager.setTheme(interfaceThemeToOfficeTheme(theme));
    void operation.catch((error) =>
      notifyOfficeEditorError(this.options.onError, toError(error), this),
    );
    return operation;
  }

  async setLanguage(lang: string): Promise<void> {
    if (this.destroyed || !this.manager) {
      throw new Error("Editor is not open");
    }
    if (!lang.trim()) throw new Error("OnlyOffice language is required");
    this.options.lang = lang;
    await this.manager.setLanguage(lang as OnlyOfficeLang);
  }

  getState(): OfficeEditorState {
    return { ...this.state };
  }

  getHostIdentity(): OfficeHostIdentity {
    if (!this.hostIdentity) {
      throw new Error(
        "Office host identity is unavailable before the host is ready",
      );
    }
    return { ...this.hostIdentity };
  }

  destroy(): Promise<void> {
    if (this.destroyPromise) return this.destroyPromise;
    this.mountPhase = "destroyed";
    this.destroyPromise = this.cleanup(true);
    return this.destroyPromise;
  }

  private async cleanup(markDestroyed: boolean) {
    this.destroyed = true;
    this.readyWaiter?.cancel();
    this.readyWaiter = null;
    if (this.loadingEventHandler) {
      onlyofficeEventbus.off(
        ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE,
        this.loadingEventHandler,
      );
      this.loadingEventHandler = null;
    }
    if (this.dirtyPollTimer !== null) {
      this.ownerWindow.clearInterval(this.dirtyPollTimer);
      this.dirtyPollTimer = null;
    }
    this.pluginBridge?.destroy();
    this.pluginBridge = null;
    if (this.manager) {
      editorManagerFactory.destroy(this.container);
      this.manager = null;
    }
    activeInstances.delete(this.id);
    const containerIds = getActiveContainerIds(this.container.ownerDocument);
    if (containerIds.get(this.containerId) === this) {
      containerIds.delete(this.containerId);
    }
    if (activeContainers.get(this.container) === this) {
      activeContainers.delete(this.container);
      this.container.replaceChildren();
      this.container.classList.remove("office-editor-host");
      if (this.generatedContainerId && this.container.id === this.containerId) {
        this.container.removeAttribute("id");
      }
    }
    if (markDestroyed) {
      this.persistenceDirty = false;
      this.updateState({ status: "destroyed", destroyed: true, dirty: false });
      this.updateLoadingState({
        loading: false,
        phase: "destroyed",
        resourceDownload: false,
      });
    }
  }

  private updateLoadingState(patch: Partial<OfficeLoadingState>) {
    const next = { ...this.loadingState, ...patch };
    const changed = Object.keys(patch).some(
      (key) =>
        next[key as keyof OfficeLoadingState] !==
        this.loadingState[key as keyof OfficeLoadingState],
    );
    this.loadingState = next;
    if (changed) this.notifyLoadingChange(next);
  }

  private notifyLoadingChange(state: OfficeLoadingState) {
    try {
      void Promise.resolve(
        this.options.onLoadingChange?.({ ...state }, this),
      ).catch((error) =>
        notifyOfficeEditorError(this.options.onError, toError(error), this),
      );
    } catch (error) {
      notifyOfficeEditorError(this.options.onError, toError(error), this);
    }
  }

  private startDirtyPolling() {
    if (this.dirtyPollTimer !== null) return;
    this.dirtyPollTimer = this.ownerWindow.setInterval(() => {
      if (!this.manager || this.destroyed) return;
      this.updateState({
        dirty: this.persistenceDirty || this.manager.isDirty(),
      });
    }, DIRTY_POLL_INTERVAL_MS);
  }

  private updateState(patch: Partial<OfficeEditorState>) {
    const previous = this.state;
    const next = { ...previous, ...patch };
    const dirtyChanged = next.dirty !== previous.dirty;
    const stateChanged = Object.keys(patch).some(
      (key) =>
        next[key as keyof OfficeEditorState] !==
        previous[key as keyof OfficeEditorState],
    );
    this.state = next;

    if (dirtyChanged) {
      void Promise.resolve(this.options.onDirtyChange?.(next.dirty, this)).catch(
        (error) =>
          notifyOfficeEditorError(this.options.onError, toError(error), this),
      );
    }
    if (stateChanged) {
      void Promise.resolve(this.options.onStateChange?.({ ...next }, this)).catch(
        (error) =>
          notifyOfficeEditorError(this.options.onError, toError(error), this),
      );
    }
  }
}

export function mountOfficeEditor(
  container: HTMLElement,
  options: CreateOfficeEditorOptions,
): OfficeEditorMount {
  if (!isHTMLElementContainer(container)) {
    throw new Error("mountOfficeEditor requires an HTMLElement container");
  }

  try {
    const ownerWindow = container.ownerDocument.defaultView || window;
    const sessionId = makeSessionId();
    const descriptor = describeInput(options, sessionId, ownerWindow);
    const existing = activeContainers.get(container);
    const containerIds = getActiveContainerIds(container.ownerDocument);
    const containerIdOwner = container.id
      ? containerIds.get(container.id)
      : undefined;
    if (existing || containerIdOwner) {
      throw new OfficeHostIsolationError(
        descriptor.hostUrl.origin,
        (existing || containerIdOwner)!.id,
        sessionId,
      );
    }
    const instance = new DirectEmbedOfficeEditor(
      container,
      options,
      descriptor,
    );
    activeContainers.set(container, instance);
    containerIds.set(container.id, instance);
    activeInstances.set(instance.id, instance);
    return {
      id: instance.id,
      activate: () => instance.activate(),
      destroy: () => instance.destroy(),
      getState: () => instance.getMountState(),
      getLoadingState: () => instance.getLoadingState(),
    };
  } catch (error) {
    const normalized = toError(error);
    notifyOfficeEditorError(options.onError, normalized);
    throw normalized;
  }
}

export async function createOfficeEditor(
  container: HTMLElement,
  options: CreateOfficeEditorOptions,
): Promise<OfficeEditorInstance> {
  const mount = mountOfficeEditor(container, options);
  try {
    return await mount.activate();
  } catch (error) {
    await mount.destroy();
    throw toError(error);
  }
}

export function loadOfficeEditorApi(): Promise<void> {
  return initializeOnlyOffice();
}

export function getActiveOfficeEditorCount(): number {
  return activeInstances.size;
}
