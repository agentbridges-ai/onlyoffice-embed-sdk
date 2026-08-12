import type {
  CreateOfficeEditorOptions as DirectCreateOfficeEditorOptions,
  OfficeEditorInstance as DirectOfficeEditorInstance,
  OfficeEditorMount as DirectOfficeEditorMount,
  OfficeEditorMountState,
  OfficeEditorState,
  OfficeHostIdentity,
  OfficeHostUrlContext,
  OfficeInterfaceTheme,
  OfficeSaveToNewFormatConfirmationOptions,
} from "./editor";
import {
  OfficeEditorStartupError,
  OfficeHostIdentityMismatchError,
  OfficeHostIsolationError,
} from "./editor";
import {
  ONLYOFFICE_EMBED_HOST_ASSET_DIGEST,
  ONLYOFFICE_EMBED_HOST_BUILD_ID,
  ONLYOFFICE_EMBED_HOST_MANIFEST,
  ONLYOFFICE_EMBED_SDK_VERSION,
} from "./version";
import {
  COMPAT_SUBFRAME_PROTOCOL_SOURCE,
  COMPAT_SUBFRAME_PROTOCOL_VERSION,
  COMPAT_SUBFRAME_RUNTIME,
  isCompatSubframeChildMessage,
  isCompatSubframeSerializedError,
  serializeCompatSubframeError,
  type CompatSubframeAction,
  type CompatSubframeCallbackRequest,
  type CompatSubframeCallbackResponse,
  type CompatSubframeChildMessage,
  type CompatSubframeDocument,
  type CompatSubframeEventMessage,
  type CompatSubframeOpenPayload,
  type CompatSubframeOpenResult,
  type CompatSubframeParentMessage,
  type CompatSubframeRequest,
  type CompatSubframeResponse,
  type CompatSubframeSerializedError,
} from "./subframe-protocol";
import { openOfficePrintWindow, printOfficePdfFile } from "./print";

export * from "./subframe-protocol";
export {
  OfficeEditorStartupError,
  OfficeHostIdentityMismatchError,
  OfficeHostIsolationError,
} from "./editor";
export {
  ONLYOFFICE_EMBED_HOST_ASSET_DIGEST,
  ONLYOFFICE_EMBED_HOST_BUILD_ID,
  ONLYOFFICE_EMBED_HOST_MANIFEST,
  ONLYOFFICE_EMBED_SDK_VERSION,
} from "./version";
export type {
  OfficeEditorMode,
  OfficeEditorInput,
  OfficeEditorMountPhase,
  OfficeEditorMountState,
  OfficeEditorSourceKind,
  OfficeEditorState,
  OfficeHostIdentity,
  OfficeHostUrlContext,
  OfficeHostUrlResolver,
  OfficeInterfaceTheme,
  OfficePluginOptions,
  OfficeDownloadCallbackResult,
  OfficeSaveBehavior,
  OfficeSaveAsCallbackResult,
  OfficeSaveCallbackResult,
  OfficeSaveToNewFormatConfirmationOptions,
} from "./editor";

const DEFAULT_REQUEST_TIMEOUT_MS = 75_000;
const DEFAULT_STARTUP_TIMEOUT_MS = 5 * 60_000;
const DEFAULT_CALLBACK_TIMEOUT_MS = 75_000;
const DEFAULT_DESTROY_TIMEOUT_MS = 10_000;
const DEFAULT_SUBFRAME_PATH = "/subframe";

export const OFFICE_SUBFRAME_SLOTS = [
  "rat",
  "ox",
  "tiger",
  "rabbit",
  "dragon",
  "snake",
  "horse",
  "goat",
  "monkey",
  "rooster",
  "dog",
  "pig",
] as const;
export type OfficeSubframeSlot = (typeof OFFICE_SUBFRAME_SLOTS)[number];

const OFFICE_SUBFRAME_SLOT_SET = new Set<string>(OFFICE_SUBFRAME_SLOTS);
const PRODUCTION_HOST = "onlyoffice.agent-bridges.com";
const LOCAL_HOST = "onlyoffice.localhost";

export const HOSTED_COMPAT_SUBFRAME_IDENTITY: OfficeHostIdentity = Object.freeze({
  packageVersion: ONLYOFFICE_EMBED_SDK_VERSION,
  hostBuildId: ONLYOFFICE_EMBED_HOST_BUILD_ID,
  assetManifestDigest: ONLYOFFICE_EMBED_HOST_ASSET_DIGEST,
});

export function isOfficeSubframeSlot(value: string): value is OfficeSubframeSlot {
  return OFFICE_SUBFRAME_SLOT_SET.has(value);
}

/** Build one of the fixed production/local isolated origins. */
export function getOfficeSubframeOrigin(
  slot: OfficeSubframeSlot,
  base: string | URL = `https://${PRODUCTION_HOST}`,
) {
  if (!isOfficeSubframeSlot(slot)) {
    throw new OfficeSubframeConfigurationError(`Unknown Office subframe slot: ${slot}`);
  }
  const url = new URL(String(base), `https://${PRODUCTION_HOST}/`);
  if (url.username || url.password) {
    throw new OfficeSubframeConfigurationError(
      "Office subframe base URLs must not contain credentials",
    );
  }
  const rootHost =
    url.hostname === LOCAL_HOST || url.hostname.endsWith(`.${LOCAL_HOST}`)
      ? LOCAL_HOST
      : url.hostname === PRODUCTION_HOST ||
          url.hostname.endsWith(`.${PRODUCTION_HOST}`)
        ? PRODUCTION_HOST
        : null;
  if (!rootHost) {
    throw new OfficeSubframeConfigurationError(
      `Unsupported Office subframe base host: ${url.hostname}`,
    );
  }
  if (rootHost === PRODUCTION_HOST && (url.protocol !== "https:" || url.port)) {
    throw new OfficeSubframeConfigurationError(
      "Production Office subframes require HTTPS on the default port",
    );
  }
  if (rootHost === LOCAL_HOST && !/^https?:$/.test(url.protocol)) {
    throw new OfficeSubframeConfigurationError(
      "Local Office subframes require HTTP(S)",
    );
  }
  url.hostname = `${slot}.${rootHost}`;
  url.pathname = "/";
  url.search = "";
  url.hash = "";
  return url.origin;
}

export interface CreateOfficeEditorOptions
  extends DirectCreateOfficeEditorOptions {
  /** Path served by the embed-sdk site. hostUrl supplies the isolated origin. */
  subframePath?: string;
  /**
   * Shared static-resource origin. Defaults to the canonical production/local
   * root while the outer editor iframe remains on its isolated zodiac origin.
   */
  resourceOrigin?: string | URL;
  requestTimeoutMs?: number;
  startupTimeoutMs?: number;
  callbackTimeoutMs?: number;
  frameTitle?: string;
  frameClassName?: string;
}

export interface OfficeEditorInstance extends DirectOfficeEditorInstance {
  /** Explicitly request a Save As export from the isolated runtime. */
  saveAs(targetExt?: string): Promise<File>;
  /** Explicitly request a download export without crossing editor internals. */
  download(targetExt?: string): Promise<File>;
}

export interface OfficeEditorMount
  extends Omit<DirectOfficeEditorMount, "activate"> {
  activate(): Promise<OfficeEditorInstance>;
}

export class OfficeSubframeConfigurationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "OfficeSubframeConfigurationError";
  }
}

export class OfficeSubframeRpcError extends Error {
  readonly remoteName: string;
  readonly code?: string;

  constructor(error: CompatSubframeSerializedError) {
    super(error.message);
    this.name = "OfficeSubframeRpcError";
    this.remoteName = error.name;
    this.code = error.code;
    if (error.stack) this.stack = error.stack;
  }
}

type PendingRequest = {
  action: CompatSubframeAction;
  message: CompatSubframeRequest;
  resolve: (value: unknown) => void;
  reject: (error: Error) => void;
  timer: number;
  sent: boolean;
  transfer: Transferable[];
};

type UrlDocumentInput = {
  kind: "url";
  url: string;
  fileName: string;
  fetchOptions?: RequestInit;
};

type DescribedDocument = CompatSubframeDocument | UrlDocumentInput;

let nextInstanceId = 1;
const activeMounts = new WeakMap<HTMLElement, CompatSubframeMount>();
const activeInstances = new Set<CompatSubframeMount>();
const activeOrigins = new Map<string, CompatSubframeMount>();

function createOpaqueId(prefix: string, ownerWindow: Window) {
  const crypto = ownerWindow.crypto;
  if (typeof crypto?.randomUUID === "function") {
    return `${prefix}-${crypto.randomUUID()}`;
  }
  if (!crypto?.getRandomValues) {
    throw new OfficeSubframeConfigurationError(
      "A cryptographically secure random source is required",
    );
  }
  const bytes = crypto.getRandomValues(new Uint8Array(24));
  return `${prefix}-${Array.from(bytes, (value) =>
    value.toString(16).padStart(2, "0"),
  ).join("")}`;
}

function toError(error: unknown) {
  if (error instanceof Error) return error;
  if (
    error &&
    typeof error === "object" &&
    "message" in error &&
    typeof error.message === "string"
  ) {
    const normalized = new Error(error.message);
    if ("name" in error && typeof error.name === "string") {
      normalized.name = error.name;
    }
    return normalized;
  }
  return new Error(String(error));
}

function positiveTimeout(value: number | undefined, fallback: number) {
  return Number.isFinite(value) && (value as number) > 0
    ? (value as number)
    : fallback;
}

function extensionOf(fileName: string) {
  const name = fileName.split(/[?#]/, 1)[0] || fileName;
  const index = name.lastIndexOf(".");
  return index > -1 ? name.slice(index + 1).toLowerCase() : "";
}

function nameFromUrl(value: string, ownerWindow: Window) {
  try {
    const pathname = new URL(value, ownerWindow.location.href).pathname;
    const name = pathname.split("/").filter(Boolean).at(-1);
    return name ? decodeURIComponent(name) : "Document.docx";
  } catch {
    return "Document.docx";
  }
}

function describeDocument(
  options: CreateOfficeEditorOptions,
  ownerWindow: Window,
): { document: DescribedDocument; fileName: string; fileType: string } {
  const sources = [options.file, options.buffer, options.url, options.emptyType].filter(
    (value) => value !== undefined,
  );
  if (sources.length !== 1) {
    throw new OfficeSubframeConfigurationError(
      "Exactly one of file, buffer, url, or emptyType is required",
    );
  }

  if (options.file !== undefined) {
    const fileName =
      options.fileName ||
      (typeof (options.file as { name?: unknown }).name === "string"
        ? (options.file as File).name
        : "Document.docx");
    return {
      document: { kind: "file", file: options.file, fileName },
      fileName,
      fileType: extensionOf(fileName),
    };
  }
  if (options.buffer !== undefined) {
    const fileName = options.fileName || "Document.docx";
    return {
      document: { kind: "buffer", buffer: options.buffer, fileName },
      fileName,
      fileType: extensionOf(fileName),
    };
  }
  if (options.url !== undefined) {
    const fileName = options.fileName || nameFromUrl(options.url, ownerWindow);
    return {
      document: {
        kind: "url",
        url: options.url,
        fileName,
        ...(options.fetchOptions ? { fetchOptions: options.fetchOptions } : {}),
      },
      fileName,
      fileType: extensionOf(fileName),
    };
  }

  const emptyType = options.emptyType!;
  const fileName = options.fileName || `New Document.${emptyType}`;
  return {
    document: { kind: "empty", emptyType, fileName },
    fileName,
    fileType: emptyType,
  };
}

function resolveSubframeUrl(
  options: CreateOfficeEditorOptions,
  context: OfficeHostUrlContext,
  instanceId: string,
  sessionToken: string,
  ownerWindow: Window,
) {
  const resolved =
    typeof options.hostUrl === "function"
      ? options.hostUrl(context)
      : options.hostUrl;
  const hostUrl = new URL(String(resolved), ownerWindow.location.href);
  if (
    !/^https?:$/.test(hostUrl.protocol) ||
    hostUrl.origin === "null" ||
    hostUrl.username ||
    hostUrl.password
  ) {
    throw new OfficeSubframeConfigurationError(
      "hostUrl must resolve to an HTTP(S) origin",
    );
  }
  const hostParts = hostUrl.hostname.split(".");
  const slot = hostParts[0];
  const isProduction =
    hostParts.length === 4 &&
    isOfficeSubframeSlot(slot) &&
    hostParts.slice(1).join(".") === PRODUCTION_HOST &&
    hostUrl.protocol === "https:" &&
    hostUrl.port === "";
  const isLocal =
    hostParts.length === 3 &&
    isOfficeSubframeSlot(slot) &&
    hostParts.slice(1).join(".") === LOCAL_HOST &&
    /^https?:$/.test(hostUrl.protocol);
  if (!isProduction && !isLocal) {
    throw new OfficeSubframeConfigurationError(
      "hostUrl must use one of the 12 fixed onlyoffice.agent-bridges.com or onlyoffice.localhost subframe origins",
    );
  }
  if (hostUrl.origin === ownerWindow.location.origin) {
    throw new OfficeSubframeConfigurationError(
      "hostUrl must use an origin isolated from the parent application",
    );
  }

  const subframeUrl = new URL(options.subframePath || DEFAULT_SUBFRAME_PATH, hostUrl);
  if (subframeUrl.origin !== hostUrl.origin) {
    throw new OfficeSubframeConfigurationError(
      "subframePath must stay on the hostUrl origin",
    );
  }
  subframeUrl.searchParams.set("runtime", COMPAT_SUBFRAME_RUNTIME);
  subframeUrl.searchParams.set("instance", instanceId);
  subframeUrl.searchParams.set("session", sessionToken);
  subframeUrl.searchParams.set("parentOrigin", ownerWindow.location.origin);
  const defaultResourceOrigin = new URL(hostUrl.origin);
  defaultResourceOrigin.hostname = isProduction
    ? PRODUCTION_HOST
    : LOCAL_HOST;
  const resourceOrigin = options.resourceOrigin
    ? new URL(String(options.resourceOrigin), ownerWindow.location.href)
    : defaultResourceOrigin;
  const isProductionResourceOrigin =
    resourceOrigin.hostname === PRODUCTION_HOST &&
    resourceOrigin.protocol === "https:" &&
    resourceOrigin.port === "";
  const isLocalResourceOrigin =
    resourceOrigin.hostname === LOCAL_HOST &&
    /^https?:$/.test(resourceOrigin.protocol) &&
    resourceOrigin.port === hostUrl.port;
  if (
    resourceOrigin.username ||
    resourceOrigin.password ||
    resourceOrigin.pathname !== "/" ||
    resourceOrigin.search ||
    resourceOrigin.hash ||
    (isProduction
      ? !isProductionResourceOrigin
      : !isLocalResourceOrigin)
  ) {
    throw new OfficeSubframeConfigurationError(
      "resourceOrigin must use the canonical onlyoffice.agent-bridges.com or onlyoffice.localhost origin",
    );
  }
  return { hostUrl, resourceOrigin, subframeUrl };
}

function isOfficeEditorState(value: unknown): value is OfficeEditorState {
  if (!value || typeof value !== "object") return false;
  const state = value as Partial<OfficeEditorState>;
  return (
    typeof state.id === "string" &&
    typeof state.fileName === "string" &&
    typeof state.fileType === "string" &&
    ["edit", "readonly", "preview"].includes(String(state.mode)) &&
    typeof state.readonly === "boolean" &&
    typeof state.dirty === "boolean" &&
    ["local-file", "new-document", "buffer", "url"].includes(
      String(state.sourceKind),
    ) &&
    ["opening", "ready", "destroyed", "error"].includes(
      String(state.status),
    ) &&
    typeof state.destroyed === "boolean"
  );
}

function isOfficeHostIdentity(value: unknown): value is OfficeHostIdentity {
  if (!value || typeof value !== "object") return false;
  const identity = value as Partial<OfficeHostIdentity>;
  return (
    typeof identity.packageVersion === "string" &&
    typeof identity.hostBuildId === "string" &&
    typeof identity.assetManifestDigest === "string" &&
    /^[a-f0-9]{64}$/i.test(identity.assetManifestDigest)
  );
}

function isFileLike(value: unknown) {
  return (
    value !== null &&
    typeof value === "object" &&
    typeof (value as Blob).size === "number" &&
    typeof (value as Blob).type === "string" &&
    typeof (value as Blob).arrayBuffer === "function" &&
    typeof (value as File).name === "string"
  );
}

function isValidActionResult(action: CompatSubframeAction, value: unknown) {
  switch (action) {
    case "open":
      return (
        value !== null &&
        typeof value === "object" &&
        isOfficeEditorState((value as CompatSubframeOpenResult).state) &&
        isOfficeHostIdentity((value as CompatSubframeOpenResult).hostIdentity)
      );
    case "set-readonly":
    case "set-theme":
    case "set-language":
    case "get-state":
      return isOfficeEditorState(value);
    case "save":
    case "save-as":
    case "download":
    case "print":
      return isFileLike(value);
    case "confirm-save-to-new-format":
      return typeof value === "boolean";
    case "get-host-identity":
      return isOfficeHostIdentity(value);
    case "destroy":
      return value === null || value === undefined;
    case "invoke-plugin":
      return true;
  }
}

function remoteError(error?: CompatSubframeSerializedError) {
  const normalized =
    error || { name: "Error", message: "The compatibility subframe request failed" };
  if (
    normalized.name === "OfficeHostIdentityMismatchError" &&
    isOfficeHostIdentity(normalized.expected) &&
    isOfficeHostIdentity(normalized.actual)
  ) {
    return new OfficeHostIdentityMismatchError(
      normalized.expected,
      normalized.actual,
    );
  }
  if (
    normalized.name === "OfficeHostIsolationError" &&
    typeof normalized.origin === "string" &&
    typeof normalized.existingSessionId === "string" &&
    typeof normalized.requestedSessionId === "string"
  ) {
    return new OfficeHostIsolationError(
      normalized.origin,
      normalized.existingSessionId,
      normalized.requestedSessionId,
    );
  }
  if (
    normalized.name === "OfficeEditorStartupError" &&
    typeof normalized.phase === "string"
  ) {
    return new OfficeEditorStartupError(
      normalized.phase,
      normalized.message,
      normalized.retryable,
    );
  }
  return new OfficeSubframeRpcError(normalized);
}

class CompatSubframeInstance implements OfficeEditorInstance {
  private state: OfficeEditorState;
  private readonly sourceKind: OfficeEditorState["sourceKind"];
  private hostIdentity: OfficeHostIdentity | null = null;
  private readonly returnsToPreview: boolean;

  constructor(
    readonly id: string,
    private readonly mount: CompatSubframeMount,
    initialState: OfficeEditorState,
  ) {
    this.state = { ...initialState, id };
    this.sourceKind = initialState.sourceKind;
    this.returnsToPreview =
      initialState.mode === "preview" || mount.canReturnToPreview;
  }

  updateState(state: OfficeEditorState) {
    this.state = { ...state, id: this.id, sourceKind: this.sourceKind };
  }

  setHostIdentity(identity: OfficeHostIdentity) {
    this.hostIdentity = { ...identity };
  }

  invokePlugin(pluginGuid: string, payload: unknown): Promise<unknown> {
    if (!pluginGuid) return Promise.reject(new Error("pluginGuid is required"));
    return this.mount.request("invoke-plugin", { pluginGuid, payload });
  }

  save(targetExt?: string): Promise<File> {
    return this.mount.request("save", { targetExt }) as Promise<File>;
  }

  saveAs(targetExt?: string): Promise<File> {
    return this.mount.request("save-as", { targetExt }) as Promise<File>;
  }

  download(targetExt?: string): Promise<File> {
    return this.mount.request("download", { targetExt }) as Promise<File>;
  }

  exportCopy(targetExt?: string): Promise<File> {
    return this.mount.request("save-as", { targetExt }) as Promise<File>;
  }

  print(): Promise<File> {
    return this.mount.print();
  }

  confirmSaveToNewFormat(
    options?: OfficeSaveToNewFormatConfirmationOptions,
  ): Promise<boolean> {
    return this.mount.request("confirm-save-to-new-format", options) as Promise<boolean>;
  }

  setInterfaceTheme(theme: OfficeInterfaceTheme): void {
    void this.mount.request("set-theme", { theme }).catch((error) =>
      this.mount.reportError(toError(error), this),
    );
  }

  setReadonly(readonly: boolean): void {
    const current = this.getState();
    this.updateState({
      ...current,
      readonly,
      mode: readonly
        ? this.returnsToPreview
          ? "preview"
          : "readonly"
        : "edit",
    });
    void this.mount.request("set-readonly", { readonly }).catch((error) =>
      this.mount.reportError(toError(error), this),
    );
  }

  setLanguage(lang: string): Promise<void> {
    if (!lang.trim()) return Promise.reject(new Error("lang is required"));
    return this.mount.request("set-language", { lang }).then(() => undefined);
  }

  destroy(): Promise<void> {
    return this.mount.destroy();
  }

  getState(): OfficeEditorState {
    return { ...this.state };
  }

  getHostIdentity(): OfficeHostIdentity {
    if (!this.hostIdentity) {
      throw new Error("Office host identity is unavailable before activation");
    }
    return { ...this.hostIdentity };
  }
}

class CompatSubframeMount implements OfficeEditorMount {
  readonly id: string;
  private readonly sessionToken: string;
  private readonly targetOrigin: string;
  private readonly frame: HTMLIFrameElement;
  private readonly openPayloadBase: Omit<CompatSubframeOpenPayload, "document">;
  private readonly inputDocument: DescribedDocument;
  private readonly ownerWindow: Window;
  private readonly requestTimeoutMs: number;
  private readonly startupTimeoutMs: number;
  private readonly callbackTimeoutMs: number;
  private phase: OfficeEditorMountState["phase"] = "host-loading";
  private mountError?: Error;
  private childReady = false;
  private destroying = false;
  private destroyed = false;
  private destroyPromise: Promise<void> | null = null;
  private cleanupPromise: Promise<void> | null = null;
  private activatePromise: Promise<OfficeEditorInstance> | null = null;
  private instance: CompatSubframeInstance;
  private pending = new Map<string, PendingRequest>();
  private lastReportedError: { key: string; at: number } | null = null;
  private readonly onMessage: (event: MessageEvent) => void;
  private readonly onFrameError: () => void;

  constructor(
    private readonly container: HTMLElement,
    private readonly options: CreateOfficeEditorOptions,
  ) {
    const ownerWindow = container.ownerDocument.defaultView;
    if (!ownerWindow) {
      throw new OfficeSubframeConfigurationError(
        "The editor container must belong to a live browser window",
      );
    }
    this.ownerWindow = ownerWindow;
    this.id = createOpaqueId(`office-${nextInstanceId++}`, ownerWindow);
    this.sessionToken = createOpaqueId("session", ownerWindow);
    const described = describeDocument(options, ownerWindow);
    const mode = options.mode || (options.readonly ? "readonly" : "edit");
    const context: OfficeHostUrlContext = {
      sessionId: this.id,
      fileName: described.fileName,
      fileType: described.fileType,
      mode,
    };
    const { hostUrl, resourceOrigin, subframeUrl } = resolveSubframeUrl(
      options,
      context,
      this.id,
      this.sessionToken,
      ownerWindow,
    );
    this.targetOrigin = subframeUrl.origin;
    const existingOriginOwner = activeOrigins.get(this.targetOrigin);
    if (existingOriginOwner) {
      throw new OfficeHostIsolationError(
        this.targetOrigin,
        existingOriginOwner.id,
        this.id,
      );
    }
    this.requestTimeoutMs = positiveTimeout(
      options.requestTimeoutMs,
      DEFAULT_REQUEST_TIMEOUT_MS,
    );
    this.startupTimeoutMs = positiveTimeout(
      options.startupTimeoutMs,
      DEFAULT_STARTUP_TIMEOUT_MS,
    );
    this.callbackTimeoutMs = positiveTimeout(
      options.callbackTimeoutMs,
      DEFAULT_CALLBACK_TIMEOUT_MS,
    );
    this.inputDocument = described.document;
    this.openPayloadBase = {
      hostUrl: hostUrl.toString(),
      resourceOrigin: resourceOrigin.origin,
      expectedHostIdentity: options.expectedHostIdentity,
      mode,
      readonly: mode !== "edit",
      canReturnToPreview: options.canReturnToPreview,
      spellcheck: options.spellcheck,
      interfaceTheme: options.interfaceTheme,
      lang: options.lang,
      plugins: options.plugins,
      saveBehavior: options.saveBehavior,
      callbacks: {
        onSave: typeof options.onSave === "function",
        onSaveAs: typeof options.onSaveAs === "function",
        onDownload: typeof options.onDownload === "function",
      },
    };

    const initialState: OfficeEditorState = {
      id: this.id,
      fileName: described.fileName,
      fileType: described.fileType,
      mode,
      readonly: mode !== "edit",
      dirty: false,
      sourceKind:
        described.document.kind === "file"
          ? "local-file"
          : described.document.kind === "empty"
            ? "new-document"
            : described.document.kind === "buffer"
              ? described.document.sourceKind || "buffer"
              : "url",
      status: "opening",
      destroyed: false,
    };
    this.instance = new CompatSubframeInstance(this.id, this, initialState);
    this.frame = container.ownerDocument.createElement("iframe");
    this.frame.name = this.id;
    this.frame.title = options.frameTitle || "ONLYOFFICE editor";
    this.frame.className = options.frameClassName || "onlyoffice-compat-subframe";
    this.frame.allow = "clipboard-read; clipboard-write; fullscreen";
    this.frame.referrerPolicy = "strict-origin";
    this.frame.style.width = "100%";
    this.frame.style.height = "100%";
    this.frame.style.minWidth = "0";
    this.frame.style.minHeight = "0";
    this.frame.style.display = "block";
    this.frame.style.border = "0";
    this.frame.src = subframeUrl.toString();
    this.onMessage = (event) => this.handleMessage(event);
    this.onFrameError = () => {
      const error = new Error(
        `Failed to load compatibility subframe at ${subframeUrl.origin}`,
      );
      this.fail(error);
      void this.cleanupFailedActivation();
    };
    ownerWindow.addEventListener("message", this.onMessage);
    this.frame.addEventListener("error", this.onFrameError, { once: true });
    container.replaceChildren(this.frame);
    activeOrigins.set(this.targetOrigin, this);
  }

  get canReturnToPreview() {
    return this.options.canReturnToPreview === true;
  }

  async print(): Promise<File> {
    const printWindow = openOfficePrintWindow(
      this.ownerWindow,
      this.container.ownerDocument,
    );
    try {
      const file = (await this.request("print")) as File;
      await printOfficePdfFile({
        ownerWindow: this.ownerWindow,
        ownerDocument: this.container.ownerDocument,
        file,
        printWindow,
      });
      return file;
    } catch (error) {
      printWindow?.close();
      throw error;
    }
  }

  activate(): Promise<OfficeEditorInstance> {
    if (this.activatePromise) return this.activatePromise;
    if (this.destroyed) {
      return Promise.reject(new DOMException("Editor was destroyed", "AbortError"));
    }
    this.phase = "waiting-for-activation";
    this.activatePromise = this.prepareOpenPayload()
      .then((payload) => this.request("open", payload, this.startupTimeoutMs))
      .then((value) => {
        const result = value as Partial<CompatSubframeOpenResult>;
        if (!isOfficeEditorState(result?.state) || !isOfficeHostIdentity(result?.hostIdentity)) {
          throw new Error("Compatibility subframe returned an invalid activation result");
        }
        this.instance.updateState(result.state);
        this.instance.setHostIdentity(result.hostIdentity);
        this.phase = "ready";
        this.notifyObserver(() => this.options.onReady?.(this.instance));
        return this.instance;
      })
      .catch((error) => {
        const normalized = toError(error);
        this.fail(normalized);
        return this.cleanupFailedActivation().then(() => {
          throw normalized;
        });
      });
    return this.activatePromise;
  }

  private async prepareOpenPayload(): Promise<CompatSubframeOpenPayload> {
    if (this.inputDocument.kind !== "url") {
      return { ...this.openPayloadBase, document: this.inputDocument };
    }
    const { url, fileName, fetchOptions } = this.inputDocument;
    const response = await this.withTimeout(
      this.ownerWindow.fetch(url, fetchOptions),
      this.startupTimeoutMs,
      "Office document fetch timed out",
    );
    if (!response.ok) {
      throw new Error(`Failed to fetch Office document (${response.status} ${response.statusText})`);
    }
    const buffer = await response.arrayBuffer();
    return {
      ...this.openPayloadBase,
      document: { kind: "buffer", buffer, fileName, sourceKind: "url" },
    };
  }

  request(
    action: CompatSubframeAction,
    payload?: unknown,
    timeoutMs = this.requestTimeoutMs,
  ): Promise<unknown> {
    if ((this.destroyed || this.destroying) && action !== "destroy") {
      return Promise.reject(new DOMException("Editor was destroyed", "AbortError"));
    }
    const requestId = createOpaqueId("request", this.ownerWindow);
    const message: CompatSubframeRequest = {
      source: COMPAT_SUBFRAME_PROTOCOL_SOURCE,
      version: COMPAT_SUBFRAME_PROTOCOL_VERSION,
      type: "request",
      instanceId: this.id,
      sessionToken: this.sessionToken,
      requestId,
      action,
      payload,
    };
    return new Promise((resolve, reject) => {
      const timer = this.ownerWindow.setTimeout(() => {
        const pending = this.pending.get(requestId);
        if (!pending) return;
        this.pending.delete(requestId);
        const timeoutError = new Error(
          `Compatibility subframe ${action} request timed out`,
        );
        reject(timeoutError);
        if (action === "open") {
          this.fail(timeoutError);
        } else if (
          action === "save" ||
          action === "save-as" ||
          action === "download" ||
          action === "invoke-plugin"
        ) {
          this.fail(timeoutError);
          void this.cleanupFailedActivation();
        }
      }, timeoutMs);
      this.pending.set(requestId, {
        action,
        message,
        resolve,
        reject,
        timer,
        sent: false,
        transfer: this.getTransferList(message),
      });
      if (this.childReady) this.sendPending(requestId);
    });
  }

  destroy(): Promise<void> {
    if (this.destroyPromise) return this.destroyPromise;
    if (this.cleanupPromise) return this.cleanupPromise;
    if (this.destroyed) return Promise.resolve();
    this.destroyed = true;
    this.destroying = true;
    this.phase = "destroyed";
    this.instance.updateState({
      ...this.instance.getState(),
      dirty: false,
      status: "destroyed",
      destroyed: true,
    });
    const abortError = new DOMException("Editor was destroyed", "AbortError");
    this.rejectPending(abortError);
    this.destroyPromise = (async () => {
      try {
        if (this.childReady) await this.requestDestroy();
      } finally {
        this.destroying = false;
        this.dispose(abortError);
      }
    })();
    return this.destroyPromise;
  }

  private requestDestroy() {
    const requestId = createOpaqueId("request", this.ownerWindow);
    const message: CompatSubframeRequest = {
      source: COMPAT_SUBFRAME_PROTOCOL_SOURCE,
      version: COMPAT_SUBFRAME_PROTOCOL_VERSION,
      type: "request",
      instanceId: this.id,
      sessionToken: this.sessionToken,
      requestId,
      action: "destroy",
    };
    return new Promise<void>((resolve) => {
      const timer = this.ownerWindow.setTimeout(() => {
        this.pending.delete(requestId);
        resolve();
      }, positiveTimeout(this.options.destroyTimeoutMs, DEFAULT_DESTROY_TIMEOUT_MS));
      this.pending.set(requestId, {
        action: "destroy",
        message,
        resolve: () => resolve(),
        reject: () => resolve(),
        timer,
        sent: false,
        transfer: [],
      });
      this.sendPending(requestId);
    });
  }

  getState(): OfficeEditorMountState {
    return {
      id: this.id,
      origin: this.targetOrigin,
      phase: this.phase,
      ...(this.mountError ? { error: this.mountError } : {}),
    };
  }

  reportError(error: Error, instance: OfficeEditorInstance = this.instance) {
    const now = Date.now();
    const key = `${error.name}:${error.message}`;
    if (
      this.lastReportedError?.key === key &&
      now - this.lastReportedError.at < 1_000
    ) {
      return;
    }
    this.lastReportedError = { key, at: now };
    try {
      this.options.onError?.(error, instance);
    } catch {
      // Error observers cannot compromise the transport lifecycle.
    }
  }

  private handleMessage(event: MessageEvent) {
    if (
      event.origin !== this.targetOrigin ||
      event.source !== this.frame.contentWindow ||
      !isCompatSubframeChildMessage(event.data)
    ) {
      return;
    }
    const message = event.data;
    if (
      message.instanceId !== this.id ||
      message.sessionToken !== this.sessionToken
    ) {
      return;
    }

    if (this.destroying || this.destroyed) {
      if (
        message.type === "response" &&
        this.pending.get(message.requestId)?.action === "destroy"
      ) {
        this.handleResponse(message);
      }
      return;
    }

    if (message.type === "ready") {
      if (this.childReady) return;
      this.childReady = true;
      this.phase = "runtime-loading";
      for (const requestId of this.pending.keys()) this.sendPending(requestId);
      return;
    }
    if (message.type === "response") {
      this.handleResponse(message);
      return;
    }
    if (message.type === "callback-request") {
      void this.handleCallback(message);
      return;
    }
    this.handleEvent(message);
  }

  private sendPending(requestId: string) {
    const pending = this.pending.get(requestId);
    if (!pending || pending.sent || !this.frame.contentWindow) return;
    try {
      this.frame.contentWindow.postMessage(pending.message, {
        targetOrigin: this.targetOrigin,
        transfer: pending.transfer,
      });
      pending.sent = true;
    } catch (error) {
      this.pending.delete(requestId);
      this.ownerWindow.clearTimeout(pending.timer);
      pending.reject(toError(error));
    }
  }

  private send(message: CompatSubframeParentMessage) {
    const target = this.frame.contentWindow;
    if (!target || this.destroyed) return false;
    try {
      target.postMessage(message, this.targetOrigin);
      return true;
    } catch {
      return false;
    }
  }

  private handleResponse(message: CompatSubframeResponse) {
    const pending = this.pending.get(message.requestId);
    if (!pending) return;
    this.pending.delete(message.requestId);
    this.ownerWindow.clearTimeout(pending.timer);
    if (!message.ok) {
      pending.reject(remoteError(message.error));
      return;
    }
    if (!isValidActionResult(pending.action, message.result)) {
      pending.reject(
        new OfficeSubframeRpcError({
          name: "ProtocolError",
          message: `Compatibility subframe returned an invalid ${pending.action} result`,
        }),
      );
      return;
    }
    pending.resolve(message.result);
  }

  private async handleCallback(message: CompatSubframeCallbackRequest) {
    const callback =
      message.callback === "save"
        ? this.options.onSave
        : message.callback === "save-as"
          ? this.options.onSaveAs
          : this.options.onDownload;
    let response: CompatSubframeCallbackResponse;
    try {
      if (!callback) throw new Error(`No ${message.callback} callback is registered`);
      const result = await this.withTimeout(
        Promise.resolve(callback(message.payload.file, this.instance)),
        this.callbackTimeoutMs,
        `${message.callback} callback timed out`,
      );
      response = {
        ...this.envelope(),
        type: "callback-response",
        callbackRequestId: message.callbackRequestId,
        ok: true,
        result:
          message.callback === "download"
            ? undefined
            : typeof result === "boolean"
              ? result
              : undefined,
      };
    } catch (error) {
      const normalized = toError(error);
      response = {
        ...this.envelope(),
        type: "callback-response",
        callbackRequestId: message.callbackRequestId,
        ok: false,
        error: serializeCompatSubframeError(normalized),
      };
    }
    if (!this.send(response) && response.ok) {
      this.send({
        ...this.envelope(),
        type: "callback-response",
        callbackRequestId: message.callbackRequestId,
        ok: false,
        error: serializeCompatSubframeError(
          new DOMException("Callback result is not cloneable", "DataCloneError"),
        ),
      });
    }
  }

  private handleEvent(message: CompatSubframeEventMessage) {
    if (message.event === "plugin-ready") {
      const payload = message.payload as { pluginGuid?: unknown; editorType?: unknown };
      if (typeof payload?.pluginGuid === "string" && typeof payload.editorType === "string") {
        this.notifyObserver(() =>
          this.options.onPluginReady?.(payload.pluginGuid as string, payload.editorType as string, this.instance),
        );
      }
      return;
    }
    if (message.event === "dirty-change") {
      const dirty =
        typeof message.payload === "boolean"
          ? message.payload
          : (message.payload as { dirty?: unknown })?.dirty;
      if (typeof dirty === "boolean") {
        const state = { ...this.instance.getState(), dirty };
        this.instance.updateState(state);
        this.notifyObserver(() => this.options.onDirtyChange?.(dirty, this.instance));
      }
      return;
    }
    if (message.event === "state-change" || message.event === "document-ready") {
      const payload =
        message.event === "document-ready" &&
        message.payload &&
        typeof message.payload === "object" &&
        "state" in message.payload
          ? (message.payload as { state: unknown }).state
          : message.payload;
      if (isOfficeEditorState(payload)) {
        this.instance.updateState(payload);
        if (message.event === "state-change") {
          this.notifyObserver(() =>
            this.options.onStateChange?.(this.instance.getState(), this.instance),
          );
        }
      }
      return;
    }
    if (message.event === "error") {
      if (isCompatSubframeSerializedError(message.payload)) {
        this.reportError(remoteError(message.payload));
      }
    }
  }

  private envelope() {
    return {
      source: COMPAT_SUBFRAME_PROTOCOL_SOURCE,
      version: COMPAT_SUBFRAME_PROTOCOL_VERSION,
      instanceId: this.id,
      sessionToken: this.sessionToken,
    } as const;
  }

  private withTimeout<T>(promise: Promise<T>, timeoutMs: number, message: string) {
    return new Promise<T>((resolve, reject) => {
      const timer = this.ownerWindow.setTimeout(() => reject(new Error(message)), timeoutMs);
      void promise.then(
        (value) => {
          this.ownerWindow.clearTimeout(timer);
          resolve(value);
        },
        (error) => {
          this.ownerWindow.clearTimeout(timer);
          reject(error);
        },
      );
    });
  }

  private fail(error: Error) {
    if (this.destroyed) return;
    this.mountError = error;
    this.phase = "error";
    this.reportError(error);
  }

  private async cleanupFailedActivation() {
    if (this.cleanupPromise) return this.cleanupPromise;
    if (this.destroyPromise) return this.destroyPromise;
    if (this.destroyed) return;
    this.destroyed = true;
    this.destroying = true;
    this.instance.updateState({
      ...this.instance.getState(),
      status: "error",
      destroyed: true,
    });
    const abortError = new DOMException("Editor activation failed", "AbortError");
    this.rejectPending(abortError);
    this.cleanupPromise = (async () => {
      try {
        if (this.childReady) await this.requestDestroy();
      } finally {
        this.destroying = false;
        this.dispose(abortError);
      }
    })();
    return this.cleanupPromise;
  }

  private rejectPending(error: Error) {
    this.pending.forEach((pending) => {
      this.ownerWindow.clearTimeout(pending.timer);
      pending.reject(error);
    });
    this.pending.clear();
  }

  private notifyObserver(callback: () => unknown) {
    try {
      void Promise.resolve(callback()).catch((error) =>
        this.reportError(toError(error)),
      );
    } catch (error) {
      this.reportError(toError(error));
    }
  }

  private getTransferList(message: CompatSubframeRequest): Transferable[] {
    if (message.action !== "open") return [];
    const payload = message.payload as CompatSubframeOpenPayload;
    if (payload?.document?.kind !== "buffer") return [];
    const input = payload.document.buffer;
    if (Object.prototype.toString.call(input) === "[object ArrayBuffer]") {
      const copy = (input as ArrayBuffer).slice(0);
      payload.document = { ...payload.document, buffer: copy };
      return [copy];
    }
    if (ArrayBuffer.isView(input)) {
      const view = input as Uint8Array;
      const copy = view.slice().buffer;
      payload.document = { ...payload.document, buffer: copy };
      return [copy];
    }
    return [];
  }

  private dispose(error: Error) {
    this.ownerWindow.removeEventListener("message", this.onMessage);
    this.frame.removeEventListener("error", this.onFrameError);
    this.pending.forEach((pending) => {
      this.ownerWindow.clearTimeout(pending.timer);
      pending.reject(error);
    });
    this.pending.clear();
    if (this.frame.parentNode === this.container) this.frame.remove();
    if (activeMounts.get(this.container) === this) activeMounts.delete(this.container);
    if (activeOrigins.get(this.targetOrigin) === this) {
      activeOrigins.delete(this.targetOrigin);
    }
    activeInstances.delete(this);
  }
}

export function mountOfficeEditor(
  container: HTMLElement,
  options: CreateOfficeEditorOptions,
): OfficeEditorMount {
  if (
    !container ||
    typeof container !== "object" ||
    container.nodeType !== 1 ||
    !container.ownerDocument
  ) {
    throw new OfficeSubframeConfigurationError(
      "mountOfficeEditor requires an HTMLElement container",
    );
  }
  const existing = activeMounts.get(container);
  if (existing) {
    throw new OfficeSubframeConfigurationError(
      `The editor container is already owned by ${existing.id}`,
    );
  }
  try {
    const mount = new CompatSubframeMount(container, options);
    activeMounts.set(container, mount);
    activeInstances.add(mount);
    return mount;
  } catch (error) {
    const normalized = toError(error);
    try {
      options.onError?.(normalized);
    } catch {
      // Preserve the original configuration error.
    }
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
    throw error;
  }
}

export const mountOfficeSubframeEditor = mountOfficeEditor;
export const createOfficeSubframeEditor = createOfficeEditor;

export function getActiveOfficeEditorCount() {
  return activeInstances.size;
}

export const getActiveOfficeSubframeEditorCount = getActiveOfficeEditorCount;

/** The isolated child owns runtime initialization; parent setup is immediate. */
export function loadOfficeEditorApi(): Promise<void> {
  return Promise.resolve();
}
