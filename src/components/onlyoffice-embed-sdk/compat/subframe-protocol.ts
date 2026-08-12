import type {
  OfficeEditorMode,
  OfficeEditorState,
  OfficeHostIdentity,
  OfficeInterfaceTheme,
  OfficeLoadingState,
  OfficePluginOptions,
  OfficeSaveBehavior,
  OfficeSaveToNewFormatConfirmationOptions,
} from "./editor";

/** Stable, versioned discriminator for the public compatibility transport. */
export const COMPAT_SUBFRAME_PROTOCOL_SOURCE =
  "onlyoffice-embed-sdk/compat-subframe/v1" as const;
export const COMPAT_SUBFRAME_PROTOCOL_VERSION = 1 as const;
export const COMPAT_SUBFRAME_RUNTIME = "compat" as const;

export type CompatSubframeSerializedError = {
  name: string;
  message: string;
  code?: string;
  stack?: string;
  phase?: string;
  retryable?: boolean;
  expected?: OfficeHostIdentity;
  actual?: OfficeHostIdentity;
  origin?: string;
  existingSessionId?: string;
  requestedSessionId?: string;
};

export type CompatSubframeDocument =
  | { kind: "file"; file: File | Blob; fileName: string }
  | {
      kind: "buffer";
      buffer: Blob | ArrayBuffer | Uint8Array;
      fileName: string;
      sourceKind?: "buffer" | "url";
    }
  | { kind: "empty"; emptyType: "docx" | "xlsx" | "pptx" | "csv"; fileName: string };

export type CompatSubframeOpenPayload = {
  hostUrl: string;
  /** Canonical, shared static-resource origin; never the embedding parent origin. */
  resourceOrigin: string;
  expectedHostIdentity?: OfficeHostIdentity;
  document: CompatSubframeDocument;
  mode: OfficeEditorMode;
  readonly: boolean;
  canReturnToPreview?: boolean;
  spellcheck?: boolean;
  interfaceTheme?: OfficeInterfaceTheme;
  lang?: string;
  plugins?: OfficePluginOptions;
  saveBehavior?: OfficeSaveBehavior;
  callbacks: {
    onSave: boolean;
    onSaveAs: boolean;
    onDownload: boolean;
  };
};

export type CompatSubframeAction =
  | "open"
  | "set-readonly"
  | "set-theme"
  | "set-language"
  | "invoke-plugin"
  | "save"
  | "save-as"
  | "download"
  | "print"
  | "confirm-save-to-new-format"
  | "get-state"
  | "get-host-identity"
  | "destroy";

type CompatSubframeEnvelope = {
  source: typeof COMPAT_SUBFRAME_PROTOCOL_SOURCE;
  version: typeof COMPAT_SUBFRAME_PROTOCOL_VERSION;
  instanceId: string;
  sessionToken: string;
};

export type CompatSubframeRequest = CompatSubframeEnvelope & {
  type: "request";
  requestId: string;
  action: CompatSubframeAction;
  payload?: unknown;
};

export type CompatSubframeReadyMessage = CompatSubframeEnvelope & {
  type: "ready";
};

export type CompatSubframeResponse = CompatSubframeEnvelope & {
  type: "response";
  requestId: string;
  ok: boolean;
  result?: unknown;
  error?: CompatSubframeSerializedError;
};

export type CompatSubframeEventName =
  | "loading-change"
  | "document-ready"
  | "plugin-ready"
  | "dirty-change"
  | "state-change"
  | "error";

export type CompatSubframeEventMessage = CompatSubframeEnvelope & {
  type: "event";
  event: CompatSubframeEventName;
  payload?: unknown;
};

function isLoadingState(value: unknown): value is OfficeLoadingState {
  if (!isRecord(value)) return false;
  return (
    typeof value.loading === "boolean" &&
    [
      "host-loading",
      "runtime-loading",
      "static-resources",
      "operation",
      "ready",
      "error",
      "destroyed",
    ].includes(String(value.phase)) &&
    ["checking", "cache-hit", "downloading", "downloaded", "not-observed"].includes(
      String(value.resourceStatus),
    ) &&
    typeof value.resourceDownload === "boolean" &&
    typeof value.transferredBytes === "number" &&
    Number.isFinite(value.transferredBytes) &&
    value.transferredBytes >= 0 &&
    typeof value.resourceCount === "number" &&
    Number.isInteger(value.resourceCount) &&
    value.resourceCount >= 0
  );
}

export type CompatSubframeCallbackName = "save" | "save-as" | "download";

export type CompatSubframeCallbackRequest = CompatSubframeEnvelope & {
  type: "callback-request";
  callbackRequestId: string;
  callback: CompatSubframeCallbackName;
  payload: { file: File; fileName: string };
};

export type CompatSubframeCallbackResponse = CompatSubframeEnvelope & {
  type: "callback-response";
  callbackRequestId: string;
  ok: boolean;
  result?: unknown;
  error?: CompatSubframeSerializedError;
};

export type CompatSubframeParentMessage =
  | CompatSubframeRequest
  | CompatSubframeCallbackResponse;

export type CompatSubframeChildMessage =
  | CompatSubframeReadyMessage
  | CompatSubframeResponse
  | CompatSubframeEventMessage
  | CompatSubframeCallbackRequest;

export type CompatSubframeMessage =
  | CompatSubframeParentMessage
  | CompatSubframeChildMessage;

export type CompatSubframeOpenResult = {
  state: OfficeEditorState;
  hostIdentity: OfficeHostIdentity;
};

export type CompatSubframePluginInvocation = {
  pluginGuid: string;
  payload: unknown;
};

export type CompatSubframeSaveRequest = { targetExt?: string };
export type CompatSubframeReadonlyRequest = { readonly: boolean };
export type CompatSubframeThemeRequest = { theme: OfficeInterfaceTheme };
export type CompatSubframeConfirmSaveRequest =
  OfficeSaveToNewFormatConfirmationOptions | undefined;

const ACTIONS = new Set<string>([
  "open",
  "set-readonly",
  "set-theme",
  "set-language",
  "invoke-plugin",
  "save",
  "save-as",
  "download",
  "print",
  "confirm-save-to-new-format",
  "get-state",
  "get-host-identity",
  "destroy",
]);

const EVENTS = new Set<string>([
  "loading-change",
  "document-ready",
  "plugin-ready",
  "dirty-change",
  "state-change",
  "error",
]);

const CALLBACKS = new Set<string>(["save", "save-as", "download"]);

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object";
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === "string" && value.length > 0;
}

function isBlobLike(value: unknown) {
  return (
    isRecord(value) &&
    typeof value.size === "number" &&
    typeof value.type === "string" &&
    typeof value.arrayBuffer === "function"
  );
}

function isIdentity(value: unknown): value is OfficeHostIdentity {
  return (
    isRecord(value) &&
    isNonEmptyString(value.packageVersion) &&
    isNonEmptyString(value.hostBuildId) &&
    typeof value.assetManifestDigest === "string" &&
    /^[a-f0-9]{64}$/i.test(value.assetManifestDigest)
  );
}

export function isCompatSubframeSerializedError(
  value: unknown,
): value is CompatSubframeSerializedError {
  if (
    !isRecord(value) ||
    !isNonEmptyString(value.name) ||
    !isNonEmptyString(value.message)
  ) {
    return false;
  }
  return (
    (value.code === undefined || typeof value.code === "string") &&
    (value.stack === undefined || typeof value.stack === "string") &&
    (value.phase === undefined || typeof value.phase === "string") &&
    (value.retryable === undefined || typeof value.retryable === "boolean") &&
    (value.expected === undefined || isIdentity(value.expected)) &&
    (value.actual === undefined || isIdentity(value.actual)) &&
    (value.origin === undefined || typeof value.origin === "string") &&
    (value.existingSessionId === undefined ||
      typeof value.existingSessionId === "string") &&
    (value.requestedSessionId === undefined ||
      typeof value.requestedSessionId === "string")
  );
}

function isEnvelope(
  value: unknown,
): value is CompatSubframeEnvelope & Record<string, unknown> {
  if (!isRecord(value)) return false;
  return (
    value.source === COMPAT_SUBFRAME_PROTOCOL_SOURCE &&
    value.version === COMPAT_SUBFRAME_PROTOCOL_VERSION &&
    isNonEmptyString(value.instanceId) &&
    isNonEmptyString(value.sessionToken)
  );
}

export function isCompatSubframeRequest(
  value: unknown,
): value is CompatSubframeRequest {
  if (!isEnvelope(value) || value.type !== "request") return false;
  return (
    isNonEmptyString(value.requestId) &&
    typeof value.action === "string" &&
    ACTIONS.has(value.action)
  );
}

export function isCompatSubframeCallbackResponse(
  value: unknown,
): value is CompatSubframeCallbackResponse {
  return (
    isEnvelope(value) &&
    value.type === "callback-response" &&
    isNonEmptyString(value.callbackRequestId) &&
    typeof value.ok === "boolean" &&
    (value.ok || isCompatSubframeSerializedError(value.error))
  );
}

export function isCompatSubframeParentMessage(
  value: unknown,
): value is CompatSubframeParentMessage {
  return (
    isCompatSubframeRequest(value) || isCompatSubframeCallbackResponse(value)
  );
}

export function isCompatSubframeChildMessage(
  value: unknown,
): value is CompatSubframeChildMessage {
  if (!isEnvelope(value) || typeof value.type !== "string") return false;
  if (value.type === "ready") return true;
  if (value.type === "response") {
    return (
      isNonEmptyString(value.requestId) &&
      typeof value.ok === "boolean" &&
      (value.ok || isCompatSubframeSerializedError(value.error))
    );
  }
  if (value.type === "event") {
    if (typeof value.event !== "string" || !EVENTS.has(value.event)) return false;
    if (value.event === "loading-change") {
      return isLoadingState(value.payload);
    }
    if (value.event === "plugin-ready") {
      return (
        isRecord(value.payload) &&
        isNonEmptyString(value.payload.pluginGuid) &&
        isNonEmptyString(value.payload.editorType)
      );
    }
    if (value.event === "dirty-change") {
      return isRecord(value.payload) && typeof value.payload.dirty === "boolean";
    }
    if (value.event === "error") {
      return isCompatSubframeSerializedError(value.payload);
    }
    return isRecord(value.payload);
  }
  if (value.type === "callback-request") {
    return (
      isNonEmptyString(value.callbackRequestId) &&
      typeof value.callback === "string" &&
      CALLBACKS.has(value.callback) &&
      isRecord(value.payload) &&
      isBlobLike(value.payload.file) &&
      isNonEmptyString(value.payload.fileName)
    );
  }
  return false;
}

export function serializeCompatSubframeError(
  error: unknown,
): CompatSubframeSerializedError {
  if (error instanceof Error) {
    const details = error as Error & Record<string, unknown>;
    const code =
      "code" in error && typeof error.code === "string" ? error.code : undefined;
    return {
      name: error.name || "Error",
      message: error.message,
      ...(code ? { code } : {}),
      ...(error.stack ? { stack: error.stack } : {}),
      ...(typeof details.phase === "string" ? { phase: details.phase } : {}),
      ...(typeof details.retryable === "boolean"
        ? { retryable: details.retryable }
        : {}),
      ...(isRecord(details.expected)
        ? { expected: details.expected as unknown as OfficeHostIdentity }
        : {}),
      ...(isRecord(details.actual)
        ? { actual: details.actual as unknown as OfficeHostIdentity }
        : {}),
      ...(typeof details.origin === "string" ? { origin: details.origin } : {}),
      ...(typeof details.existingSessionId === "string"
        ? { existingSessionId: details.existingSessionId }
        : {}),
      ...(typeof details.requestedSessionId === "string"
        ? { requestedSessionId: details.requestedSessionId }
        : {}),
    };
  }
  return { name: "Error", message: String(error) };
}
