import type {
  FileType,
  OfficeTheme,
  OfficeXmlEventConfig,
} from "@/components/onlyoffice-web-comp";

/**
 * Parent/subframe communication is deliberately kept small and explicit.
 * The editor runtime never crosses the subframe boundary; only serializable
 * commands and results do.
 */
export const SUBFRAME_MESSAGE_SOURCE = "onlyoffice-subframe" as const;
export const SUBFRAME_EDITOR_CONTAINER_ID = "onlyoffice-subframe-editor";

export type SubframeDocumentInput = {
  fileName: string;
  file?: File;
  isNew?: boolean;
  readOnly?: boolean;
};

export type SubframeOpenPayload = {
  containerId: string;
  fileType: FileType;
  defaultFileName: string;
  readOnly: boolean;
  theme: OfficeTheme;
  officeXmlEvent?: OfficeXmlEventConfig;
  document: SubframeDocumentInput;
};

export type SubframeEditorAction =
  | "get-all-comments"
  | "add-comment"
  | "update-comment"
  | "remove-comment"
  | "get-all-revisions"
  | "set-track-revisions"
  | "add-demo-revision"
  | "go-to-revision"
  | "accept-revision"
  | "reject-revision"
  | "accept-all-revisions"
  | "reject-all-revisions";

export type SubframeConnectorOperation =
  | "connect"
  | "disconnect"
  | "call-command"
  | "execute-method";

export type SubframeRequestAction =
  | "open"
  | "set-read-only"
  | "set-theme"
  | "download"
  | "print-logs"
  | "editor"
  | "connector"
  | "destroy";

export type SubframeRequest = {
  source: typeof SUBFRAME_MESSAGE_SOURCE;
  type: "request";
  instanceId: string;
  requestId: string;
  action: SubframeRequestAction;
  payload?: unknown;
};

export type SubframeReadyMessage = {
  source: typeof SUBFRAME_MESSAGE_SOURCE;
  type: "ready";
  instanceId: string;
};

export type SubframeResponse = {
  source: typeof SUBFRAME_MESSAGE_SOURCE;
  type: "response";
  instanceId: string;
  requestId: string;
  ok: boolean;
  result?: unknown;
  error?: string;
};

export type SubframeEventName =
  | "loading-change"
  | "document-ready"
  | "office-xml-size-limit-exceeded";

export type SubframeEventMessage = {
  source: typeof SUBFRAME_MESSAGE_SOURCE;
  type: "event";
  instanceId: string;
  event: SubframeEventName;
  payload?: unknown;
};

export type SubframeMessage =
  | SubframeReadyMessage
  | SubframeResponse
  | SubframeEventMessage;
