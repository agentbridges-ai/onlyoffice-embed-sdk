import type {
  CommentData,
  CommentItem,
  RevisionItem,
} from "@/components/onlyoffice-web-comp";
import type {
  FileType,
  OfficeTheme,
  OfficeXmlEventConfig,
} from "@/components/onlyoffice-web-comp";
import type {
  OnlyOfficeConnector,
  OnlyOfficeConnectorOptions,
} from "@/components/onlyoffice-web-comp";
import { downloadBlob } from "@/components/onlyoffice-web-comp";
import {
  SUBFRAME_MESSAGE_SOURCE,
  type SubframeConnectorOperation,
  type SubframeDocumentInput,
  type SubframeEditorAction,
  type SubframeEventName,
  type SubframeMessage,
  type SubframeOpenPayload,
  type SubframeRequest,
  type SubframeRequestAction,
} from "./subframe-protocol";

const RPC_TIMEOUT_MS = 30_000;

type PendingRequest = {
  action: SubframeRequestAction;
  resolve: (value: unknown) => void;
  reject: (reason?: unknown) => void;
  timer: number;
};

export type SubframeEventHandler = (
  event: SubframeEventName,
  payload: unknown,
) => void;

/**
 * The demo controls only need this stable editor surface. Both the regular
 * OnlyOffice EditorManager and the RPC-backed editor proxy satisfy it.
 */
export type MultiInstanceEditor = {
  createConnector: (
    options?: OnlyOfficeConnectorOptions,
  ) => OnlyOfficeConnector;
  getAllComments: () => Promise<CommentItem[]>;
  addComment: (input: CommentData | string) => Promise<string> | string;
  updateComment: (id: string, data: CommentData) => Promise<void> | void;
  removeComment: (id: string) => Promise<void> | void;
  getAllRevisions: () => Promise<RevisionItem[]>;
  setTrackRevisions: (enabled: boolean) => Promise<unknown> | void;
  addDemoRevision: (text?: string) => Promise<RevisionItem[]> | RevisionItem[];
  goToRevision: (id: string) => Promise<unknown> | void;
  acceptRevision: (revision: RevisionItem | string) => Promise<unknown> | void;
  rejectRevision: (revision: RevisionItem | string) => Promise<unknown> | void;
  acceptAllRevisions: () => Promise<unknown> | void;
  rejectAllRevisions: () => Promise<unknown> | void;
};

export type MultiInstanceManager = {
  getEditor: () => MultiInstanceEditor;
};

export type SubframeManagerOptions = {
  containerId: string;
  fileType: FileType;
  defaultFileName: string;
  readOnly: boolean;
  theme: OfficeTheme;
  officeXmlEvent?: OfficeXmlEventConfig;
  instanceId: string;
  targetOrigin: string;
  frame?: HTMLIFrameElement | null;
  onEvent?: SubframeEventHandler;
};

type ConfigureOptions = Partial<
  Pick<
    SubframeOpenPayload,
    "defaultFileName" | "readOnly" | "theme" | "officeXmlEvent"
  >
>;

function isSubframeMessage(value: unknown): value is SubframeMessage {
  if (!value || typeof value !== "object") {
    return false;
  }

  return (value as { source?: unknown }).source === SUBFRAME_MESSAGE_SOURCE;
}

function requestId() {
  return `subframe-${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

function getRevisionId(revision: RevisionItem | string) {
  return typeof revision === "string" ? revision : revision.Id;
}

class SubframeConnectorProxy implements OnlyOfficeConnector {
  private connected = false;

  constructor(
    private readonly manager: SubframeManager,
    private readonly options?: OnlyOfficeConnectorOptions,
  ) {}

  get isConnected() {
    return this.connected;
  }

  connect() {
    this.connected = true;
    void this.manager
      .request("connector", {
        operation: "connect" satisfies SubframeConnectorOperation,
        options: this.options,
      })
      .catch(() => {
        this.connected = false;
      });
  }

  disconnect() {
    this.connected = false;
    void this.manager
      .request("connector", {
        operation: "disconnect" satisfies SubframeConnectorOperation,
      })
      .catch(() => {});
  }

  callCommand(
    command: string | (() => void),
    callback?: (result: unknown) => void,
    recalculate?: boolean,
  ) {
    void this.manager
      .request("connector", {
        operation: "call-command" satisfies SubframeConnectorOperation,
        command: typeof command === "function" ? command.toString() : command,
        recalculate,
      })
      .then((result) => callback?.(result))
      .catch(() => callback?.(undefined));
  }

  executeMethod(
    method: string,
    args: unknown[],
    callback?: (result: unknown) => void,
  ) {
    void this.manager
      .request("connector", {
        operation: "execute-method" satisfies SubframeConnectorOperation,
        method,
        args,
      })
      .then((result) => callback?.(result))
      .catch(() => callback?.(undefined));
  }

  attachEvent() {}

  detachEvent() {}

  callCommandAsync(command: string | (() => void)) {
    return this.manager.request("connector", {
      operation: "call-command" satisfies SubframeConnectorOperation,
      command: typeof command === "function" ? command.toString() : command,
    });
  }

  callMethodAsync(method: string, args: unknown[] = []) {
    return this.manager.request("connector", {
      operation: "execute-method" satisfies SubframeConnectorOperation,
      method,
      args,
    });
  }
}

class SubframeEditorProxy implements MultiInstanceEditor {
  constructor(private readonly manager: SubframeManager) {}

  createConnector(options?: OnlyOfficeConnectorOptions) {
    return new SubframeConnectorProxy(this.manager, options);
  }

  private request(action: SubframeEditorAction, payload?: unknown) {
    return this.manager.request("editor", { action, payload });
  }

  getAllComments() {
    return this.request("get-all-comments") as Promise<CommentItem[]>;
  }

  addComment(input: CommentData | string) {
    return this.request("add-comment", { input }) as Promise<string>;
  }

  updateComment(id: string, data: CommentData) {
    return this.request("update-comment", { id, data }) as Promise<void>;
  }

  removeComment(id: string) {
    return this.request("remove-comment", { id }) as Promise<void>;
  }

  getAllRevisions() {
    return this.request("get-all-revisions") as Promise<RevisionItem[]>;
  }

  setTrackRevisions(enabled: boolean) {
    return this.request("set-track-revisions", { enabled }) as Promise<void>;
  }

  addDemoRevision(text?: string) {
    return this.request("add-demo-revision", { text }) as Promise<RevisionItem[]>;
  }

  goToRevision(id: string) {
    return this.request("go-to-revision", { id }) as Promise<void>;
  }

  acceptRevision(revision: RevisionItem | string) {
    return this.request("accept-revision", { id: getRevisionId(revision) }) as Promise<void>;
  }

  rejectRevision(revision: RevisionItem | string) {
    return this.request("reject-revision", { id: getRevisionId(revision) }) as Promise<void>;
  }

  acceptAllRevisions() {
    return this.request("accept-all-revisions") as Promise<void>;
  }

  rejectAllRevisions() {
    return this.request("reject-all-revisions") as Promise<void>;
  }
}

export class SubframeManager implements MultiInstanceManager {
  readonly containerId: string;
  readonly fileType: FileType;

  private defaultFileName: string;
  private readOnly: boolean;
  private theme: OfficeTheme;
  private officeXmlEvent?: OfficeXmlEventConfig;
  private readonly instanceId: string;
  private readonly targetOrigin: string;
  private readonly onEvent?: SubframeEventHandler;
  private frame: HTMLIFrameElement | null;
  private childReady = false;
  private editorReady = false;
  private disposed = false;
  private requestQueue: Array<{
    action: SubframeRequestAction;
    payload: unknown;
    resolve: (value: unknown) => void;
    reject: (reason?: unknown) => void;
  }> = [];
  private pending = new Map<string, PendingRequest>();
  private editorProxy = new SubframeEditorProxy(this);

  constructor(options: SubframeManagerOptions) {
    this.containerId = options.containerId;
    this.fileType = options.fileType;
    this.defaultFileName = options.defaultFileName;
    this.readOnly = options.readOnly;
    this.theme = options.theme;
    this.officeXmlEvent = options.officeXmlEvent;
    this.instanceId = options.instanceId;
    this.targetOrigin = options.targetOrigin;
    this.frame = options.frame ?? null;
    this.onEvent = options.onEvent;
  }

  configure(options: ConfigureOptions) {
    if (options.defaultFileName) this.defaultFileName = options.defaultFileName;
    if (options.readOnly !== undefined) this.readOnly = options.readOnly;
    if (options.theme) this.theme = options.theme;
    if (options.officeXmlEvent !== undefined) {
      this.officeXmlEvent = options.officeXmlEvent;
    }
  }

  attachFrame(frame: HTMLIFrameElement) {
    if (this.frame === frame) return;
    this.frame = frame;
    this.childReady = false;
    this.editorReady = false;
  }

  detachFrame() {
    this.frame = null;
    this.childReady = false;
    this.editorReady = false;
  }

  /** Called by the parent window's one message listener. */
  handleMessage(event: MessageEvent) {
    if (this.disposed || !this.frame || event.source !== this.frame.contentWindow) {
      return false;
    }
    if (this.targetOrigin !== "*" && event.origin !== this.targetOrigin) {
      return false;
    }
    if (!isSubframeMessage(event.data)) {
      return false;
    }

    const message = event.data;
    if (message.instanceId !== this.instanceId) {
      return false;
    }

    if (message.type === "ready") {
      this.childReady = true;
      this.flushQueue();
      return true;
    }

    if (message.type === "event") {
      this.onEvent?.(message.event, message.payload);
      return true;
    }

    const pending = this.pending.get(message.requestId);
    if (!pending) {
      return true;
    }

    this.pending.delete(message.requestId);
    window.clearTimeout(pending.timer);
    if (message.ok) {
      if (pending.action === "open") this.editorReady = true;
      if (pending.action === "destroy") this.editorReady = false;
      pending.resolve(message.result);
    } else {
      if (pending.action === "open") this.editorReady = false;
      pending.reject(new Error(message.error || "Subframe request failed"));
    }
    return true;
  }

  request(action: SubframeRequestAction, payload?: unknown): Promise<unknown> {
    if (this.disposed) {
      return Promise.reject(new Error("OnlyOffice subframe was disposed"));
    }

    return new Promise((resolve, reject) => {
      const queued = { action, payload, resolve, reject };
      if (this.childReady) {
        this.dispatch(queued);
      } else {
        this.requestQueue.push(queued);
      }
    });
  }

  private flushQueue() {
    if (!this.childReady) return;
    const queue = this.requestQueue.splice(0);
    queue.forEach((request) => this.dispatch(request));
  }

  private dispatch(request: {
    action: SubframeRequestAction;
    payload: unknown;
    resolve: (value: unknown) => void;
    reject: (reason?: unknown) => void;
  }) {
    const frameWindow = this.frame?.contentWindow;
    if (!frameWindow) {
      this.requestQueue.unshift(request);
      return;
    }

    const id = requestId();
    const message: SubframeRequest = {
      source: SUBFRAME_MESSAGE_SOURCE,
      type: "request",
      instanceId: this.instanceId,
      requestId: id,
      action: request.action,
      payload: request.payload,
    };

    const timer = window.setTimeout(() => {
      this.pending.delete(id);
      request.reject(
        new Error(`OnlyOffice subframe request timed out: ${request.action}`),
      );
    }, RPC_TIMEOUT_MS);
    this.pending.set(id, {
      action: request.action,
      resolve: request.resolve,
      reject: request.reject,
      timer,
    });

    try {
      frameWindow.postMessage(message, this.targetOrigin);
    } catch (error) {
      window.clearTimeout(timer);
      this.pending.delete(id);
      request.reject(error);
    }
  }

  private postOneWay(action: SubframeRequestAction, payload?: unknown) {
    const frameWindow = this.frame?.contentWindow;
    if (!this.childReady || !frameWindow) return;
    frameWindow.postMessage(
      {
        source: SUBFRAME_MESSAGE_SOURCE,
        type: "request",
        instanceId: this.instanceId,
        requestId: requestId(),
        action,
        payload,
      } satisfies SubframeRequest,
      this.targetOrigin,
    );
  }

  async openDocument(input: SubframeDocumentInput) {
    const readOnly = input.readOnly ?? this.readOnly;
    this.readOnly = readOnly;
    this.editorReady = false;
    return this.request("open", {
      containerId: this.containerId,
      fileType: this.fileType,
      defaultFileName: this.defaultFileName,
      readOnly,
      theme: this.theme,
      officeXmlEvent: this.officeXmlEvent,
      document: { ...input, readOnly },
    } satisfies SubframeOpenPayload);
  }

  isReady() {
    return this.editorReady;
  }

  getReadOnly() {
    return this.readOnly;
  }

  async setReadOnly(readOnly: boolean) {
    await this.request("set-read-only", { readOnly });
    this.readOnly = readOnly;
  }

  getTheme() {
    return this.theme;
  }

  async setTheme(theme: OfficeTheme) {
    if (this.theme === theme) return;
    this.theme = theme;
    await this.request("set-theme", { theme });
  }

  async toggleLanguage() {
    return (await this.request("toggle-language")) as string;
  }

  getEditor() {
    return this.editorProxy;
  }

  getLogger() {
    return {
      operation: (message: string, context?: unknown) =>
        console.info(`[OnlyOffice subframe:${this.instanceId}] ${message}`, context),
      error: (category: string, message: string, context?: unknown) =>
        console.error(
          `[OnlyOffice subframe:${this.instanceId}] ${category}: ${message}`,
          context,
        ),
    };
  }

  async downloadExport() {
    const result = (await this.request("download")) as
      | { blob?: Blob; fileName?: string }
      | undefined;
    if (result?.blob && result.fileName) {
      downloadBlob(result.blob, result.fileName);
    }
  }

  printLogs() {
    void this.request("print-logs").catch(() => {});
  }

  destroy() {
    this.postOneWay("destroy");
    this.dispose();
  }

  private dispose() {
    this.disposed = true;
    this.requestQueue.splice(0).forEach((request) => {
      request.reject(new Error("OnlyOffice subframe was disposed"));
    });
    this.pending.forEach((request) => {
      window.clearTimeout(request.timer);
      request.reject(new Error("OnlyOffice subframe was disposed"));
    });
    this.pending.clear();
    this.detachFrame();
  }
}

export function getSubframeOrigin(instanceHost: string) {
  if (typeof window === "undefined") return "";

  const current = new URL(window.location.href);
  const hostname = current.hostname;
  const safeHost = instanceHost.replace(/[^a-z0-9-]/gi, "-").toLowerCase();

  if (
    hostname === "localhost" ||
    hostname.endsWith(".localhost") ||
    /^(?:127\.0\.0\.1|0\.0\.0\.0)$/.test(hostname)
  ) {
    current.hostname = `${safeHost}.onlyoffice.localhost`;
    return current.origin;
  }

  const rootHost =
    hostname === "onlyoffice.agent-bridges.com" ||
    hostname.endsWith(".onlyoffice.agent-bridges.com")
      ? "onlyoffice.agent-bridges.com"
      : (() => {
          const domainParts = hostname.split(".");
          const baseDomain =
            domainParts.length > 1
              ? domainParts.slice(-2).join(".")
              : hostname;
          return `onlyoffice.${baseDomain}`;
        })();

  current.hostname = `${safeHost}.${rootHost}`;
  return current.origin;
}
