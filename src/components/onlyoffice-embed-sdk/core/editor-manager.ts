import {
  installOnlyOfficeProxies,
  installReporterWindowHook,
  registerScopedIo,
  type OnlyOfficeProxyWindow,
  type ReporterHookWindow,
  type ScopedIoFactory,
} from "../internal/editor/runtime-bridge";
import {
  callCrossOriginEditor,
  canAccessIframeWindow,
  setCrossOriginReadOnly,
  subscribeEditorResourceLoading,
  subscribeCrossOriginEditorEvent,
  watchCrossOriginIframe,
} from "../internal/editor/runtime-bridge";
import {
  CROSS_ORIGIN_EDITOR_COMMAND,
  CROSS_ORIGIN_EDITOR_EVENT,
} from "../internal/editor/runtime-bridge";
import { EditorServer } from "../internal/editor/server";
import io, { type MockSocketOptions } from "../internal/editor/runtime-bridge";
import { EditorLogger } from "../internal/editor/logger";
import {
  type DocEditor,
  DocumentType,
  type EditorCapturedDocumentSnapshot,
  type EditorDownloadOutput,
  isOfficeXmlSizeLimitExceededError,
  type OfficeXmlSizeLimitExceededPayload,
  type OfficeTheme,
  type OnlyOfficeConnector,
  type OnlyOfficeConnectorOptions,
  type OnlyOfficePluginOptions,
  type User,
} from "../internal/editor/types";
import {
  getDocumentType,
  isOnlyOfficeCdnMode,
} from "../const";
import type { AscWordApiCallback, AscWordApiMethod } from "../type/word-api";
import { type OnlyOfficeIframeWindow } from "../type/sdk-internal";
import {
  ONLYOFFICE_CONTAINER_CONFIG,
  ONLYOFFICE_EVENT_KEYS,
  ONLYOFFICE_ID,
  ASC_RESTRICTION_NONE,
  ASC_RESTRICTION_VIEW,
  type OfficeXmlEventConfig,
} from "../const";
import {
  type CommentChangeHandlers,
  type CommentData,
  type CommentInput,
  type CommentItem,
  isResolvedComment,
  normalizeCommentInput,
  toPluginCommentPayload,
} from "../feature/comments";
import { onlyofficeEventbus, type LoadingChangeData } from "./eventbus";
import {
  type RevisionChangeHandlers,
  type RevisionItem,
  type RevisionsEditorApi,
  collectRevisionItems,
  resolveRevisionShowChanges,
  goToRevision as goToRevisionInSdk,
  applyRevisionChange,
  prepareRevisionReviewDisplay,
} from "../feature/revisions";
import { getOnlyOfficeLang, type OnlyOfficeLang } from "../store/lang";
import { getDocumentObj, setDocumentObj } from "../store/document";
import { initializeOnlyOffice } from "../util/initialize";
import {
  removeOfficeXmlSizeLimitOverlay,
  showOfficeXmlSizeLimitOverlay,
} from "../internal/ui/office-xml-size-limit-overlay";

export type CreateEditorViewOptions = {
  isNew: boolean;
  fileName: string;
  file?: File;
  url?: string;
  loader?: (url: string) => Promise<ArrayBuffer>;
  fileType?: string;
  readOnly?: boolean;
  /** Public compatibility mode; preview uses the native embedded viewer shell. */
  mode?: "edit" | "readonly" | "preview";
  /** Initial native spellcheck policy. */
  spellcheck?: boolean;
  user?: User;
  lang?: string;
  containerId?: string;
  /** Explicit host element; keeps editor DOM and DocsAPI bound to its Document. */
  container?: HTMLElement;
  editorManager?: EditorManager;
  editing?: boolean;
  theme?: OfficeTheme;
  /** Plugin descriptors forwarded to editorConfig.plugins. */
  plugins?: OnlyOfficePluginOptions;
  /** Native Save Copy As / Download As output interceptor. */
  onDownloadOutput?: (
    output: EditorDownloadOutput,
  ) => boolean | Promise<boolean>;
  /** Native toolbar Save persistence handler. It must settle only after storage commits. */
  onUserSave?: (output: EditorDocumentExport) => void | Promise<void>;
  /** 由 EditorManagerFactory.beginLoadSession 生成，用于丢弃过期的异步初始化 */
  loadSession?: number;
  /** 修订审阅页：开启 markup 显示与页边修订气泡 */
  revisionReview?: boolean;
  officeXmlEvent?: OfficeXmlEventConfig;
};

export type EditorDocumentExport = Omit<
  EditorCapturedDocumentSnapshot,
  "capturedDirtyRevision"
> & {
  binData: Uint8Array;
  instanceId: string;
};

function isResourceLoadingPayload(value: unknown): value is Pick<
  LoadingChangeData,
  | "loading"
  | "resourceStatus"
  | "resourceDownload"
  | "transferredBytes"
  | "resourceCount"
> {
  if (!value || typeof value !== "object") return false;
  const payload = value as Partial<LoadingChangeData>;
  return (
    typeof payload.loading === "boolean" &&
    ["checking", "cache-hit", "downloading", "downloaded", "not-observed"].includes(
      String(payload.resourceStatus),
    ) &&
    typeof payload.resourceDownload === "boolean" &&
    typeof payload.transferredBytes === "number" &&
    Number.isFinite(payload.transferredBytes) &&
    payload.transferredBytes >= 0 &&
    typeof payload.resourceCount === "number" &&
    Number.isInteger(payload.resourceCount) &&
    payload.resourceCount >= 0
  );
}

type OnlyOfficeNativeThemeWindow = Window & {
  Common?: {
    UI?: {
      Themes?: {
        setTheme?: (theme: OfficeTheme, source?: string) => void;
      };
    };
  };
};

function getFileType(fileName: string, fileType?: string) {
  return fileType || fileName.split(".").pop()?.toLowerCase() || "docx";
}

type OnlyOfficeSdkApi = {
  i1f?: (priority?: number) => void;
  asyncFontsDocumentEndLoaded?: (priority?: number) => void;
  ra?: { Ghj?: () => void };
  asc_registerCallback?: (type: string, fn: AscWordApiCallback) => void;
  asc_unregisterCallback?: (type: string, fn: AscWordApiCallback) => void;
  asc_nativeGetPDF?: (options?: Record<string, unknown>) => Uint8Array;
  asc_addComment?: (data: CommentData) => string | undefined;
  asc_changeComment?: (id: string, data: CommentData) => void;
  asc_removeComment?: (id: string) => void;
  sync_ChangeCommentData?: (
    id: string,
    data: unknown,
    ...args: unknown[]
  ) => unknown;
  __ONLYOFFICE_RESOLVE_PATCHED__?: boolean;
  asc_selectComment?: (id: string) => void;
  asc_showComment?: (id: string) => void;
  asc_showComments?: () => void;
  asc_hideComments?: () => void;
  asc_SetGlobalTrackRevisions?: (enabled: boolean) => void;
  asc_GetGlobalTrackRevisions?: () => boolean;
  asc_GetRevisionsChangesStack?: () => unknown[];
  asc_HaveRevisionsChanges?: (all?: boolean) => boolean;
  asc_GetTrackRevisionsReportByAuthors?: () => Record<string, unknown[]>;
  asc_BeginViewModeInReview?: (finalMode?: boolean) => void;
  asc_EndViewModeInReview?: () => void;
  asc_SetLocalTrackRevisions?: (enabled: boolean) => void;
  asc_GetNextRevisionsChange?: () => unknown;
  asc_GetPrevRevisionsChange?: () => unknown;
  asc_FollowRevisionMove?: (change: unknown) => void;
  pluginMethod_MoveToNextReviewChange?: (next: boolean) => void;
  pluginMethod_SetDisplayModeInReview?: (mode: string) => void;
  te?: () => unknown;
  asc_AcceptChanges?: (change?: unknown) => void;
  asc_RejectChanges?: (change?: unknown) => void;
  asc_AcceptChangesBySelection?: (all?: boolean) => void;
  asc_RejectChangesBySelection?: (all?: boolean) => void;
  pluginMethod_GetAllComments?: () => Array<{ Id: string; Data: CommentData }>;
  pluginMethod_AddComment?: (data: CommentData) => string | null;
  pluginMethod_ChangeComment?: (id: string, data: CommentData) => void;
  pluginMethod_InputText?: (text: string) => void;
  pluginMethod_PasteText?: (text: string) => void;
  asc_AddText?: (text: string) => void;
  /** OnlyOffice 内部 WOPI 重命名通道；通过 socket rpc 请求宿主重命名。 */
  asc_wopi_renameFile?: (fileName: string) => void;
};

/** iframe 内运行时；混淆字段见 type/sdk-internal.ts，asc_* 为公开 API */
type OnlyOfficeWindow = OnlyOfficeIframeWindow & {
  Asc?: Omit<NonNullable<OnlyOfficeIframeWindow["Asc"]>, "editor"> & {
    editor?: OnlyOfficeSdkApi & {
      asc_Recalculate?: () => void;
    };
  };
};

type ShellMainController = {
  mode?: { isEdit?: boolean; canEdit?: boolean };
};

type WordHeaderView = {
  btnDocMode?: { setVisible?: (visible: boolean) => void };
  btnPDFMode?: { setVisible?: (visible: boolean) => void };
};

export class EditorManager {
  private editor: DocEditor | null = null;
  /**
   * Connector 和 logger 一样属于一个 EditorManager（也就是一个编辑器 iframe）。
   * 9.4 的 DocsAPI 每创建一个 Connector 都会注册一组 postMessage 监听器；
   * 因此不能由业务调用方重复创建。
   */
  private connector: OnlyOfficeConnector | null = null;
  private server: EditorServer;
  private dirty = false;
  /**
   * Monotonic edit revision used to keep an older async snapshot from marking
   * edits that arrived while it was in flight as persisted.
   */
  private dirtyRevision = 0;
  private readOnly = false;
  private editorMode: "edit" | "readonly" | "preview" = "edit";
  private spellcheck = false;
  private editorLang: OnlyOfficeLang = getOnlyOfficeLang();
  private uiTheme: OfficeTheme = "theme-white";
  private plugins?: OnlyOfficePluginOptions;
  private pluginConfigUrls: string[] = [];
  private downloadOutputHandler?: CreateEditorViewOptions["onDownloadOutput"];
  private userSaveHandler?: CreateEditorViewOptions["onUserSave"];
  private userSaveQueue: Promise<void> = Promise.resolve();
  /** 与容器一一对应，供事件与 Connector 使用同一稳定路由键。 */
  private instanceId: string;
  private containerId: string;
  private containerElement: HTMLElement | null;
  private ownerDocument: Document | null;
  private ownerWindow: Window | null;
  private logger: EditorLogger;
  private fileName = "New Document.docx";
  private fileType = "docx";
  private media: Record<string, Uint8Array> = {};
  private comments = new Map<string, CommentData>();
  private crossOriginCommentRefreshPromise: Promise<CommentItem[]> | null =
    null;
  private revisions: RevisionItem[] = [];
  private refreshingRevisions = false;
  private crossOriginRevisionRefreshPromise: Promise<RevisionItem[]> | null =
    null;
  private trackRevisions = false;
  private revisionReviewMode = false;
  private wordContentSyncPromise: Promise<void> | null = null;
  private wordContentSyncTeardown: (() => void) | null = null;
  private scopedIoTeardown: (() => void) | null = null;
  private crossOriginBridgeTeardown: (() => void) | null = null;
  private crossOriginResourceLoadingTeardown: (() => void) | null = null;
  private resourceLoadingState: Pick<
    LoadingChangeData,
    | "resourceStatus"
    | "resourceDownload"
    | "transferredBytes"
    | "resourceCount"
  > = {
    resourceStatus: "checking",
    resourceDownload: false,
    transferredBytes: 0,
    resourceCount: 0,
  };
  private officeXmlSizeLimitOverlayTeardown: (() => void) | null = null;
  private officeXmlSizeLimitPayload: OfficeXmlSizeLimitExceededPayload | null =
    null;
  /** 同一实例的异步 create 串行执行；新请求到达后旧请求只负责清理，不再挂载 iframe。 */
  private createQueue: Promise<void> = Promise.resolve();
  private createGeneration = 0;
  private startupTimeline: {
    generation: number;
    startedAt: number;
  } | null = null;
  private crossOriginReadOnlySyncGeneration = 0;
  private destroyed = false;
  private pendingRename: {
    resolve: (fileName: string) => void;
    reject: (error: Error) => void;
    timer: number;
  } | null = null;

  constructor(
    container: string | HTMLElement = ONLYOFFICE_ID,
    ownerDocument?: Document,
  ) {
    const containerElement = typeof container === "string" ? null : container;
    const containerId =
      typeof container === "string" ? container : container.id;
    const resolvedDocument =
      containerElement?.ownerDocument ??
      ownerDocument ??
      (typeof document === "undefined" ? null : document);

    this.containerId = containerId || ONLYOFFICE_ID;
    this.instanceId = this.containerId;
    this.containerElement = containerElement;
    this.ownerDocument = resolvedDocument;
    this.ownerWindow =
      resolvedDocument?.defaultView ??
      (typeof window === "undefined" ? null : window);
    this.logger = new EditorLogger(this.containerId);
    this.server = new EditorServer({
      getState: () => ({
        readOnly: this.readOnly,
        dirtyRevision: this.dirtyRevision,
      }),
      logger: this.logger,
      onUserSave: (snapshot) => this.handleUserSave(snapshot),
      onDownloadOutput: (output) =>
        this.downloadOutputHandler?.(output) ?? false,
      onLoadError: (error) => {
        this.handleServerLoadError(error);
      },
      onDocumentRename: (fileName) => {
        this.syncRenamedDocument(fileName);
      },
    });
  }

  private bindContainer(container: HTMLElement) {
    this.containerElement = container;
    this.ownerDocument = container.ownerDocument;
    this.ownerWindow =
      container.ownerDocument.defaultView ??
      (typeof window === "undefined" ? null : window);
    if (container.id) {
      this.containerId = container.id;
      this.instanceId = container.id;
    }
  }

  private getOwnerDocument() {
    const ownerDocument =
      this.containerElement?.ownerDocument ??
      this.ownerDocument ??
      (typeof document === "undefined" ? null : document);
    if (!ownerDocument) {
      throw new Error("OnlyOffice editor requires a browser Document");
    }
    this.ownerDocument = ownerDocument;
    return ownerDocument;
  }

  private getOwnerWindow() {
    const ownerWindow =
      this.containerElement?.ownerDocument.defaultView ??
      this.ownerWindow ??
      this.getOwnerDocument().defaultView ??
      (typeof window === "undefined" ? null : window);
    if (!ownerWindow) {
      throw new Error("OnlyOffice editor requires a browser Window");
    }
    this.ownerWindow = ownerWindow;
    return ownerWindow;
  }

  private isCdnMode() {
    return isOnlyOfficeCdnMode(this.getOwnerWindow());
  }

  private buildPluginConfigUrls(): string[] {
    if (!this.plugins?.configUrls.length) return [];

    const ownerWindow = this.getOwnerWindow();
    const requireConfigJson = this.isCdnMode();
    return this.plugins.configUrls.map((configUrl) => {
      const sourceUrl = new URL(configUrl, ownerWindow.location.href);
      if (
        (sourceUrl.protocol !== "http:" && sourceUrl.protocol !== "https:") ||
        sourceUrl.username ||
        sourceUrl.password
      ) {
        throw new Error(`Invalid OnlyOffice plugin config URL: ${configUrl}`);
      }
      if (
        requireConfigJson &&
        !sourceUrl.pathname.endsWith("/config.json")
      ) {
        throw new Error(
          `OnlyOffice CDN plugin config URL must end with config.json: ${configUrl}`,
        );
      }
      sourceUrl.hash = "";
      return sourceUrl.href;
    });
  }

  private getContainerElement() {
    if (
      this.containerElement &&
      this.containerElement.ownerDocument === this.getOwnerDocument()
    ) {
      return this.containerElement;
    }
    return this.getOwnerDocument().getElementById(this.containerId);
  }

  private getOfficeOverlayHostElement() {
    const container = this.getContainerElement();
    return (
      container?.closest<HTMLElement>(
        ONLYOFFICE_CONTAINER_CONFIG.PARENT_SELECTOR,
      ) ?? container
    );
  }

  private clearOfficeXmlSizeLimitOverlay() {
    this.officeXmlSizeLimitOverlayTeardown?.();
    this.officeXmlSizeLimitOverlayTeardown = null;
    this.officeXmlSizeLimitPayload = null;
    const container = this.getOfficeOverlayHostElement();
    if (container) {
      removeOfficeXmlSizeLimitOverlay(container);
    }
  }

  private renderOfficeXmlSizeLimitOverlay() {
    if (!this.officeXmlSizeLimitPayload) {
      return;
    }

    const container = this.getOfficeOverlayHostElement();
    if (container) {
      this.officeXmlSizeLimitOverlayTeardown = showOfficeXmlSizeLimitOverlay(
        container,
        this.officeXmlSizeLimitPayload,
      );
    }
  }

  private handleServerLoadError(error: Error) {
    if (!isOfficeXmlSizeLimitExceededError(error)) {
      return;
    }

    this.officeXmlSizeLimitPayload = error.payload;
    this.renderOfficeXmlSizeLimitOverlay();

    const data = {
      ...error.payload,
      instanceId: this.instanceId,
      containerId: this.containerId,
    };

    try {
      onlyofficeEventbus.emit(
        ONLYOFFICE_EVENT_KEYS.OFFICE_XML_SIZE_LIMIT_EXCEEDED,
        data,
      );
    } catch (eventError) {
      console.error(
        "[EditorManager] officeXmlSizeLimitExceeded handler failed",
        eventError,
      );
    }
  }

  private createScopedIo() {
    return (url?: string, options: MockSocketOptions = {}) => {
      const socket = io(url, {
        ...options,
        logger: this.logger,
        ownerWindow: this.getOwnerWindow(),
      });

      socket.on("connect", () => {
        this.server.handleConnect({ socket });
      });
      socket.on("disconnect", () => {
        this.server.handleDisconnect({ socket });
      });

      return socket;
    };
  }

  private createCrossOriginServerIo(): ScopedIoFactory {
    return () => {
      const socket = io(undefined, {
        deferConnect: true,
        logger: this.logger,
        ownerWindow: this.getOwnerWindow(),
      });
      socket.connected = true;
      socket.disconnected = false;
      return socket;
    };
  }

  private syncEditorBridge() {
    this.crossOriginReadOnlySyncGeneration += 1;
    this.crossOriginBridgeTeardown?.();
    this.crossOriginBridgeTeardown = null;
    this.crossOriginResourceLoadingTeardown?.();
    this.crossOriginResourceLoadingTeardown = null;
    this.scopedIoTeardown?.();
    this.scopedIoTeardown = null;

    this.scopedIoTeardown = registerScopedIo(
      this.containerId,
      this.createScopedIo(),
      this.getOwnerWindow(),
    );
    this.crossOriginResourceLoadingTeardown = subscribeEditorResourceLoading(
      this.containerId,
      () => this.getEditorFrameElement(),
      (payload) => {
        if (!isResourceLoadingPayload(payload)) return;
        this.resourceLoadingState = {
          resourceStatus: payload.resourceStatus,
          resourceDownload: payload.resourceDownload,
          transferredBytes: payload.transferredBytes,
          resourceCount: payload.resourceCount,
        };
        this.emitLoadingChange({
          loading: payload.loading,
          phase: "static-resources",
          ...this.resourceLoadingState,
        });
      },
      this.getOwnerWindow(),
    );

    if (!this.isCdnMode()) {
      return;
    }

    this.crossOriginBridgeTeardown = watchCrossOriginIframe(
      this.containerId,
      () => this.getEditorFrameElement(),
      this.server,
      this.createCrossOriginServerIo(),
      this.getOwnerWindow(),
      this.pluginConfigUrls,
    );
  }

  private emitLoadingChange(
    state: Omit<LoadingChangeData, "instanceId" | "manager">,
  ) {
    this.logger.operation("loading-change", {
      event: "loading-change",
      instanceId: this.instanceId,
      phase: state.phase,
      loading: state.loading,
      resourceStatus: state.resourceStatus,
      resourceDownload: state.resourceDownload,
      transferredBytes: state.transferredBytes,
      resourceCount: state.resourceCount,
    });
    onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, {
      ...state,
      instanceId: this.instanceId,
      manager: this,
    });
  }

  private logStartupPhase(
    generation: number,
    phase:
      | "start"
      | "document-source-ready"
      | "docs-api-ready"
      | "editor-mounted"
      | "document-ready",
  ) {
    const timeline = this.startupTimeline;
    if (!timeline || timeline.generation !== generation) return;
    this.logger.operation("startup-phase", {
      event: "startup-phase",
      instanceId: this.instanceId,
      generation,
      phase,
      elapsedMs: Math.max(
        0,
        Math.round(this.getOwnerWindow().performance.now() - timeline.startedAt),
      ),
      resourceStatus: this.resourceLoadingState.resourceStatus,
      resourceDownload: this.resourceLoadingState.resourceDownload,
    });
  }

  private emitOperationLoading(loading: boolean) {
    this.emitLoadingChange({
      loading,
      phase: loading ? "operation" : "ready",
      ...this.resourceLoadingState,
      resourceDownload: false,
    });
  }

  private syncCrossOriginReadOnly(
    readOnly: boolean,
    retries = 10,
    syncGeneration = ++this.crossOriginReadOnlySyncGeneration,
  ) {
    if (
      this.destroyed ||
      syncGeneration !== this.crossOriginReadOnlySyncGeneration
    ) {
      return false;
    }

    if (
      setCrossOriginReadOnly(
        this.containerId,
        readOnly,
        this.getOwnerWindow(),
      )
    ) {
      return true;
    }

    if (retries > 0) {
      this.getOwnerWindow().setTimeout(() => {
        this.syncCrossOriginReadOnly(readOnly, retries - 1, syncGeneration);
      }, 50);
    }

    return false;
  }

  private getEditorFrameElement() {
    const ownerDocument = this.getOwnerDocument();
    const ownerWindow = this.getOwnerWindow();
    const containerFrame = ownerDocument
      .getElementById(this.containerId)
      ?.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');

    if (containerFrame) {
      return containerFrame;
    }

    const frames = Array.from(
      ownerDocument.querySelectorAll<HTMLIFrameElement>(
        'iframe[name="frameEditor"]',
      ),
    );
    const matchedFrame = frames.find((frame) => {
      try {
        const url = new URL(frame.src, ownerWindow.location.origin);
        return url.searchParams.get("frameEditorId") === this.containerId;
      } catch {
        return false;
      }
    });

    if (matchedFrame) {
      return matchedFrame;
    }

    if (this.containerId === ONLYOFFICE_ID) {
      return frames[0];
    }

    return ownerDocument
      .querySelector<HTMLElement>(
        `${ONLYOFFICE_CONTAINER_CONFIG.PARENT_SELECTOR}[data-onlyoffice-container-id="${this.containerId}"]`,
      )
      ?.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');
  }

  /**
   * Verifies that a plugin message came from a direct child of this
   * manager's live frameEditor. DocsAPI may replace or wrap the original
   * container element, so ownership must be resolved through the manager's
   * current iframe lookup instead of a stale consumer container reference.
   */
  ownsPluginSource(source: WindowProxy) {
    const editorFrame = this.getEditorFrameElement();
    if (!editorFrame?.contentWindow) return false;
    try {
      return source.parent === editorFrame.contentWindow;
    } catch {
      return false;
    }
  }

  private installProxiesOnWindow(win: OnlyOfficeProxyWindow) {
    // 9.4 的协作客户端会把 socket.io 请求相对解析到静态 SDK 根目录。
    // 必须替换 iframe 内的 io，才能将 /doc/... 连接交给内存 EditorServer。
    installOnlyOfficeProxies(win, this.server, this.createScopedIo());
  }

  /**
   * 劫持 iframe 内 XHR/fetch/io，将协作与 downloadAs 请求路由到 mock EditorServer。
   * 必须在 downloadAs 前安装，否则 export 无法收到 /downloadas/ 分片。
   */
  private installIframeProxies() {
    const iframe = this.getEditorFrameElement();
    if (!iframe) {
      throw new Error("Iframe not loaded");
    }

    if (!canAccessIframeWindow(iframe)) {
      return;
    }

    const win = iframe.contentWindow as
      | (OnlyOfficeWindow & ReporterHookWindow)
      | undefined;
    const iframeDoc = iframe.contentDocument;

    if (!iframeDoc || !win) {
      throw new Error("Iframe not loaded");
    }

    if (win.__ONLYOFFICE_PROXIES_INSTALLED__) {
      return;
    }

    this.installProxiesOnWindow(win);
    installReporterWindowHook(win, (target) => {
      this.installProxiesOnWindow(target as OnlyOfficeProxyWindow);
    });
    this.installSaveShortcutBlocker();
  }

  private getEditorFrameWindow() {
    const iframe = this.getEditorFrameElement();

    if (!iframe || !canAccessIframeWindow(iframe)) {
      return undefined;
    }

    return iframe.contentWindow as OnlyOfficeWindow | undefined;
  }

  private getSdkApi() {
    return this.getEditorFrameWindow()?.Asc?.editor;
  }

  private requireSdkApi() {
    const api = this.getSdkApi();

    if (!api) {
      throw new Error("OnlyOffice SDK API is not ready");
    }

    return api;
  }

  private installCommentResolveCleanup() {
    const api = this.getSdkApi();

    if (!api || api.__ONLYOFFICE_RESOLVE_PATCHED__) {
      return;
    }

    api.__ONLYOFFICE_RESOLVE_PATCHED__ = true;

    const removeResolvedComment = (id: unknown) => {
      // OnlyOffice resolves comments through an internal change event first.
      // Removing synchronously during that event can race its own render pass,
      // so schedule the delete for the next tick after the resolved state lands.
      this.getOwnerWindow().setTimeout(() => {
        api.asc_removeComment?.(String(id));
        this.comments.delete(String(id));
      }, 0);
    };

    const originalSyncChange = api.sync_ChangeCommentData?.bind(api);
    if (originalSyncChange) {
      api.sync_ChangeCommentData = (id, data, ...args) => {
        const result = originalSyncChange(id, data, ...args);

        if (isResolvedComment(data)) {
          removeResolvedComment(id);
        }

        return result;
      };
    }

    api.asc_registerCallback?.("asc_onChangeCommentData", (id, data) => {
      if (!isResolvedComment(data)) {
        return;
      }

      removeResolvedComment(id);
    });
  }

  private getDocumentPermissions(editing: boolean) {
    const doc = this.server.getDocument();
    const preview = this.editorMode === "preview";
    return {
      edit: editing && doc.fileType !== "pdf",
      download: !preview,
      chat: false,
      rename: editing,
      protect: editing,
      // 允许接受/拒绝文档内已有修订；不自动进入「修订」录制模式
      review: true,
      print: true,
    };
  }

  /** 关闭 autosave 与保存按钮；保存快捷键由 installSaveShortcutBlocker 拦截。 */
  private buildEditorCustomization() {
    const preview = this.editorMode === "preview";
    const documentType = getDocumentType(this.fileType);
    const zoom =
      !preview &&
      (documentType === DocumentType.Word || documentType === DocumentType.Slide)
        ? -2
        : undefined;
    return {
      uiTheme: this.uiTheme,
      autosave: false,
      help: false,
      about: false,
      hideRightMenu: true,
      compactToolbar: true,
      zoom,
      spellcheck: this.spellcheck,
      layout: {
        header: {
          save: false,
          // Word 头部「编辑 / 审阅 / 查看」切换（PPT/Excel 无此入口）
          editMode: false,
        },
        toolbar: {
          file: {
            save: false,
          },
          save: false,
        },
      },
      review: {
        trackChanges: this.revisionReviewMode,
        // showReviewChanges:true 会在加载时弹出 asc-window review-changes modal-dlg
        showReviewChanges: false,
        ...(this.revisionReviewMode
          ? { reviewDisplay: "markup" as const }
          : {}),
      },
      features: {
        featuresTips: false,
        spellcheck: {
          change: false,
        },
      },
      plugins: Boolean(this.plugins?.configUrls.length),
      anonymous: {
        request: false,
        label: "Local User",
      },
    };
  }

  /** 禁用 Ctrl/Cmd+S 与工具栏保存，避免与 export/downloadAs 共用管道冲突。 */
  private installSaveShortcutBlocker() {
    const win = this.getEditorFrameWindow();
    const doc = win?.document;

    if (!doc || win?.__ONLYOFFICE_SAVE_BLOCKED__) {
      return;
    }

    const blockSaveShortcut = (event: KeyboardEvent) => {
      const key = event.key.toLowerCase();
      if ((event.ctrlKey || event.metaKey) && key === "s") {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
      }
    };

    doc.addEventListener("keydown", blockSaveShortcut, true);
    win.__ONLYOFFICE_SAVE_BLOCKED__ = true;
  }

  /** 文档若自带 w:trackRevisions，OnlyOffice 默认会跟进入修订模式；接入层强制关闭录制。 */
  private applyDefaultReviewSettings() {
    this.trackRevisions = false;
    if (this.isCdnMode()) {
      void this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_SET_TRACK,
        {
          enabled: false,
        },
      ).catch(() => {});
      return;
    }

    const api = this.getSdkApi();
    api?.asc_SetGlobalTrackRevisions?.(false);
  }

  private mergeCommentItems(items: CommentItem[]) {
    for (const item of items) {
      if (isResolvedComment(item.Data)) {
        this.comments.delete(item.Id);
        continue;
      }

      this.comments.set(item.Id, item.Data);
    }
  }

  private fetchCommentsFromSdk(): CommentItem[] {
    const raw = this.getSdkApi()?.pluginMethod_GetAllComments?.();
    if (!Array.isArray(raw)) {
      return [];
    }

    return raw
      .map((item, index) => {
        const source =
          item && typeof item === "object"
            ? (item as Record<string, unknown>)
            : {};
        const id = String(source.Id ?? source.id ?? `comment-${index}`);
        const data = (source.Data ?? source.data ?? source) as CommentData;

        return { Id: id, Data: data };
      })
      .filter((item) => !isResolvedComment(item.Data));
  }

  private refreshCommentsFromSdk() {
    if (this.isCdnMode()) {
      void this.refreshCrossOriginComments().catch(() => {});
      return;
    }

    this.mergeCommentItems(this.fetchCommentsFromSdk());
  }

  private normalizeCrossOriginComments(value: unknown): CommentItem[] {
    if (!Array.isArray(value)) {
      return [];
    }

    return value.map((item, index) => {
      const source =
        item && typeof item === "object"
          ? (item as Record<string, unknown>)
          : {};
      const data =
        source.Data && typeof source.Data === "object"
          ? (source.Data as CommentData)
          : {};

      return {
        Id: String(source.Id ?? `comment-${index}`),
        Data: data,
      };
    });
  }

  private refreshCrossOriginComments() {
    if (this.crossOriginCommentRefreshPromise) {
      return this.crossOriginCommentRefreshPromise;
    }

    this.crossOriginCommentRefreshPromise = this.callCrossOriginComment(
      CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_LIST,
      {},
    )
      .then((items) => {
        const normalized = this.normalizeCrossOriginComments(items);
        this.comments.clear();
        this.mergeCommentItems(normalized);
        return Array.from(this.comments.entries()).map(([Id, Data]) => ({
          Id,
          Data,
        }));
      })
      .finally(() => {
        this.crossOriginCommentRefreshPromise = null;
      });

    return this.crossOriginCommentRefreshPromise;
  }

  private refreshRevisionsFromSdk(options?: { forceRefreshStack?: boolean }) {
    if (this.isCdnMode()) {
      void this.refreshCrossOriginRevisions(options).catch(() => {});
      return;
    }

    if (this.refreshingRevisions) return;

    const api = this.getSdkApi();
    const frameWin = this.getEditorFrameWindow();
    if (!api || !frameWin) {
      this.revisions = [];
      return;
    }

    this.refreshingRevisions = true;
    try {
      this.revisions = collectRevisionItems(
        api as RevisionsEditorApi,
        frameWin,
        options,
      );
    } finally {
      this.refreshingRevisions = false;
    }
  }

  private applyRevisionsFromSdkStack(stack: unknown) {
    if (this.isCdnMode()) {
      this.revisions = this.normalizeCrossOriginRevisions(stack);
      return;
    }

    const api = this.getSdkApi();
    const frameWin = this.getEditorFrameWindow();
    if (!api || !frameWin) return;

    this.revisions = resolveRevisionShowChanges(
      stack,
      api as RevisionsEditorApi,
      frameWin,
    );
  }

  private syncRevisionsAfterMutation() {
    this.refreshRevisionsFromSdk({ forceRefreshStack: true });
  }

  private applyAllRevisionChanges(mode: "accept" | "reject") {
    if (this.isCdnMode()) {
      void this.callCrossOriginRevision(
        mode === "accept"
          ? CROSS_ORIGIN_EDITOR_COMMAND.REVISION_ACCEPT_ALL
          : CROSS_ORIGIN_EDITOR_COMMAND.REVISION_REJECT_ALL,
      ).then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
      });
      return;
    }

    const api = this.requireSdkApi() as RevisionsEditorApi;
    const frameWin = this.getEditorFrameWindow();
    if (!frameWin) {
      return;
    }

    const applyBulk =
      mode === "accept"
        ? () => api.asc_AcceptChanges?.()
        : () => api.asc_RejectChanges?.();

    this.refreshRevisionsFromSdk({ forceRefreshStack: true });
    let guard = 0;
    let stagnant = 0;
    let lastCount = this.revisions.length;

    while (this.haveRevisionsChanges() && guard++ < 20) {
      this.refreshRevisionsFromSdk({ forceRefreshStack: true });
      const [first] = this.revisions;

      if (!first) {
        applyBulk();
        this.syncRevisionsAfterMutation();
        continue;
      }

      applyRevisionChange(mode, first, api, frameWin, this.revisions);
      this.syncRevisionsAfterMutation();

      const nextCount = this.revisions.length;
      if (nextCount >= lastCount && this.haveRevisionsChanges()) {
        stagnant += 1;
        if (stagnant >= 3) {
          applyBulk();
          this.syncRevisionsAfterMutation();
          stagnant = 0;
        }
      } else {
        stagnant = 0;
      }

      lastCount = nextCount;
    }

    this.syncRevisionsAfterMutation();
  }

  private normalizeCrossOriginRevisions(value: unknown): RevisionItem[] {
    if (!Array.isArray(value)) {
      return [];
    }

    return value.map((item, index) => {
      const source =
        item && typeof item === "object"
          ? (item as Record<string, unknown>)
          : {};
      const data =
        source.Data && typeof source.Data === "object"
          ? (source.Data as RevisionItem["Data"])
          : {};

      return {
        Id: String(source.Id ?? `rev-stack-${index}`),
        Index:
          typeof source.Index === "number" && Number.isFinite(source.Index)
            ? source.Index
            : index,
        Data: data,
        Raw: (source.Raw ?? {}) as RevisionItem["Raw"],
      };
    });
  }

  private refreshCrossOriginRevisions(options?: {
    forceRefreshStack?: boolean;
  }) {
    if (this.crossOriginRevisionRefreshPromise) {
      return this.crossOriginRevisionRefreshPromise;
    }

    this.crossOriginRevisionRefreshPromise = this.callCrossOriginRevision(
      CROSS_ORIGIN_EDITOR_COMMAND.REVISION_LIST,
      {
        forceRefreshStack: !!options?.forceRefreshStack,
      },
    )
      .then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
        return this.revisions;
      })
      .finally(() => {
        this.crossOriginRevisionRefreshPromise = null;
      });

    return this.crossOriginRevisionRefreshPromise;
  }

  private callCrossOriginRevision(
    command: string,
    payload: Record<string, unknown> = {},
  ) {
    return callCrossOriginEditor(
      this.containerId,
      command,
      payload,
      5000,
      this.getOwnerWindow(),
    );
  }

  private revisionTargetId(revision: RevisionItem | string) {
    return typeof revision === "string" ? revision : revision.Id;
  }

  addDemoRevision(text = `审批修订 ${new Date().toLocaleTimeString()}`) {
    this.trackRevisions = true;
    if (this.isCdnMode()) {
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_ADD_DEMO,
        { text },
      ).then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
        return this.revisions;
      });
    }

    const api = this.requireSdkApi() as OnlyOfficeSdkApi;
    api.asc_SetGlobalTrackRevisions?.(true);
    api.asc_SetLocalTrackRevisions?.(true);
    if (api.pluginMethod_InputText) {
      api.pluginMethod_InputText(text);
    } else if (api.pluginMethod_PasteText) {
      api.pluginMethod_PasteText(text);
    } else if (api.asc_AddText) {
      api.asc_AddText(text);
    } else {
      throw new Error("OnlyOffice text insertion API is not available");
    }
    this.syncRevisionsAfterMutation();
    return this.revisions;
  }

  private teardownWordContentSync() {
    this.wordContentSyncTeardown?.();
    this.wordContentSyncTeardown = null;
    this.wordContentSyncPromise = null;
  }

  private scheduleWordContentSync() {
    this.getOwnerWindow().setTimeout(() => {
      this.refreshCommentsFromSdk();
      this.refreshRevisionsFromSdk();
    }, 0);
  }

  private ensureWordContentSync() {
    if (this.fileType !== "docx" && getDocumentType(this.fileType) !== "word") {
      return Promise.resolve();
    }

    if (this.wordContentSyncPromise) {
      return this.wordContentSyncPromise;
    }

    this.wordContentSyncPromise = (async () => {
      if (this.isCdnMode()) {
        await Promise.all([
          this.refreshCrossOriginComments(),
          this.refreshCrossOriginRevisions(),
          this.callCrossOriginComment(
            CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_SUBSCRIBE,
            {},
          ),
          this.callCrossOriginRevision(
            CROSS_ORIGIN_EDITOR_COMMAND.REVISION_SUBSCRIBE,
          ),
        ]);

        const unsubscribers = [
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.ADD_COMMENT,
            ([id, data]) => {
              const commentId = String(id);
              const commentData = data as CommentData;
              if (isResolvedComment(commentData)) {
                this.comments.delete(commentId);
                return;
              }

              this.comments.set(commentId, commentData);
            },
            this.getOwnerWindow(),
          ),
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.CHANGE_COMMENT,
            ([id, data]) => {
              const commentId = String(id);
              const commentData = data as CommentData;
              if (isResolvedComment(commentData)) {
                this.comments.delete(commentId);
                return;
              }

              this.comments.set(commentId, commentData);
            },
            this.getOwnerWindow(),
          ),
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.REMOVE_COMMENT,
            ([id]) => {
              this.comments.delete(String(id));
            },
            this.getOwnerWindow(),
          ),
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.SHOW_REVISIONS_CHANGE,
            ([items]) => {
              this.revisions = this.normalizeCrossOriginRevisions(items);
            },
            this.getOwnerWindow(),
          ),
        ];

        this.wordContentSyncTeardown = () => {
          unsubscribers.forEach((unsubscribe) => unsubscribe());
        };
        return;
      }

      const api = this.requireSdkApi();

      this.refreshCommentsFromSdk();
      this.refreshRevisionsFromSdk();

      const unsubscribers = await Promise.all([
        this.subscribe({
          type: "asc_onAddComment",
          fn: (id, data) => {
            const commentId = String(id);
            const commentData = data as CommentData;
            if (isResolvedComment(commentData)) {
              this.comments.delete(commentId);
              return;
            }

            this.comments.set(commentId, commentData);
          },
        }),
        this.subscribe({
          type: "asc_onChangeCommentData",
          fn: (id, data) => {
            const commentId = String(id);
            const commentData = data as CommentData;
            if (isResolvedComment(commentData)) {
              this.comments.delete(commentId);
              return;
            }

            this.comments.set(commentId, commentData);
          },
        }),
        this.subscribe({
          type: "asc_onRemoveComment",
          fn: (id) => {
            this.comments.delete(String(id));
          },
        }),
        this.subscribe({
          type: CROSS_ORIGIN_EDITOR_EVENT.SHOW_REVISIONS_CHANGE,
          fn: (stack) => {
            this.applyRevisionsFromSdkStack(stack);
          },
        }),
      ]);

      this.wordContentSyncTeardown = () => {
        unsubscribers.forEach((unsubscribe) => unsubscribe?.());
      };

      this.scheduleWordContentSync();
      api.asc_Recalculate?.();
    })().catch(() => {
      this.teardownWordContentSync();
    });

    return this.wordContentSyncPromise;
  }

  private getRestrictionSdkApi() {
    return this.getSdkApi() as
      | {
          asc_setRestriction?: (type: number) => void;
          asc_removeRestriction?: (type: number) => void;
          asc_setCanSendChanges?: (enabled: boolean) => void;
        }
      | undefined;
  }

  private getShellMainController() {
    const win = this.getEditorFrameWindow() as OnlyOfficeIframeWindow & {
      PE?: { getController?: (name: string) => ShellMainController };
      DE?: { getController?: (name: string) => ShellMainController };
      getApplication?: () => {
        getController?: (name: string) => ShellMainController;
      };
    };

    return (
      win?.PE?.getController?.("Main") ??
      win?.getApplication?.()?.getController?.("Main") ??
      win?.DE?.getController?.("Main")
    );
  }

  private getWordHeaderView() {
    const win = this.getEditorFrameWindow() as OnlyOfficeIframeWindow & {
      DE?: {
        getController?: (name: string) => {
          getView?: (name: string) => WordHeaderView;
        };
      };
    };

    return win?.DE?.getController?.("Viewport")?.getView?.(
      "Common.Views.Header",
    );
  }

  /** 隐藏 Word 头部「编辑 / 审阅 / 查看」切换（customization + 运行时兜底）。 */
  private hideWordDocModeSwitcher() {
    if (getDocumentType(this.fileType) !== DocumentType.Word) {
      return;
    }

    const header = this.getWordHeaderView();
    header?.btnDocMode?.setVisible?.(false);
    header?.btnPDFMode?.setVisible?.(false);

    const hedset = this.getEditorFrameWindow()?.document?.querySelector(
      '[data-layout-name="header-editMode"]',
    );
    if (hedset instanceof HTMLElement) {
      hedset.style.display = "none";
    }
  }

  private scheduleWordDocModeHide() {
    this.hideWordDocModeSwitcher();
    this.getOwnerWindow().setTimeout(() => this.hideWordDocModeSwitcher(), 0);
  }

  /** 同步 web-apps 工具栏/侧栏的 editing:disable（viewMode 与只读一致）。 */
  private syncShellEditingDisable(
    disabled: boolean,
    documentType = getDocumentType(this.fileType),
  ) {
    const nc = this.getEditorFrameWindow()?.Common?.NotificationCenter;
    if (!nc?.trigger) {
      return;
    }

    if (documentType === DocumentType.Slide) {
      nc.trigger("editing:disable", disabled, {
        viewMode: disabled,
        allowSignature: !disabled,
        rightMenu: { clear: false, disable: true },
        statusBar: true,
        leftMenu: { disable: disabled, previewMode: disabled },
        fileMenu: false,
        comments: { disable: false, previewMode: disabled },
        chat: false,
        review: true,
        viewport: disabled,
        documentHolder: { clear: disabled, disable: true },
        toolbar: true,
        header: { search: false },
        shortcuts: disabled ? false : undefined,
      });
      return;
    }

    if (documentType === DocumentType.Word) {
      if (disabled) {
        nc.trigger("editing:disable", true, {
          viewMode: true,
          reviewMode: false,
          fillFormMode: false,
          viewDocMode: false,
          allowMerge: false,
          allowSignature: false,
          allowProtect: false,
          rightMenu: { clear: true, disable: true },
          statusBar: true,
          leftMenu: { disable: true, previewMode: true },
          fileMenu: { protect: true, history: false },
          navigation: { disable: false, previewMode: true },
          comments: { disable: false, previewMode: true },
          chat: false,
          review: true,
          viewport: true,
          documentHolder: { clear: true, disable: true },
          toolbar: true,
          protect: true,
          header: { search: false, startfill: false },
          shortcuts: false,
        });
      } else {
        nc.trigger("editing:disable", false, {
          viewMode: false,
          reviewMode: false,
          fillFormMode: false,
          viewDocMode: false,
          allowMerge: true,
          allowSignature: false,
          allowProtect: false,
          rightMenu: { clear: false, disable: true },
          statusBar: true,
          leftMenu: { disable: false, previewMode: false },
          fileMenu: false,
          navigation: { disable: false, previewMode: false },
          comments: { disable: false, previewMode: false },
          chat: false,
          review: true,
          viewport: false,
          documentHolder: { clear: false, disable: true },
          toolbar: true,
          protect: true,
        });
      }

      this.scheduleWordDocModeHide();
      return;
    }

    nc.trigger("editing:disable", disabled, {
      viewMode: disabled,
      reviewMode: false,
      fillFormMode: false,
      viewDocMode: false,
      allowMerge: true,
      allowSignature: false,
      allowProtect: false,
      rightMenu: { clear: false, disable: true },
      statusBar: true,
      leftMenu: { disable: false, previewMode: disabled },
      fileMenu: false,
      navigation: { disable: false, previewMode: disabled },
      comments: { disable: false, previewMode: disabled },
      chat: false,
      review: true,
      viewport: false,
      documentHolder: { clear: false, disable: true },
      toolbar: true,
    });
  }

  /** PPT 在通用只读逻辑之上追加：锁定 Main 控制器 + 禁用幻灯片侧栏。 */
  private syncSlideReadOnlyExtras(locked: boolean) {
    if (getDocumentType(this.fileType) !== DocumentType.Slide || !locked) {
      return;
    }
    this.lockShellEditMode();
  }

  /**
   * processRightsChange(true) 在 OnlyOffice 内无效果；false 会 asc_coAuthoringDisconnect 且 mode.isEdit=false。
   * 本地/mock 场景用 asc_setRestriction + 外壳 UI 同步，避免切回编辑仍停留在只读。
   */
  private restoreShellEditMode() {
    const main = this.getShellMainController();

    if (main?.mode) {
      main.mode.isEdit = true;
      main.mode.canEdit = true;
    }

    this.getRestrictionSdkApi()?.asc_setCanSendChanges?.(true);
  }

  private lockShellEditMode() {
    const main = this.getShellMainController();

    if (main?.mode) {
      main.mode.isEdit = false;
      main.mode.canEdit = false;
    }

    this.getRestrictionSdkApi()?.asc_setCanSendChanges?.(false);
  }

  /** 兜底：拦截 SDK 层新增/复制幻灯片（toolbar 锁定之外）。 */
  private installSlideStructureEditBlocker() {
    if (getDocumentType(this.fileType) !== DocumentType.Slide) {
      return;
    }

    const patchApi = (
      api:
        | {
            AddSlide?: (...args: unknown[]) => unknown;
            DublicateSlide?: (...args: unknown[]) => unknown;
            __ONLYOFFICE_SLIDE_BLOCK_PATCHED__?: boolean;
          }
        | undefined,
    ) => {
      if (!api || api.__ONLYOFFICE_SLIDE_BLOCK_PATCHED__) {
        return;
      }

      const guard = <T extends (...args: unknown[]) => unknown>(fn?: T) => {
        if (!fn) {
          return fn;
        }

        const bound = fn.bind(api);
        return (...args: unknown[]) => {
          if (this.readOnly) {
            return undefined;
          }
          return bound(...args);
        };
      };

      api.AddSlide = guard(api.AddSlide);
      api.DublicateSlide = guard(api.DublicateSlide);
      api.__ONLYOFFICE_SLIDE_BLOCK_PATCHED__ = true;
    };

    patchApi(
      this.getSdkApi() as unknown as {
        AddSlide?: (...args: unknown[]) => unknown;
        DublicateSlide?: (...args: unknown[]) => unknown;
        __ONLYOFFICE_SLIDE_BLOCK_PATCHED__?: boolean;
      },
    );
    patchApi(
      (
        this.getShellMainController() as {
          api?: {
            AddSlide?: (...args: unknown[]) => unknown;
            DublicateSlide?: (...args: unknown[]) => unknown;
            __ONLYOFFICE_SLIDE_BLOCK_PATCHED__?: boolean;
          };
        }
      )?.api,
    );
  }

  /** downloadAs → /downloadas/ → 更新 fsMap 中的 Editor.bin。 */
  private async captureDocumentSnapshot(): Promise<EditorCapturedDocumentSnapshot> {
    if (!this.editor) {
      return this.server.getDocumentSnapshot();
    }

    return await this.server.captureCurrentDocument(() => {
      this.installIframeProxies();
      this.editor?.downloadAs("bin");
    });
  }

  /**
   * 只读模式下 downloadAs 可能被 SDK 拦截；导出前临时恢复编辑权再抓取。
   */
  private async captureDocumentSnapshotAllowingReadOnly() {
    if (!this.editor) {
      return this.server.getDocumentSnapshot();
    }

    const locked = this.readOnly;
    if (locked) {
      this.syncEditingRights(true);
    }

    try {
      return await this.captureDocumentSnapshot();
    } finally {
      if (locked) {
        this.syncEditingRights(false);
        this.syncSlideReadOnlyExtras(true);
      }
    }
  }

  private async captureDocumentIfDirty() {
    if (!this.editor || this.readOnly || !this.dirty) {
      return;
    }
    await this.captureStableDocumentSnapshot();
  }

  private async captureStableDocumentSnapshot() {
    const capturedEditor = this.editor;
    if (!capturedEditor || this.destroyed) {
      throw new DOMException(
        "OnlyOffice editor changed while capturing a document snapshot",
        "AbortError",
      );
    }
    const snapshot = await this.captureDocumentSnapshot();
    if (this.destroyed || this.editor !== capturedEditor) {
      throw new DOMException(
        "OnlyOffice editor changed while capturing a document snapshot",
        "AbortError",
      );
    }
    if (
      snapshot.capturedDirtyRevision === undefined ||
      !this.clearDirtyAtRevision(snapshot.capturedDirtyRevision)
    ) {
      throw new Error(
        "OnlyOffice document changed after its snapshot was captured",
      );
    }
    return snapshot;
  }

  private markDirty() {
    this.dirtyRevision += 1;
    this.dirty = true;
  }

  private clearDirtyAtRevision(revision: number) {
    if (this.dirtyRevision !== revision) {
      return false;
    }
    this.dirty = false;
    return true;
  }

  private resetDirtyState() {
    // Invalidate any snapshot started by the previous mounted document.
    this.dirtyRevision += 1;
    this.dirty = false;
  }

  private destroyDocEditorInstance() {
    this.disconnectConnector();
    this.editor?.destroyEditor?.();
    this.editor = null;
    this.comments.clear();
    this.revisions = [];
    this.teardownWordContentSync();
  }

  /** 重建 iframe 前先重置跨域 watcher，确保 about:blank 导航完成后绑定到新窗口。 */
  private remountDocEditor() {
    this.destroyDocEditorInstance();
    this.syncEditorBridge();
    this.mountDocEditor();
  }

  /**
   * 创建 Developer Edition Connector。
   * Connector 运行在父页面，借助 DocsAPI 与编辑器 iframe 通信，因此可用于 CDN 跨域场景。
   */
  createConnector(options?: OnlyOfficeConnectorOptions): OnlyOfficeConnector {
    if (!this.editor) {
      throw new Error("OnlyOffice editor is not ready");
    }

    // 同一编辑器始终复用同一个 Connector。调用方可以在不再使用时主动
    // disconnect；下一次默认创建请求会重新 connect，而不会注册第二套监听器。
    if (this.connector) {
      if (options?.autoconnect !== false && !this.connector.isConnected) {
        this.connector.connect();
      }
      return this.connector;
    }

    const iframe = this.getEditorFrameElement();
    if (!iframe?.contentWindow) {
      throw new Error("OnlyOffice editor iframe is not ready");
    }

    // 9.4 的 DocsAPI 把 Connector 消息固定发送给 `iframeEditor`，但本组件
    // 以 containerId 生成 frameEditorId。先禁止自动连接，替换发送函数后再 connect。
    const connector = this.editor.createConnector({
      ...options,
      autoconnect: false,
    }) as OnlyOfficeConnector & {
      guid?: string;
      sendMessage?: (data: Record<string, unknown>) => void;
    };
    // DocsAPI 将 guid 仅作为 connector 回调的路由键。一个 EditorManager
    // 只维护一个 connector，因此用 containerId 可稳定地与编辑器一一对应。
    connector.guid = this.containerId;
    const iframeUrl = new URL(iframe.src, this.getOwnerWindow().location.href);
    const frameEditorId =
      iframeUrl.searchParams.get("frameEditorId") ?? "iframeEditor";
    connector.sendMessage = (data) => {
      iframe.contentWindow?.postMessage(
        JSON.stringify({
          frameEditorId,
          type: "onExternalPluginMessage",
          subType: "connector",
          data: { ...data, guid: connector.guid },
        }),
        iframeUrl.origin,
      );
    };
    if (options?.autoconnect !== false) {
      connector.connect();
    }
    this.connector = connector;
    return this.connector;
  }

  private disconnectConnector() {
    if (this.connector?.isConnected) {
      this.connector.disconnect();
    }
    this.connector = null;
  }

  /**
   * 初始只读与运行时切只读走同一套 asc_setRestriction。
   * 挂载阶段若 permissions.edit=false，xlsx 等会在打开时样式/格式异常；
   * 因此挂载时保持完整编辑权限，documentReady 后再施加只读限制。
   * CDN 跨域模式通过 iframe 内 bridge 执行同样的 SDK/UI 同步，避免跨域访问 window.Asc。
   */
  private applyInitialReadOnlyState(documentType: DocumentType) {
    if (this.isCdnMode()) {
      this.syncCrossOriginReadOnly(this.readOnly);
      return;
    }

    this.installSlideStructureEditBlocker();
    this.syncEditingRights(false);

    if (documentType === DocumentType.Slide) {
      // 工具栏 delayed render 后再锁一次，确保「新增幻灯片」按钮被 DisableToolbar 处理。
      this.getOwnerWindow().setTimeout(() => {
        if (this.readOnly) {
          this.syncShellEditingDisable(true, documentType);
        }
      }, 0);
    }

    if (documentType === DocumentType.Cell) {
      this.getSdkApi()?.asc_Recalculate?.();
    }
  }

  /** 语言写在 iframe URL 的 lang 参数里，运行时 refreshFile 不会更新界面语言。 */
  private mountDocEditor() {
    const doc = this.server.getDocument();
    const user = this.server.getUser();
    const documentType = getDocumentType(doc.fileType);
    const ownerWindow = this.getOwnerWindow();
    const startupGeneration = this.startupTimeline?.generation;

    this.server.setClient({
      buildVersion: ownerWindow.DocsAPI!.DocEditor.version(),
    });

    this.editor = new ownerWindow.DocsAPI!.DocEditor(this.containerId, {
      document: {
        fileType: doc.fileType,
        key: doc.key,
        title: doc.title,
        url: doc.url,
        permissions: this.getDocumentPermissions(
          this.editorMode !== "preview",
        ),
      },
      documentType,
      editorConfig: {
        // Readonly keeps the native desktop editor shell and is enforced by
        // asc_setRestriction after documentReady. Preview alone uses the
        // upstream embedded viewer shell.
        mode: this.editorMode === "preview" ? "view" : "edit",
        lang: this.editorLang,
        canCoAuthoring: false,
        coEditing: {
          mode: "strict",
          change: false,
        },
        user: {
          ...user,
        },
        customization: this.buildEditorCustomization(),
        embedded:
          this.editorMode === "preview"
            ? { autostart: "document", toolbarDocked: "top" }
            : undefined,
        plugins: this.plugins?.configUrls.length
          ? {
              pluginsData: this.pluginConfigUrls,
              autostart: this.plugins.autostart,
            }
          : undefined,
      },
      events: {
        onAppReady: () => {
          // 尽早安装代理，供 WebSocket auth 与后续 downloadAs 使用。
          this.installIframeProxies();
        },
        onDocumentReady: () => {
          if (startupGeneration !== undefined) {
            this.logStartupPhase(startupGeneration, "document-ready");
            if (this.startupTimeline?.generation === startupGeneration) {
              this.startupTimeline = null;
            }
          }
          this.installSaveShortcutBlocker();
          this.installCommentResolveCleanup();
          this.installSlideStructureEditBlocker();
          this.applyDefaultReviewSettings();
          void this.ensureWordContentSync();
          onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, {
            fileName: doc.title,
            fileType: doc.fileType,
            instanceId: this.instanceId,
            manager: this,
          });

          if (this.readOnly) {
            this.applyInitialReadOnlyState(documentType);
          } else if (documentType === DocumentType.Word) {
            this.scheduleWordDocModeHide();
          }
        },
        onDocumentStateChange: (event: { data: boolean }) => {
          if (event.data) {
            this.markDirty();
          }
        },
        // DocsAPI 仅在注册此回调时暴露「文件 → 重命名」入口。
        // 回调由 OnlyOffice 的内部 Gateway 发出；标题显示已由 iframe 更新，
        // 使用 iframe 内 asc_wopi_renameFile 发起 RPC，由内存服务按 WOPI 协议回包。
        // onRequestRename: (event: { data?: unknown }) => {
        //   const fileName =
        //     typeof event.data === "string" ? event.data : "";
        //   void this.renameDocument(fileName).catch((error) => {
        //     this.logger.error("operation", "OnlyOffice document rename failed", {
        //       fileName,
        //       instanceId: this.instanceId,
        //       containerId: this.containerId,
        //       error: error instanceof Error ? error.message : String(error),
        //     });
        //   });
        // },
        // 不注册 onSave/onSaveDocument：内部保存已禁用，导出统一走 export() → downloadAs。
        onDownloadAs: () => {
          // Required so DocsAPI.downloadAs can request the current editor binary.
        },
      },
      type: this.editorMode === "preview" ? "embedded" : "desktop",
      width: "100%",
      height: "100%",
    });
  }

  private buildRefreshFileConfig(editing: boolean) {
    const doc = this.server.getDocument();
    return {
      document: {
        fileType: doc.fileType,
        key: doc.key,
        title: doc.title,
        url: doc.url,
        permissions: this.getDocumentPermissions(editing),
      },
      documentType: getDocumentType(doc.fileType),
      editorConfig: {
        mode: editing ? "edit" : "view",
        lang: this.editorLang,
      },
      type: "desktop",
      width: "100%",
      height: "100%",
    };
  }

  /** 就地切换编辑权限（主路径 asc_setRestriction；PPT 只读额外 syncSlideReadOnlyExtras）。 */
  private syncEditingRights(editing: boolean) {
    if (!this.editor) {
      return;
    }

    const documentType = getDocumentType(this.fileType);

    if (this.isCdnMode()) {
      this.syncCrossOriginReadOnly(!editing);
      return;
    }

    const sdk = this.getRestrictionSdkApi();

    if (sdk?.asc_setRestriction) {
      if (editing) {
        sdk.asc_removeRestriction?.(ASC_RESTRICTION_VIEW);
        sdk.asc_setRestriction(ASC_RESTRICTION_NONE);
        this.restoreShellEditMode();
      } else {
        sdk.asc_setRestriction(ASC_RESTRICTION_VIEW);
        this.syncSlideReadOnlyExtras(true);
      }
      this.syncShellEditingDisable(!editing, documentType);
      return;
    }

    if (editing) {
      this.restoreShellEditMode();
      this.syncShellEditingDisable(false, documentType);
      this.editor.refreshFile?.(this.buildRefreshFileConfig(true));
    } else {
      this.syncSlideReadOnlyExtras(true);
      if (documentType !== DocumentType.Slide) {
        this.editor.denyEditingRights?.("");
      }
      this.syncShellEditingDisable(true, documentType);
    }
  }

  private createExportData(
    snapshot: ReturnType<EditorServer["getDocumentSnapshot"]>,
  ) {
    const binData = snapshot.binData;

    if (!binData) {
      throw new Error("No OnlyOffice document data is available to export");
    }

    return {
      fileName: snapshot.fileName || this.fileName,
      fileType: snapshot.fileType || this.fileType,
      binData,
      instanceId: this.instanceId,
      media: {
        ...snapshot.media,
        ...this.media,
      },
      themes: snapshot.themes,
    };
  }

  /**
   * Native toolbar Save is a persistence transaction, not an in-memory ACK.
   * Serialize saves, wait for the consumer storage callback, then clear only
   * the dirty revision represented by the accepted Editor.bin snapshot.
   */
  private handleUserSave(snapshot: EditorCapturedDocumentSnapshot) {
    const handler = this.userSaveHandler;
    const operation = this.userSaveQueue.then(async () => {
      const data = this.createExportData(snapshot);
      await handler?.(data);
      if (snapshot.capturedDirtyRevision !== undefined) {
        this.clearDirtyAtRevision(snapshot.capturedDirtyRevision);
      }
      onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.SAVE_DOCUMENT, data);
      onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.ONSAVE, {
        fileName: data.fileName,
        instanceId: this.instanceId,
      });
    });
    this.userSaveQueue = operation.catch(() => undefined);
    return operation;
  }

  private isLoadSessionActive(containerId: string, loadSession?: number) {
    return (
      loadSession === undefined ||
      editorManagerFactory.isLoadSessionActive(
        containerId,
        loadSession,
        this.getOwnerDocument(),
      )
    );
  }

  private isCreateActive(
    containerId: string,
    createGeneration: number,
    loadSession?: number,
  ) {
    return (
      createGeneration === this.createGeneration &&
      this.isLoadSessionActive(containerId, loadSession)
    );
  }

  /**
   * 同一容器快速连续 open 时采用 latest-wins：WASM 转换仍按顺序收尾，但过期请求不会重新挂载编辑器。
   */
  create(options: CreateEditorViewOptions): Promise<this> {
    const containerId =
      options.containerId || this.containerId || ONLYOFFICE_ID;
    const createGeneration = ++this.createGeneration;
    this.destroyed = false;

    const createTask = this.createQueue.then(() =>
      this.createInternal(options, containerId, createGeneration),
    );
    this.createQueue = createTask.then(
      () => undefined,
      () => undefined,
    );
    return createTask;
  }

  private async createInternal(
    options: CreateEditorViewOptions,
    containerId: string,
    createGeneration: number,
  ) {
    const isActive = () =>
      this.isCreateActive(containerId, createGeneration, options.loadSession);

    if (!isActive()) {
      return this;
    }

    this.teardown();
    if (options.container) {
      this.bindContainer(options.container);
      containerId = options.container.id || containerId;
    }
    this.readOnly = !!options.readOnly;
    this.editorMode =
      options.mode ?? (this.readOnly ? "readonly" : "edit");
    this.spellcheck = options.spellcheck ?? false;
    this.revisionReviewMode = !!options.revisionReview;
    this.downloadOutputHandler = options.onDownloadOutput;
    this.userSaveHandler = options.onUserSave;
    this.server.setOfficeXmlEventConfig(options.officeXmlEvent);
    if (options.user) {
      this.server.setUser(options.user);
    }
    this.containerId = containerId;
    this.startupTimeline = {
      generation: createGeneration,
      startedAt: this.getOwnerWindow().performance.now(),
    };
    this.logStartupPhase(createGeneration, "start");

    const fileType = getFileType(options.fileName, options.fileType);
    this.fileName = options.fileName;
    this.fileType = fileType;
    this.media = {};
    this.comments.clear();
    this.revisions = [];
    this.clearOfficeXmlSizeLimitOverlay();
    this.teardownWordContentSync();

    if (!options.isNew && !options.file && !options.url) {
      throw new Error("OnlyOffice requires a file, url, or new document type");
    }

    // Start the small DocsAPI bootstrap before File.arrayBuffer()/URL loading.
    // EditorServer starts x2t asynchronously after the source bytes arrive, so
    // this also preserves the existing x2t/runtime overlap without changing the
    // document load promise or its error ownership.
    const initializeResult = initializeOnlyOffice(this.getOwnerWindow()).then(
      () => ({ ok: true as const }),
      (error: unknown) => ({ ok: false as const, error }),
    );

    if (options.isNew) {
      this.server.openNew(fileType, options.fileName);
    } else if (options.file) {
      await this.server.open(options.file, {
        fileName: options.fileName,
        fileType,
      });
    } else {
      await this.server.openUrl(options.url!, {
        fileName: options.fileName,
        fileType,
        loader: options.loader,
      });
    }
    this.logStartupPhase(createGeneration, "document-source-ready");

    if (this.officeXmlSizeLimitPayload) {
      this.renderOfficeXmlSizeLimitOverlay();
      return this;
    }

    if (!isActive()) {
      if (!this.destroyed) this.teardown();
      return this;
    }

    const initialized = await initializeResult;
    if ("error" in initialized) throw initialized.error;
    this.logStartupPhase(createGeneration, "docs-api-ready");

    if (!isActive()) {
      if (!this.destroyed) this.teardown();
      return this;
    }

    this.editorLang =
      (options.lang as OnlyOfficeLang | undefined) || getOnlyOfficeLang();
    this.uiTheme = options.theme || "theme-white";
    this.plugins = options.plugins
      ? {
          configUrls: [...options.plugins.configUrls],
          autostart: options.plugins.autostart
            ? [...options.plugins.autostart]
            : undefined,
        }
      : undefined;
    this.pluginConfigUrls = this.buildPluginConfigUrls();

    this.syncEditorBridge();
    this.mountDocEditor();
    this.logStartupPhase(createGeneration, "editor-mounted");
    this.renderOfficeXmlSizeLimitOverlay();

    return this;
  }

  exists() {
    return !!this.editor;
  }

  /** 导出链路：downloadAs("bin") → server.resolvePendingExport → SAVE_DOCUMENT 事件。 */
  async export() {
    let snapshot;
    if (this.editor && (!this.readOnly || this.dirty)) {
      snapshot = await this.captureDocumentSnapshotAllowingReadOnly();
      if (snapshot.capturedDirtyRevision !== undefined) {
        this.clearDirtyAtRevision(snapshot.capturedDirtyRevision);
      }
    } else {
      snapshot = this.server.getDocumentSnapshot();
    }
    const data = this.createExportData(snapshot);

    onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.SAVE_DOCUMENT, data);

    return data;
  }

  /**
   * Capture an exportable copy without acknowledging external persistence.
   * Save Copy As / Download As must not clear the source document's dirty
   * revision or emit SAVE_DOCUMENT.
   */
  async exportCopy() {
    const snapshot = this.editor
      ? await this.captureDocumentSnapshotAllowingReadOnly()
      : this.server.getDocumentSnapshot();
    return this.createExportData(snapshot);
  }

  /** Render the live editor model through ONLYOFFICE's native PDF pipeline. */
  async exportPrintPdf() {
    let raw: unknown;
    if (this.isCdnMode()) {
      raw = await callCrossOriginEditor(
        this.containerId,
        CROSS_ORIGIN_EDITOR_COMMAND.DOCUMENT_PRINT_PDF,
        {},
        60_000,
        this.getOwnerWindow(),
      );
    } else {
      const frameWindow = this.getEditorFrameWindow() as
        | (OnlyOfficeWindow & {
            native?: { Save_End?: (...args: unknown[]) => void };
          })
        | undefined;
      const api = frameWindow?.Asc?.editor;
      if (!frameWindow || !api?.asc_nativeGetPDF) {
        throw new Error("OnlyOffice native PDF API is not available");
      }
      const previousNative = frameWindow.native;
      frameWindow.native = {
        ...previousNative,
        Save_End: previousNative?.Save_End ?? (() => undefined),
      };
      try {
        raw = api.asc_nativeGetPDF({ isPrint: true });
      } finally {
        if (previousNative) frameWindow.native = previousNative;
        else delete frameWindow.native;
      }
    }
    let bytes: Uint8Array;
    if (raw instanceof Uint8Array) {
      bytes = raw;
    } else if (raw instanceof ArrayBuffer) {
      bytes = new Uint8Array(raw);
    } else if (ArrayBuffer.isView(raw)) {
      bytes = new Uint8Array(raw.buffer, raw.byteOffset, raw.byteLength);
    } else {
      throw new Error("OnlyOffice native PDF API is not available");
    }
    let header = new TextDecoder("ascii").decode(bytes.subarray(0, 5));
    if (header !== "%PDF-") {
      bytes = await this.server.convertPrintRendererToPdf(bytes);
      header = new TextDecoder("ascii").decode(bytes.subarray(0, 5));
    }
    if (header !== "%PDF-") {
      throw new Error("OnlyOffice print pipeline did not produce a PDF");
    }
    return bytes.slice();
  }

  getUser(): User {
    return this.server.getUser();
  }

  setUser(user: User) {
    this.server.setUser(user);
    this.editor?.setUsers?.([{ id: user.id, name: user.name }]);
  }

  /** 就地切换只读；切到只读前先 downloadAs 落盘，避免后续导出仍是打开时的 Editor.bin。 */
  async setReadOnly(readOnly: boolean) {
    if (this.readOnly === readOnly) {
      return;
    }

    this.emitOperationLoading(true);

    try {
      if (readOnly && this.editor) {
        await this.captureStableDocumentSnapshot();
      }

      this.readOnly = readOnly;
      if (this.isCdnMode()) {
        this.syncCrossOriginReadOnly(readOnly);
        return;
      }

      this.installSlideStructureEditBlocker();
      this.syncEditingRights(!readOnly);
    } finally {
      this.emitOperationLoading(false);
    }
  }

  getReadOnly() {
    return this.readOnly;
  }

  /**
   * Switch between the upstream desktop editor and embedded preview shells.
   * This is a programmatic compatibility API only; it does not inject the
   * custom preview/edit toggle used by older browser-host implementations.
   */
  async setMode(mode: "edit" | "readonly" | "preview") {
    const readOnly = mode !== "edit";
    if (this.editorMode === mode && this.readOnly === readOnly) {
      return;
    }
    if (!this.editor) {
      this.editorMode = mode;
      this.readOnly = readOnly;
      return;
    }

    this.emitOperationLoading(true);
    try {
      await this.captureDocumentIfDirty();
      this.editorMode = mode;
      this.readOnly = readOnly;
      this.remountDocEditor();
    } finally {
      this.emitOperationLoading(false);
    }
  }

  async setLanguage(lang: OnlyOfficeLang) {
    if (this.editorLang === lang) {
      return;
    }

    this.editorLang = lang;

    if (!this.editor) {
      return;
    }

    this.emitOperationLoading(true);

    try {
      await this.captureDocumentIfDirty();
      this.remountDocEditor();
    } finally {
      this.emitOperationLoading(false);
    }
  }

  getTheme(): OfficeTheme {
    return this.uiTheme;
  }

  /**
   * Prefer the editor's native theme controller so a UI-only change does not
   * capture the document, rebuild DocsAPI, or replace the editor iframe.
   * Cross-origin editors expose the same controller through the strict bridge.
   */
  private async applyNativeInterfaceTheme(theme: OfficeTheme) {
    if (this.isCdnMode()) {
      const applied = await callCrossOriginEditor(
        this.containerId,
        CROSS_ORIGIN_EDITOR_COMMAND.INTERFACE_SET_THEME,
        { theme },
        5_000,
        this.getOwnerWindow(),
      );
      return applied === theme;
    }

    const editorWindow = this.getEditorFrameElement()
      ?.contentWindow as OnlyOfficeNativeThemeWindow | null;
    const nativeThemes = editorWindow?.Common?.UI?.Themes;
    if (typeof nativeThemes?.setTheme !== "function") {
      return false;
    }
    nativeThemes.setTheme(theme, "sdk");
    return true;
  }

  /** Apply modern light/dark in place, with remount only as a compatibility fallback. */
  async setTheme(theme: OfficeTheme) {
    if (!this.editor) {
      this.uiTheme = theme;
      return;
    }

    try {
      if (await this.applyNativeInterfaceTheme(theme)) {
        this.uiTheme = theme;
        return;
      }
    } catch {
      // Mixed-version or not-yet-ready editor frames use the existing safe
      // remount path. Hosted releases normally stay on the native path.
    }

    const previousTheme = this.uiTheme;

    this.emitOperationLoading(true);

    try {
      await this.captureDocumentIfDirty();
      this.uiTheme = theme;
      this.remountDocEditor();
    } catch (error) {
      this.uiTheme = previousTheme;
      throw error;
    } finally {
      this.emitOperationLoading(false);
    }
  }

  getInstanceId() {
    return this.instanceId;
  }

  isOfficeXmlSizeLimitExceeded() {
    return !!this.officeXmlSizeLimitPayload;
  }

  getContainerId() {
    return this.containerId;
  }

  getLogger() {
    return this.logger;
  }

  printLogs() {
    this.logger.print();
  }

  getFileName() {
    return this.fileName;
  }

  /**
   * @description 通过 iframe 内 asc_wopi_renameFile 重命名当前实例。
   * SDK 经 socket rpc 收到内存 WOPI 回包后才更新本地标题，因此返回 Promise。
   */
  renameDocument(fileName: string): Promise<string> {
    const requestedName = typeof fileName === "string" ? fileName.trim() : "";
    if (!requestedName) {
      return Promise.reject(
        new Error("OnlyOffice document name cannot be empty"),
      );
    }

    if (this.pendingRename) {
      return Promise.reject(new Error("OnlyOffice document rename is pending"));
    }

    return new Promise<string>((resolve, reject) => {
      const timer = this.getOwnerWindow().setTimeout(() => {
        if (this.pendingRename?.timer !== timer) {
          return;
        }
        this.pendingRename = null;
        reject(new Error("OnlyOffice document rename timed out"));
      }, 5000);

      this.pendingRename = { resolve, reject, timer };

      try {
        if (this.isCdnMode()) {
          void callCrossOriginEditor(
            this.containerId,
            CROSS_ORIGIN_EDITOR_COMMAND.DOCUMENT_RENAME,
            { fileName: requestedName },
            5000,
            this.getOwnerWindow(),
          )
            .then(() => {
              // 跨域 bridge 只能确认 iframe 已调用 SDK API；原生「文件 → 重命名」
              // 路径不一定把 WOPI RPC 回包转回父页。由宿主提交标题，确保导出
              // 快照与当前文件名同步；若 RPC 已先到达，pendingRename 已被清空。
              if (this.pendingRename) {
                this.server.rename(requestedName);
              }
            })
            .catch((error) => this.rejectPendingRename(error));
          return;
        }

        const api = this.requireSdkApi();
        if (!api.asc_wopi_renameFile) {
          throw new Error("OnlyOffice WOPI rename API is not available");
        }
        api.asc_wopi_renameFile(requestedName);
      } catch (error) {
        this.rejectPendingRename(error);
      }
    });
  }

  private syncRenamedDocument(renamedFileName: string) {
    this.fileName = renamedFileName;

    const document = getDocumentObj(this.containerId);
    setDocumentObj(
      {
        ...document,
        fileName: renamedFileName,
      },
      this.containerId,
    );

    const pendingRename = this.pendingRename;
    if (pendingRename) {
      this.getOwnerWindow().clearTimeout(pendingRename.timer);
      this.pendingRename = null;
      pendingRename.resolve(renamedFileName);
    }

    this.logger.operation("OnlyOffice document renamed through WOPI RPC", {
      fileName: renamedFileName,
      instanceId: this.instanceId,
      containerId: this.containerId,
    });
  }

  private rejectPendingRename(error: unknown) {
    const pendingRename = this.pendingRename;
    if (!pendingRename) {
      return;
    }

    this.getOwnerWindow().clearTimeout(pendingRename.timer);
    this.pendingRename = null;
    pendingRename.reject(
      error instanceof Error ? error : new Error(String(error)),
    );
  }

  getContainerParentSelector() {
    return `${ONLYOFFICE_CONTAINER_CONFIG.PARENT_SELECTOR}[data-onlyoffice-container-id="${this.containerId}"], ${ONLYOFFICE_CONTAINER_CONFIG.PARENT_SELECTOR}`;
  }

  getContainerStyle() {
    return ONLYOFFICE_CONTAINER_CONFIG.STYLE;
  }

  updateMedia(key: string, data: Uint8Array) {
    this.media[key] = data;
  }

  getMedia() {
    return { ...this.media };
  }

  isDirty() {
    return this.dirty;
  }

  async subscribe({
    type,
    fn,
  }: {
    type: AscWordApiMethod;
    fn: AscWordApiCallback;
  }) {
    if (this.isCdnMode()) {
      if (
        type === CROSS_ORIGIN_EDITOR_EVENT.ADD_COMMENT ||
        type === CROSS_ORIGIN_EDITOR_EVENT.CHANGE_COMMENT ||
        type === CROSS_ORIGIN_EDITOR_EVENT.REMOVE_COMMENT
      ) {
        await this.callCrossOriginComment(
          CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_SUBSCRIBE,
          {},
        );
        return subscribeCrossOriginEditorEvent(
          this.containerId,
          type,
          (args) => fn(...args),
          this.getOwnerWindow(),
        );
      }

      if (
        type === CROSS_ORIGIN_EDITOR_EVENT.SHOW_REVISIONS_CHANGE ||
        type === CROSS_ORIGIN_EDITOR_EVENT.TRACK_REVISIONS_CHANGE
      ) {
        await this.callCrossOriginRevision(
          CROSS_ORIGIN_EDITOR_COMMAND.REVISION_SUBSCRIBE,
        );
        return subscribeCrossOriginEditorEvent(
          this.containerId,
          type,
          (args) => fn(...args),
          this.getOwnerWindow(),
        );
      }

      if (type === CROSS_ORIGIN_EDITOR_EVENT.DOCUMENT_MODIFIED_CHANGED) {
        await callCrossOriginEditor(
          this.containerId,
          CROSS_ORIGIN_EDITOR_COMMAND.EDITOR_SUBSCRIBE,
          { event: type },
          5000,
          this.getOwnerWindow(),
        );
        return subscribeCrossOriginEditorEvent(
          this.containerId,
          type,
          (args) => fn(...args),
          this.getOwnerWindow(),
        );
      }

      throw new Error(
        `OnlyOffice cross-origin callback is not supported: ${type}`,
      );
    }

    const api = this.requireSdkApi();

    if (!api.asc_registerCallback || !api.asc_unregisterCallback) {
      throw new Error("OnlyOffice callback subscription is not supported");
    }

    api.asc_registerCallback(type, fn);

    return () => {
      api.asc_unregisterCallback?.(type, fn);
    };
  }

  async getAllComments(): Promise<CommentItem[]> {
    if (this.isCdnMode()) {
      return this.refreshCrossOriginComments();
    }

    this.refreshCommentsFromSdk();
    return Array.from(this.comments.entries()).map(([Id, Data]) => ({
      Id,
      Data,
    }));
  }

  refreshComments() {
    return this.getAllComments();
  }

  private createSdkCommentPayload(data: CommentData): unknown {
    const asc = this.getEditorFrameWindow()?.Asc as
      | (Record<string, unknown> & {
          asc_CCommentDataWord?: new (value: unknown) => {
            asc_putText?: (value: string) => void;
            asc_putUserName?: (value: string) => void;
            asc_putTime?: (value: string) => void;
            asc_putQuoteText?: (value: string) => void;
            asc_putSolved?: (value: boolean) => void;
            asc_putUserData?: (value: string) => void;
          };
        })
      | undefined;
    const CommentDataWord = asc?.asc_CCommentDataWord;

    if (!CommentDataWord) {
      return data;
    }

    const comment = new CommentDataWord(null);
    const payload = toPluginCommentPayload(data);

    if (payload.Text != null) {
      comment.asc_putText?.(String(payload.Text));
    }
    if (payload.UserName != null) {
      comment.asc_putUserName?.(String(payload.UserName));
    }
    if (payload.Time != null) {
      comment.asc_putTime?.(String(payload.Time));
    }
    if (payload.QuoteText != null) {
      comment.asc_putQuoteText?.(String(payload.QuoteText));
    }
    if (typeof payload.Solved === "boolean") {
      comment.asc_putSolved?.(payload.Solved);
    }
    if (payload.UserData != null) {
      comment.asc_putUserData?.(String(payload.UserData));
    }

    return comment;
  }

  async callCrossOriginComment(
    command: string,
    payload: Record<string, unknown>,
  ) {
    return callCrossOriginEditor(
      this.containerId,
      command,
      payload,
      5000,
      this.getOwnerWindow(),
    );
  }

  addComment(input: CommentInput) {
    if (this.isCdnMode()) {
      const data = toPluginCommentPayload(normalizeCommentInput(input));
      return this.callCrossOriginComment(
        CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_ADD,
        { data },
      ).then((id) => {
        if (id) {
          this.comments.set(String(id), data);
        }
        return id ? String(id) : "";
      });
    }

    const api = this.requireSdkApi();
    const data = toPluginCommentPayload(normalizeCommentInput(input));
    const id =
      api.pluginMethod_AddComment?.(data) ??
      api.asc_addComment?.(this.createSdkCommentPayload(data) as CommentData);
    if (id) {
      this.comments.set(String(id), data);
    }
    return id ? String(id) : "";
  }

  updateComment(id: string, data: CommentData) {
    if (isResolvedComment(data)) {
      return this.removeComment(id);
    }

    if (this.isCdnMode()) {
      const payload = toPluginCommentPayload(data);
      return this.callCrossOriginComment(
        CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_UPDATE,
        {
          id,
          data: payload,
        },
      ).then(() => {
        this.comments.set(id, payload);
      });
    }

    const api = this.requireSdkApi();
    const payload = toPluginCommentPayload(data);

    if (typeof api.pluginMethod_ChangeComment === "function") {
      api.pluginMethod_ChangeComment(id, payload);
    } else {
      api.asc_changeComment?.(
        id,
        this.createSdkCommentPayload(payload) as CommentData,
      );
    }
    this.comments.set(id, payload);
  }

  removeComment(id: string) {
    if (this.isCdnMode()) {
      return this.callCrossOriginComment(
        CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_REMOVE,
        { id },
      ).then(() => {
        this.comments.delete(id);
      });
    }

    this.requireSdkApi().asc_removeComment?.(id);
    this.comments.delete(id);
  }

  goToComment(
    id: string,
    { showBalloon = false }: { showBalloon?: boolean } = {},
  ) {
    if (this.isCdnMode()) {
      return this.callCrossOriginComment(
        CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_GO_TO,
        { id, showBalloon },
      );
    }

    const api = this.requireSdkApi();
    api.asc_selectComment?.(id);
    if (showBalloon) {
      api.asc_showComment?.(id);
    }
  }

  async registerCommentCallbacks(handlers: CommentChangeHandlers) {
    if (this.isCdnMode()) {
      const unsubscribers: Array<() => void> = [];
      await this.callCrossOriginComment(
        CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_SUBSCRIBE,
        {},
      );

      if (handlers.onAdd) {
        unsubscribers.push(
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.ADD_COMMENT,
            ([id, data]) => {
              const commentId = String(id);
              const commentData = data as CommentData;
              this.comments.set(commentId, commentData);
              handlers.onAdd?.(commentId, commentData);
            },
            this.getOwnerWindow(),
          ),
        );
      }

      if (handlers.onChange) {
        unsubscribers.push(
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.CHANGE_COMMENT,
            ([id, data]) => {
              const commentId = String(id);
              const commentData = data as CommentData;
              if (isResolvedComment(commentData)) {
                this.comments.delete(commentId);
                handlers.onRemove?.(commentId);
                return;
              }

              this.comments.set(commentId, commentData);
              handlers.onChange?.(commentId, commentData);
            },
            this.getOwnerWindow(),
          ),
        );
      }

      if (handlers.onRemove) {
        unsubscribers.push(
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.REMOVE_COMMENT,
            ([id]) => {
              const commentId = String(id);
              this.comments.delete(commentId);
              handlers.onRemove?.(commentId);
            },
            this.getOwnerWindow(),
          ),
        );
      }

      return () => {
        unsubscribers.forEach((unsubscribe) => unsubscribe());
      };
    }

    const unsubscribers = await Promise.all([
      handlers.onAdd
        ? this.subscribe({
            type: "asc_onAddComment",
            fn: (id, data) => {
              const commentId = String(id);
              const commentData = data as CommentData;
              this.comments.set(commentId, commentData);
              handlers.onAdd?.(commentId, commentData);
            },
          })
        : undefined,
      handlers.onChange
        ? this.subscribe({
            type: "asc_onChangeCommentData",
            fn: (id, data) => {
              const commentId = String(id);
              const commentData = data as CommentData;
              if (isResolvedComment(commentData)) {
                this.getOwnerWindow().setTimeout(() => {
                  this.removeComment(commentId);
                  handlers.onRemove?.(commentId);
                }, 0);
                return;
              }

              this.comments.set(commentId, commentData);
              handlers.onChange?.(commentId, commentData);
            },
          })
        : undefined,
      handlers.onRemove
        ? this.subscribe({
            type: "asc_onRemoveComment",
            fn: (id) => {
              const commentId = String(id);
              this.comments.delete(commentId);
              handlers.onRemove?.(commentId);
            },
          })
        : undefined,
    ]);

    return () => {
      unsubscribers.forEach((unsubscribe) => unsubscribe?.());
    };
  }

  setTrackRevisions(enabled: boolean) {
    this.trackRevisions = enabled;
    if (this.isCdnMode()) {
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_SET_TRACK,
        { enabled },
      );
    }

    const api = this.requireSdkApi() as RevisionsEditorApi;
    api.asc_SetGlobalTrackRevisions?.(enabled);
    api.asc_SetLocalTrackRevisions?.(enabled);
  }

  /** 修订审阅页初始化：开启追踪 + markup 显示模式（不会批量接受/拒绝） */
  prepareRevisionReview() {
    this.trackRevisions = true;
    if (this.isCdnMode()) {
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_PREPARE_REVIEW,
      ).then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
      });
    }

    const api = this.requireSdkApi() as RevisionsEditorApi;
    api.asc_SetGlobalTrackRevisions?.(true);
    prepareRevisionReviewDisplay(api, this.getEditorFrameWindow());
  }

  isTrackRevisions() {
    if (this.isCdnMode()) {
      void this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_IS_TRACK,
      ).then((enabled) => {
        this.trackRevisions = !!enabled;
      });
      return this.trackRevisions;
    }

    return !!this.getSdkApi()?.asc_GetGlobalTrackRevisions?.();
  }

  haveRevisionsChanges() {
    if (this.isCdnMode()) {
      void this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_HAVE_CHANGES,
      ).then((hasChanges) => {
        if (!hasChanges) {
          this.revisions = [];
        } else {
          void this.refreshCrossOriginRevisions();
        }
      });
      return this.revisions.length > 0;
    }

    const api = this.getSdkApi() as RevisionsEditorApi | undefined;
    if (typeof api?.asc_HaveRevisionsChanges === "function") {
      if (api.asc_HaveRevisionsChanges(true)) {
        return true;
      }
      if (api.asc_HaveRevisionsChanges()) {
        return true;
      }
    }

    return this.revisions.length > 0;
  }

  async getAllRevisions(): Promise<RevisionItem[]> {
    if (this.isCdnMode()) {
      return this.refreshCrossOriginRevisions({ forceRefreshStack: true });
    }

    this.refreshRevisionsFromSdk({ forceRefreshStack: true });
    return this.revisions;
  }

  refreshRevisions() {
    return this.getAllRevisions();
  }

  goToNextRevision() {
    if (this.isCdnMode()) {
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_NEXT,
      );
    }

    (
      this.getSdkApi() as RevisionsEditorApi | undefined
    )?.asc_GetNextRevisionsChange?.();
  }

  goToPrevRevision() {
    if (this.isCdnMode()) {
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_PREV,
      );
    }

    (
      this.getSdkApi() as RevisionsEditorApi | undefined
    )?.asc_GetPrevRevisionsChange?.();
  }

  goToRevision(id: string) {
    if (this.isCdnMode()) {
      const cached = this.revisions.find((entry) => entry.Id === id);
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_GO_TO,
        {
          id,
          index: cached?.Index,
        },
      );
    }

    const api = this.getSdkApi();
    const frameWin = this.getEditorFrameWindow();
    if (!api || !frameWin) return;

    let cached = this.revisions.find((entry) => entry.Id === id);
    if (!cached) {
      this.refreshRevisionsFromSdk();
      cached = this.revisions.find((entry) => entry.Id === id);
    }

    goToRevisionInSdk(
      cached ?? id,
      api as RevisionsEditorApi,
      frameWin,
      this.revisions,
    );
  }

  acceptRevision(revision: RevisionItem | string) {
    if (this.isCdnMode()) {
      const id = this.revisionTargetId(revision);
      const cached = this.revisions.find((entry) => entry.Id === id);
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_ACCEPT,
        {
          id,
          index: cached?.Index,
        },
      ).then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
      });
    }

    const api = this.getSdkApi();
    const frameWin = this.getEditorFrameWindow();
    if (!api || !frameWin) {
      return;
    }

    if (
      applyRevisionChange(
        "accept",
        revision,
        api as RevisionsEditorApi,
        frameWin,
        this.revisions,
      )
    ) {
      this.syncRevisionsAfterMutation();
    }
  }

  rejectRevision(revision: RevisionItem | string) {
    if (this.isCdnMode()) {
      const id = this.revisionTargetId(revision);
      const cached = this.revisions.find((entry) => entry.Id === id);
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_REJECT,
        {
          id,
          index: cached?.Index,
        },
      ).then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
      });
    }

    const api = this.getSdkApi();
    const frameWin = this.getEditorFrameWindow();
    if (!api || !frameWin) {
      return;
    }

    if (
      applyRevisionChange(
        "reject",
        revision,
        api as RevisionsEditorApi,
        frameWin,
        this.revisions,
      )
    ) {
      this.syncRevisionsAfterMutation();
    }
  }

  acceptAllRevisions() {
    this.applyAllRevisionChanges("accept");
  }

  rejectAllRevisions() {
    this.applyAllRevisionChanges("reject");
  }

  acceptRevisionsBySelection(all?: boolean) {
    if (this.isCdnMode()) {
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_ACCEPT_SELECTION,
        { all },
      ).then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
      });
    }

    this.requireSdkApi().asc_AcceptChangesBySelection?.(all);
  }

  rejectRevisionsBySelection(all?: boolean) {
    if (this.isCdnMode()) {
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_REJECT_SELECTION,
        { all },
      ).then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
      });
    }

    this.requireSdkApi().asc_RejectChangesBySelection?.(all);
  }

  async registerRevisionCallbacks(handlers: RevisionChangeHandlers) {
    if (this.isCdnMode()) {
      const unsubscribers: Array<() => void> = [];
      await this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_SUBSCRIBE,
      );

      if (handlers.onShowChanges) {
        unsubscribers.push(
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.SHOW_REVISIONS_CHANGE,
            ([items]) => {
              this.revisions = this.normalizeCrossOriginRevisions(items);
              handlers.onShowChanges?.(this.revisions);
            },
            this.getOwnerWindow(),
          ),
        );
      }

      if (handlers.onTrackRevisionsChange) {
        unsubscribers.push(
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.TRACK_REVISIONS_CHANGE,
            ([enabled]) => {
              this.trackRevisions = !!enabled;
              handlers.onTrackRevisionsChange?.(!!enabled);
            },
            this.getOwnerWindow(),
          ),
        );
      }

      return () => {
        unsubscribers.forEach((unsubscribe) => unsubscribe());
      };
    }

    const unsubscribers = await Promise.all([
      handlers.onShowChanges
        ? this.subscribe({
            type: CROSS_ORIGIN_EDITOR_EVENT.SHOW_REVISIONS_CHANGE,
            fn: (stack) => {
              const api = this.getSdkApi();
              const frameWin = this.getEditorFrameWindow();
              if (!api || !frameWin) return;

              handlers.onShowChanges?.(
                resolveRevisionShowChanges(
                  stack,
                  api as RevisionsEditorApi,
                  frameWin,
                ),
              );
            },
          })
        : undefined,
      handlers.onTrackRevisionsChange
        ? this.subscribe({
            type: CROSS_ORIGIN_EDITOR_EVENT.TRACK_REVISIONS_CHANGE,
            fn: (enabled) => handlers.onTrackRevisionsChange?.(!!enabled),
          })
        : undefined,
    ]);
    return () => {
      unsubscribers.forEach((unsubscribe) => unsubscribe?.());
    };
  }

  private teardown() {
    this.crossOriginReadOnlySyncGeneration += 1;
    this.rejectPendingRename(new Error("OnlyOffice editor was destroyed"));
    this.teardownWordContentSync();
    this.crossOriginBridgeTeardown?.();
    this.crossOriginBridgeTeardown = null;
    this.crossOriginResourceLoadingTeardown?.();
    this.crossOriginResourceLoadingTeardown = null;
    this.scopedIoTeardown?.();
    this.scopedIoTeardown = null;
    this.clearOfficeXmlSizeLimitOverlay();
    this.disconnectConnector();
    this.editor?.destroyEditor?.();
    this.editor = null;
    this.downloadOutputHandler = undefined;
    this.userSaveHandler = undefined;
    this.resetDirtyState();
    this.comments.clear();
    this.revisions = [];
    this.server.reset();
  }

  destroy() {
    this.createGeneration += 1;
    this.destroyed = true;
    this.teardown();
  }
}

class EditorManagerFactory {
  private defaultManager = new EditorManager();
  private managersByDocument = new WeakMap<
    Document,
    Map<string, EditorManager>
  >();
  private orphanManagers = new Map<string, EditorManager>();
  private managers = new Set<EditorManager>();
  private loadSessionsByDocument = new WeakMap<Document, Map<string, number>>();
  private orphanLoadSessions = new Map<string, number>();

  private resolveDocument(
    context?: HTMLElement | Document | Window,
  ): Document | null {
    if (!context) {
      return typeof document === "undefined" ? null : document;
    }
    if ("document" in context) {
      return context.document;
    }
    if (context.nodeType === 9) {
      return context as Document;
    }
    return context.ownerDocument;
  }

  private getManagerMap(ownerDocument: Document | null) {
    if (!ownerDocument) return this.orphanManagers;
    let managers = this.managersByDocument.get(ownerDocument);
    if (!managers) {
      managers = new Map();
      this.managersByDocument.set(ownerDocument, managers);
    }
    return managers;
  }

  private getLoadSessionMap(ownerDocument: Document | null) {
    if (!ownerDocument) return this.orphanLoadSessions;
    let sessions = this.loadSessionsByDocument.get(ownerDocument);
    if (!sessions) {
      sessions = new Map();
      this.loadSessionsByDocument.set(ownerDocument, sessions);
    }
    return sessions;
  }

  beginLoadSession(
    containerId: string,
    context?: HTMLElement | Document | Window,
  ) {
    const sessions = this.getLoadSessionMap(this.resolveDocument(context));
    const next = (sessions.get(containerId) ?? 0) + 1;
    sessions.set(containerId, next);
    return next;
  }

  isLoadSessionActive(
    containerId: string,
    loadSession: number,
    context?: HTMLElement | Document | Window,
  ) {
    return (
      this.getLoadSessionMap(this.resolveDocument(context)).get(containerId) ===
      loadSession
    );
  }

  getDefault() {
    return this.defaultManager;
  }

  create(container: string | HTMLElement, context?: Document | Window) {
    const containerId =
      typeof container === "string" ? container : container.id;
    const ownerDocument = this.resolveDocument(
      typeof container === "string" ? context : container,
    );
    const managers = this.getManagerMap(ownerDocument);
    const manager =
      managers.get(containerId) ||
      new EditorManager(
        typeof container === "string" ? containerId : container,
        ownerDocument ?? undefined,
      );
    managers.set(containerId, manager);
    this.managers.add(manager);
    return manager;
  }

  get(container: string | HTMLElement, context?: Document | Window) {
    const containerId =
      typeof container === "string" ? container : container.id;
    const ownerDocument = this.resolveDocument(
      typeof container === "string" ? context : container,
    );
    return (
      this.getManagerMap(ownerDocument).get(containerId) ||
      this.create(container, context)
    );
  }

  getAll() {
    return [this.defaultManager, ...this.managers];
  }

  destroy(container: string | HTMLElement, context?: Document | Window) {
    const containerId =
      typeof container === "string" ? container : container.id;
    const ownerDocument = this.resolveDocument(
      typeof container === "string" ? context : container,
    );
    this.beginLoadSession(containerId, ownerDocument ?? undefined);
    const managers = this.getManagerMap(ownerDocument);
    const manager = managers.get(containerId);
    manager?.destroy();
    if (manager) this.managers.delete(manager);
    managers.delete(containerId);
  }

  destroyAll() {
    this.beginLoadSession(this.defaultManager.getInstanceId());
    this.defaultManager.destroy();
    for (const manager of this.managers) {
      manager.destroy();
    }
    this.managers.clear();
    this.managersByDocument = new WeakMap();
    this.orphanManagers.clear();
    this.loadSessionsByDocument = new WeakMap();
    this.orphanLoadSessions.clear();
  }
}
export const editorManagerFactory = new EditorManagerFactory();
export const editorManager = editorManagerFactory.getDefault();
if (typeof window !== "undefined") {
  (window as any).editorManagerFactory = editorManagerFactory;
}
