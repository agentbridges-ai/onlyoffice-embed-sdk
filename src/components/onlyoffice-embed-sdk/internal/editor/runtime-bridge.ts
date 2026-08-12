import { EventEmitter } from "../../util/event-emitter";
import type { EditorLogger } from "./logger";
import type { EditorServer } from "./server";

type Callback = (...args: any[]) => void;

export interface MockSocketOptions {
  /**
   * @description 是否开启调试日志。
   */
  debug?: boolean;
  /**
   * @description 跨域 bridge 父页侧仅路由消息，不自动 connect
   */
  deferConnect?: boolean;
  logger?: EditorLogger;
  ownerWindow?: Pick<Window, "setTimeout">;
}

/**
 * @description 基于内部 EventEmitter 模拟 socket.io-client 的最小运行时。
 */
export class MockSocket<
  ListenEvents extends Record<string, Callback> = any,
  EmitEvents extends Record<string, Callback> = any,
> {
  private static _staticEmitter = new EventEmitter();
  static on<E extends string>(event: E, listener: Callback) {
    MockSocket._staticEmitter.on(event, listener);
  }
  static off<E extends string>(event: E, listener?: Callback) {
    MockSocket._staticEmitter.off(event, listener);
  }

  public active = true;
  public connected: boolean = false;
  public disconnected: boolean = true;
  public recovered = false;
  public id: string = "";
  public io = {
    setOpenToken: () => {},
    setSessionToken: () => {},
    on: function () {
      return this;
    },
    reconnectionAttempts: function () {
      return this;
    },
    reconnectionDelay: function () {
      return this;
    },
    reconnectionDelayMax: function () {
      return this;
    },
    timeout: function () {
      return this;
    },
    transports: function () {
      return this;
    },
    upgrade: function () {
      return this;
    },
    upgradeTransport: function () {
      return this;
    },
    upgradeTimeout: function () {
      return this;
    },
  };

  private _clientEmitter = new EventEmitter();
  private _serverEmitter = new EventEmitter();

  private _debug: boolean;
  private _logger?: EditorLogger;
  private _ownerWindow?: Pick<Window, "setTimeout">;

  constructor(options: MockSocketOptions = {}) {
    this._debug = options.debug;
    this._logger = options.logger;
    this._ownerWindow = options.ownerWindow;
    if (!options.deferConnect) {
      this.connect();
    }
  }

  private _log(...args: any[]): void {
    if (this._logger) {
      this._logger.raw("log", "socket", "mock socket", [
        "[MockSocket]",
        ...args,
      ]);
      return;
    }
    if (this._debug) {
      console.log("[MockSocket]", ...args);
    }
  }

  open() {
    return this.connect();
  }

  compress() {}

  /**
   * @description 模拟连接建立并生成新的 session id。
   */
  connect() {
    this.connected = true;
    this.disconnected = false;
    this.id = Math.random().toString(36).substring(2, 15);
    const onConnected = () => {
      this._trigger("connect");
      MockSocket._staticEmitter.emit("connect", { socket: this });
    };
    if (this._ownerWindow) {
      this._ownerWindow.setTimeout(onConnected, 0);
    } else {
      globalThis.setTimeout(onConnected, 0);
    }
    return this;
  }

  disconnect() {
    this.connected = false;
    this.disconnected = true;
    this._trigger("disconnect");
    MockSocket._staticEmitter.emit("disconnect", { socket: this });
    return this;
  }

  close(): this {
    return this.disconnect();
  }

  /**
   * @description 触发本地监听器，用于模拟服务端下行事件。
   */
  private _trigger(event: string, ...args: any[]): this {
    this._log(`trigger event: ${event}`, ...args);
    this._clientEmitter.emit(event, ...args);
    return this;
  }

  /**
   * @description 注册服务端下行事件监听器。
   */
  on<E extends keyof ListenEvents & string>(
    event: E,
    listener: ListenEvents[E],
  ): this {
    this._clientEmitter.on(event, listener);
    return this;
  }

  /**
   * @description 注册一次性的服务端下行事件监听器。
   */
  once<E extends keyof ListenEvents & string>(
    event: E,
    listener: ListenEvents[E],
  ): this {
    this._clientEmitter.once(event, listener);
    return this;
  }

  /**
   * @description 移除事件监听器。
   */
  off<E extends keyof ListenEvents & string>(
    event: E,
    listener?: ListenEvents[E],
  ): this {
    this._clientEmitter.off(event, listener);
    return this;
  }

  /**
   * @description 移除全部监听器，或移除指定事件的监听器。
   */
  removeAllListeners(event?: string): this {
    this._clientEmitter.removeAllListeners(event);
    return this;
  }

  /**
   * @description 使用 message 事件向服务端发送消息。
   */
  send(...args: Parameters<EmitEvents["message"]>): this {
    if (!this.connected) return this;
    this.emit("message", ...args);
    return this;
  }

  /**
   * @description 向服务端发送事件消息。
   */
  emit<E extends keyof EmitEvents & string>(
    event: E,
    ...args: Parameters<EmitEvents[E]>
  ): this {
    this._log(`emit: ${event}`, ...args);

    if (!this.connected) return this;

    const processEmit = async () => {
      this._serverEmitter.emit(event, ...args);
    };

    if (this._ownerWindow) {
      this._ownerWindow.setTimeout(() => processEmit(), 0);
    } else {
      globalThis.setTimeout(() => processEmit(), 0);
    }
    return this;
  }

  public server = {
    on: (event: string, listener: Callback) => {
      this._serverEmitter.on(event, listener);
    },
    off: (event: string, listener?: Callback) => {
      this._serverEmitter.off(event, listener);
    },
    emit: (event: string, ...args: any[]) => {
      this._clientEmitter.emit(event, ...args);
    },
  };
}

/**
 * @description 兼容 socket.io-client 调用方式的工厂函数。
 */
export function io(_url?: string, options?: MockSocketOptions): MockSocket {
  return new MockSocket(options);
}

/**
 * @description 为 socket.io 兼容层保留函数命名空间类型。
 */
export interface SocketIOStatic {
  (url?: string, options?: MockSocketOptions): MockSocket;
}

const ioWithStatics = io as SocketIOStatic;

/**
 * @description 默认导出保持 socket.io-client 兼容。
 */
export default ioWithStatics;

export interface XHRMiddleware {
  (request: Request): Response | null | Promise<Response | null>;
}

export interface XHRProxyOptions {
  baseUrl?: string;
  shouldBypass?: (url: string, method: string) => boolean;
}

/**
 * Public static surface of the XMLHttpRequest proxy constructor.
 *
 * Keeping the anonymous implementation behind this type prevents private
 * implementation fields from leaking into generated package declarations.
 */
export type XHRProxyConstructor = typeof XMLHttpRequest & {
  use(middleware: XHRMiddleware): void;
  clearMiddlewares(): void;
};

export interface FetchProxyOptions {
  baseUrl?: string;
}

function isForbiddenRequestHeader(name: string) {
  const lowerName = name.toLowerCase();
  return (
    lowerName === "accept-charset" ||
    lowerName === "accept-encoding" ||
    lowerName === "access-control-request-headers" ||
    lowerName === "access-control-request-method" ||
    lowerName === "connection" ||
    lowerName === "content-length" ||
    lowerName === "cookie" ||
    lowerName === "cookie2" ||
    lowerName === "date" ||
    lowerName === "dnt" ||
    lowerName === "expect" ||
    lowerName === "host" ||
    lowerName === "keep-alive" ||
    lowerName === "origin" ||
    lowerName === "referer" ||
    lowerName === "te" ||
    lowerName === "trailer" ||
    lowerName === "transfer-encoding" ||
    lowerName === "upgrade" ||
    lowerName === "via" ||
    lowerName.startsWith("proxy-") ||
    lowerName.startsWith("sec-")
  );
}

/**
 * @description 创建支持中间件拦截的 XMLHttpRequest 代理类。
 * @param BaseXHR 原始 XMLHttpRequest 构造器。
 */
export function createXHRProxy(
  BaseXHR = globalThis.XMLHttpRequest,
  options: XHRProxyOptions = {},
): XHRProxyConstructor {
  return class ProxyXMLHttpRequest extends BaseXHR {
    private static _middlewares: XHRMiddleware[] = [];

    private _isMocked: boolean = false;
    private _requestMethod: string = "GET";
    private _requestUrl: string = "";
    private _requestHeaders: Headers = new Headers();
    private _requestBody: any = null;
    private _responseHeaders: Headers = new Headers();

    /**
     * @description 注册全局 XHR 中间件。
     */
    static use(middleware: XHRMiddleware) {
      this._middlewares.push(middleware);
    }

    /**
     * @description 清空全部 XHR 中间件。
     */
    static clearMiddlewares() {
      this._middlewares = [];
    }

    open(
      method: string,
      url: string | URL,
      async: boolean = true,
      username?: string | null,
      password?: string | null,
    ): void {
      const normalizedUrl = (() => {
        try {
          return options.baseUrl
            ? new URL(url.toString(), options.baseUrl).href
            : url.toString();
        } catch {
          return url.toString();
        }
      })();

      this._requestMethod = method;
      this._requestUrl = normalizedUrl;
      this._requestHeaders = new Headers();
      this._responseHeaders = new Headers();
      this._isMocked = false;

      super.open(
        method,
        normalizedUrl,
        async,
        username ?? undefined,
        password ?? undefined,
      );
    }

    setRequestHeader(name: string, value: string): void {
      if (isForbiddenRequestHeader(name)) {
        return;
      }

      this._requestHeaders.append(name, value);

      if (!this._isMocked) {
        super.setRequestHeader(name, value);
      }
    }

    getResponseHeader(name: string): string | null {
      if (this._isMocked) {
        return this._responseHeaders.get(name);
      }
      return super.getResponseHeader(name);
    }

    getAllResponseHeaders(): string {
      if (!this._isMocked) {
        return super.getAllResponseHeaders();
      }
      return Array.from(this._responseHeaders.entries())
        .map(([key, value]) => `${key}: ${value}`)
        .join("\r\n");
    }

    send(body?: Document | XMLHttpRequestBodyInit | null): void {
      this._requestBody = body;

      if (options.shouldBypass?.(this._requestUrl, this._requestMethod)) {
        super.send(body);
        return;
      }

      this._tryMiddlewares()
        .then((handled) => {
          if (!handled) {
            super.send(body);
          }
        })
        .catch((err) => {
          console.error("ProxyXMLHttpRequest middleware error:", err);
          super.send(body);
        });
    }

    private async _tryMiddlewares(): Promise<boolean> {
      let request: Request;
      try {
        const reqInit: RequestInit = {
          method: this._requestMethod,
          headers: this._requestHeaders,
          body: this._requestBody as BodyInit,
          mode: "cors",
        };

        if (this.withCredentials) {
          reqInit.credentials = "include";
        }

        request = new Request(this._requestUrl, reqInit);
        console.log("ProxyXHR created request:", {
          url: this._requestUrl,
          method: request.method,
          hasBody: !!request.body,
          originalBody: this._requestBody,
        });
      } catch (e) {
        return false;
      }

      for (const mw of ProxyXMLHttpRequest._middlewares) {
        const response = await mw(request.clone());
        if (response) {
          this._isMocked = true;
          await this._handleMockResponse(response);
          return true;
        }
      }

      return false;
    }

    private async _handleMockResponse(response: Response) {
      this._responseHeaders = new Headers(response.headers);

      const emit = (event: Event) => {
        this.dispatchEvent(event);
      };

      emit(new ProgressEvent("loadstart"));

      Object.defineProperty(this, "readyState", {
        value: 2,
        writable: false,
        configurable: true,
      });
      emit(new Event("readystatechange"));

      Object.defineProperty(this, "readyState", {
        value: 3,
        writable: false,
        configurable: true,
      });
      emit(new Event("readystatechange"));

      try {
        let responseData: any;

        if (this.responseType === "json") {
          responseData = await response.json();
        } else if (this.responseType === "arraybuffer") {
          responseData = await response.arrayBuffer();
        } else if (this.responseType === "blob") {
          responseData = await response.blob();
        } else if (this.responseType === "document") {
          const text = await response.text();
          responseData = new DOMParser().parseFromString(text, "text/xml");
        } else {
          responseData = await response.text();
        }

        Object.defineProperty(this, "status", {
          value: response.status,
          writable: false,
          configurable: true,
        });

        Object.defineProperty(this, "statusText", {
          value: response.statusText,
          writable: false,
          configurable: true,
        });

        Object.defineProperty(this, "response", {
          value: responseData,
          writable: false,
          configurable: true,
        });

        Object.defineProperty(this, "responseText", {
          value:
            typeof responseData === "string"
              ? responseData
              : JSON.stringify(responseData),
          writable: false,
          configurable: true,
        });

        Object.defineProperty(this, "responseURL", {
          value: response.url,
          writable: false,
          configurable: true,
        });

        emit(
          new ProgressEvent("progress", {
            lengthComputable: true,
            loaded: 100,
            total: 100,
          }),
        );

        Object.defineProperty(this, "readyState", {
          value: 4,
          writable: false,
          configurable: true,
        });
        emit(new Event("readystatechange"));

        emit(new ProgressEvent("load"));

        emit(new ProgressEvent("loadend"));
      } catch (e) {
        console.error("ProxyXHR: error handling response", e);

        Object.defineProperty(this, "readyState", {
          value: 4,
          writable: false,
          configurable: true,
        });
        emit(new Event("readystatechange"));

        emit(new ProgressEvent("error"));
        emit(new ProgressEvent("loadend"));
      }
    }
  };
}

export type FetchProxy = typeof fetch & {
  use(middleware: XHRMiddleware): void;
  clearMiddlewares(): void;
};

/**
 * @description 创建支持中间件拦截的 fetch 代理函数。
 */
export function createFetchProxy(
  target: (Window & { fetch: typeof fetch }) | typeof fetch = globalThis.fetch,
  options: FetchProxyOptions = {},
): FetchProxy {
  const middlewares: XHRMiddleware[] = [];
  const BaseFetch =
    typeof target === "function" ? target : target.fetch.bind(target);

  const proxy = (async (input: RequestInfo | URL, init?: RequestInit) => {
    let request: Request;
    try {
      const normalizedInput =
        options.baseUrl && !(input instanceof Request)
          ? new URL(input.toString(), options.baseUrl).href
          : input;
      request = new Request(normalizedInput, init);
    } catch (e) {
      return BaseFetch(input, init);
    }

    try {
      for (const mw of middlewares) {
        const response = await mw(request.clone());
        if (response) {
          return response;
        }
      }
    } catch (err) {
      console.error("ProxyFetch middleware error:", err);
      return BaseFetch(request);
    }

    return BaseFetch(request);
  }) as FetchProxy;

  proxy.use = (middleware: XHRMiddleware) => {
    middlewares.push(middleware);
  };

  proxy.clearMiddlewares = () => {
    middlewares.length = 0;
  };

  return proxy;
}

export const CROSS_ORIGIN_BRIDGE_MESSAGE = {
  EDITOR_COMMAND: "editor:command",
  EDITOR_RESPONSE: "editor:response",
  EDITOR_EVENT: "editor:event",
  EDITOR_SET_READONLY: "editor:set-readonly",
} as const;

export const CROSS_ORIGIN_EDITOR_COMMAND = {
  EDITOR_SUBSCRIBE: "editor:subscribe",
  INTERFACE_SET_THEME: "interface:set-theme",
  DOCUMENT_PRINT_PDF: "document:print-pdf",
  DOCUMENT_RENAME: "document:rename",
  COMMENT_ADD: "comment:add",
  COMMENT_UPDATE: "comment:update",
  COMMENT_REMOVE: "comment:remove",
  COMMENT_GO_TO: "comment:go-to",
  COMMENT_LIST: "comment:list",
  COMMENT_SUBSCRIBE: "comment:subscribe",
  REVISION_ADD_DEMO: "revision:add-demo",
  REVISION_LIST: "revision:list",
  REVISION_SET_TRACK: "revision:set-track",
  REVISION_IS_TRACK: "revision:is-track",
  REVISION_HAVE_CHANGES: "revision:have-changes",
  REVISION_PREPARE_REVIEW: "revision:prepare-review",
  REVISION_NEXT: "revision:next",
  REVISION_PREV: "revision:prev",
  REVISION_GO_TO: "revision:go-to",
  REVISION_ACCEPT: "revision:accept",
  REVISION_REJECT: "revision:reject",
  REVISION_ACCEPT_ALL: "revision:accept-all",
  REVISION_REJECT_ALL: "revision:reject-all",
  REVISION_ACCEPT_SELECTION: "revision:accept-selection",
  REVISION_REJECT_SELECTION: "revision:reject-selection",
  REVISION_SUBSCRIBE: "revision:subscribe",
} as const;

export const CROSS_ORIGIN_EDITOR_EVENT = {
  ADD_COMMENT: "asc_onAddComment",
  CHANGE_COMMENT: "asc_onChangeCommentData",
  REMOVE_COMMENT: "asc_onRemoveComment",
  SHOW_REVISIONS_CHANGE: "asc_onShowRevisionsChange",
  TRACK_REVISIONS_CHANGE: "asc_onOnTrackRevisionsChange",
  DOCUMENT_MODIFIED_CHANGED: "asc_onDocumentModifiedChanged",
} as const;

export function shouldBypassOnlyOfficeProxy(url: string, baseUrl: string) {
  const pathname = new URL(url, baseUrl).pathname;

  return (
    pathname.includes("/sdkjs/common/AllFonts.js") ||
    pathname.includes("/sdkjs/common/libfont/") ||
    pathname.includes("/fonts/")
  );
}

export type ScopedIoFactory = (
  url?: string,
  options?: MockSocketOptions,
) => MockSocket;

export type OnlyOfficeParentWindow = Window & {
  __ONLYOFFICE_SCOPED_IO__?: Record<string, ScopedIoFactory>;
};

export function getScopedIoRegistry(
  win: Window = window,
): Record<string, ScopedIoFactory> {
  const parent = win as OnlyOfficeParentWindow;
  if (!parent.__ONLYOFFICE_SCOPED_IO__) {
    parent.__ONLYOFFICE_SCOPED_IO__ = {};
  }
  return parent.__ONLYOFFICE_SCOPED_IO__;
}

export function registerScopedIo(
  containerId: string,
  factory: ScopedIoFactory,
  win: Window = window,
) {
  const registry = getScopedIoRegistry(win);
  registry[containerId] = factory;
  return () => {
    if (registry[containerId] === factory) {
      delete registry[containerId];
    }
  };
}

export function unregisterScopedIo(containerId: string, win: Window = window) {
  const registry = (win as OnlyOfficeParentWindow).__ONLYOFFICE_SCOPED_IO__;
  if (registry) {
    delete registry[containerId];
  }
}

export type OnlyOfficeProxyWindow = Window & {
  __ONLYOFFICE_PROXIES_INSTALLED__?: boolean;
  __ONLYOFFICE_GETFILE_PATCHED__?: boolean;
  __ONLYOFFICE_PRINT_FRAME_PATCHED__?: boolean;
  __ONLYOFFICE_PROXY_SERVER__?: EditorServer;
  HTMLIFrameElement: typeof HTMLIFrameElement;
  URL: typeof URL;
  XMLHttpRequest: typeof XMLHttpRequest;
  Worker: typeof Worker;
  AscCommon?: {
    getFile?: (url: string) => void;
  };
};

export type InstallOnlyOfficeProxyOptions = {
  installIo?: boolean;
};

function extractDownloadFileName(url: string, baseUrl: string) {
  if (!url) {
    return "download";
  }

  try {
    const parsed = new URL(url, baseUrl);
    const fromQuery = parsed.searchParams.get("filename");
    if (fromQuery) {
      return decodeURIComponent(fromQuery);
    }

    const pathname = parsed.pathname;
    const name = decodeURIComponent(pathname.split("/").pop() || "");
    if (name.startsWith("output.")) {
      return name;
    }
    if (name) {
      return name;
    }
  } catch {
    const fallback = url.split("/").pop() || url.split("?")[0];
    if (fallback) {
      return decodeURIComponent(fallback);
    }
  }

  return "download";
}

function parseContentDispositionFileName(header: string | null) {
  if (!header) {
    return "";
  }

  const utf8Match = /filename\*=UTF-8''([^;]+)/i.exec(header);
  if (utf8Match?.[1]) {
    try {
      return decodeURIComponent(utf8Match[1]);
    } catch {
      return utf8Match[1];
    }
  }

  const asciiMatch = /filename="([^"]+)"/i.exec(header);
  if (asciiMatch?.[1]) {
    return asciiMatch[1];
  }

  return "";
}

function scheduleNamedDownloadPatch(
  win: OnlyOfficeProxyWindow,
  server: EditorServer,
  retries = 100,
) {
  if (win.__ONLYOFFICE_GETFILE_PATCHED__) {
    return;
  }

  if (win.AscCommon?.getFile) {
    installNamedDownloadPatch(win, server);
    return;
  }

  if (retries > 0) {
    win.setTimeout(
      () => scheduleNamedDownloadPatch(win, server, retries - 1),
      50,
    );
  }
}

function extractOutputNameFromCacheUrl(url: string) {
  const match = /\/cache\/files\/data\/[^/]+\/(output\.[^/?#]+)/i.exec(url);
  return match?.[1] ?? "";
}

function triggerBlobDownload(
  win: OnlyOfficeProxyWindow,
  blob: Blob,
  fileName: string,
) {
  const objectUrl = win.URL.createObjectURL(blob);
  const link = win.document.createElement("a");
  link.href = objectUrl;
  link.download = fileName;
  link.style.display = "none";
  win.document.body.appendChild(link);
  link.click();
  link.remove();
  win.URL.revokeObjectURL(objectUrl);
}

/**
 * @description mock cache URL 无法通过 iframe 下载携带文件名，改为 fetch + <a download>。
 */
function installNamedDownloadPatch(
  win: OnlyOfficeProxyWindow,
  server: EditorServer,
) {
  win.__ONLYOFFICE_PROXY_SERVER__ = server;
  const ascCommon = win.AscCommon;
  if (!ascCommon?.getFile || win.__ONLYOFFICE_GETFILE_PATCHED__) {
    return;
  }

  const nativeGetFile = ascCommon.getFile.bind(ascCommon);
  const fetchFile = win.fetch.bind(win);

  ascCommon.getFile = (url: string) => {
    if (typeof url !== "string" || !url) {
      nativeGetFile(url);
      return;
    }

    const needsNamedDownload =
      url.includes("/cache/files/") || url.startsWith("blob:");

    if (!needsNamedDownload) {
      nativeGetFile(url);
      return;
    }

    const fallbackName = extractDownloadFileName(url, win.location.href);
    void fetchFile(url)
      .then((response) => {
        if (!response.ok) {
          throw new Error(`Download failed: ${response.status}`);
        }
        const fileName =
          parseContentDispositionFileName(
            response.headers.get("Content-Disposition"),
          ) || fallbackName;
        return response.blob().then((blob) => ({ blob, fileName }));
      })
      .then(({ blob, fileName }) => {
        triggerBlobDownload(win, blob, fileName);
      })
      .catch((err) => {
        console.warn("[OnlyOffice] named download fetch failed:", err);
        const outputName = extractOutputNameFromCacheUrl(url);
        const currentServer = win.__ONLYOFFICE_PROXY_SERVER__ ?? server;
        const blobUrl = outputName
          ? currentServer.getStoredOutputUrl(outputName)
          : null;
        const fileName =
          (outputName && currentServer.getStoredOutputFileName(outputName)) ||
          fallbackName;
        if (blobUrl) {
          void fetchFile(blobUrl)
            .then((response) => response.blob())
            .then((blob) => triggerBlobDownload(win, blob, fileName))
            .catch((fallbackErr) => {
              console.warn("[OnlyOffice] blob download fallback:", fallbackErr);
              nativeGetFile(url);
            });
          return;
        }
        nativeGetFile(url);
      });
  };

  win.__ONLYOFFICE_GETFILE_PATCHED__ = true;
}

/**
 * @description Print uses a hidden iframe navigation instead of XHR/fetch.
 * The mock cache URL is only handled by our proxy middleware, so translate it
 * to a Blob URL before the browser attempts the navigation.
 */
export function guardPendingPrintFrameLoad(target: HTMLIFrameElement) {
  const nativeOnLoad = target.onload;
  let printNavigationPending = true;
  if (typeof nativeOnLoad === "function") {
    target.onload = function (event) {
      if (printNavigationPending) return;
      target.onload = nativeOnLoad;
      return nativeOnLoad.call(target, event);
    };
  }
  return () => {
    printNavigationPending = false;
  };
}

function installPrintFramePatch(
  win: OnlyOfficeProxyWindow,
  server: EditorServer,
) {
  if (win.__ONLYOFFICE_PRINT_FRAME_PATCHED__) {
    return;
  }

  const iframePrototype = win.HTMLIFrameElement?.prototype;
  const srcDescriptor = iframePrototype
    ? Object.getOwnPropertyDescriptor(iframePrototype, "src")
    : undefined;
  if (!srcDescriptor?.set || !srcDescriptor.get) {
    return;
  }

  const nativeSetter = srcDescriptor.set;
  const nativeGetter = srcDescriptor.get;
  const fetchFile = win.fetch.bind(win);

  Object.defineProperty(iframePrototype, "src", {
    configurable: srcDescriptor.configurable,
    enumerable: srcDescriptor.enumerable,
    get: nativeGetter,
    set(value: string) {
      const outputName = extractOutputNameFromCacheUrl(value);
      const blobUrl = outputName ? server.getStoredOutputUrl(outputName) : null;

      if (!blobUrl || !value.includes("/cache/files/")) {
        nativeSetter.call(this, value);
        return;
      }

      const target = this;
      const allowPrintNavigation = guardPendingPrintFrameLoad(target);

      // Fetch through the patched bridge so this also works when the editor
      // iframe is loaded from the CDN origin.
      void fetchFile(blobUrl)
        .then((response) => {
          if (!response.ok) {
            throw new Error(`Print PDF fetch failed: ${response.status}`);
          }
          return response.blob();
        })
        .then((blob) => {
          const objectUrl = win.URL.createObjectURL(blob);
          allowPrintNavigation();
          nativeSetter.call(target, objectUrl);
          // Keep the URL alive through the browser print dialog and release it
          // later; the print controller owns the iframe lifecycle.
          win.setTimeout(() => win.URL.revokeObjectURL(objectUrl), 60_000);
        })
        .catch((error) => {
          console.warn("[OnlyOffice] print PDF fetch failed:", error);
          allowPrintNavigation();
          nativeSetter.call(target, value);
        });
    },
  });

  win.__ONLYOFFICE_PRINT_FRAME_PATCHED__ = true;
}

export function installOnlyOfficeProxies(
  win: OnlyOfficeProxyWindow,
  server: EditorServer,
  createIo: ScopedIoFactory,
  options: InstallOnlyOfficeProxyOptions = {},
) {
  win.__ONLYOFFICE_PROXY_SERVER__ = server;

  if (win.__ONLYOFFICE_PROXIES_INSTALLED__) {
    scheduleNamedDownloadPatch(win, server);
    installPrintFramePatch(win, server);
    return;
  }

  const xhr = createXHRProxy(win.XMLHttpRequest, {
    baseUrl: win.location.href,
    shouldBypass: (url) => shouldBypassOnlyOfficeProxy(url, win.location.href),
  });
  const fetchProxy = createFetchProxy(win, { baseUrl: win.location.href });
  const WorkerCtor = win.Worker;

  xhr.use(
    (request) =>
      win.__ONLYOFFICE_PROXY_SERVER__?.handleRequest(request) ?? null,
  );
  fetchProxy.use(
    (request) =>
      win.__ONLYOFFICE_PROXY_SERVER__?.handleRequest(request) ?? null,
  );

  const patches: Partial<OnlyOfficeProxyWindow> & { io?: ScopedIoFactory } = {
    XMLHttpRequest: xhr,
    fetch: fetchProxy,
    Worker: function Worker(url: string, options?: WorkerOptions) {
      const u = new URL(url, win.location.origin);
      return new WorkerCtor(
        u.href.replace(u.origin, win.location.origin),
        options,
      );
    } as unknown as typeof Worker,
  };
  if (options.installIo !== false) {
    patches.io = createIo;
  }

  Object.assign(win, patches);
  win.__ONLYOFFICE_PROXIES_INSTALLED__ = true;
  scheduleNamedDownloadPatch(win, server);
  installPrintFramePatch(win, server);
}

export const REPORTER_HTML = "index.reporter.html";

export type ReporterBridge = {
  install: (target: Window) => void;
};

export type ReporterHookWindow = Window & {
  open: typeof window.open;
  __ONLYOFFICE_REPORTER_HOOK__?: boolean;
  __ONLYOFFICE_REPORTER_BRIDGE__?: ReporterBridge;
};

export function installReporterWindowHook(
  win: ReporterHookWindow,
  installProxies: (target: Window) => void,
) {
  if (win.__ONLYOFFICE_REPORTER_HOOK__) {
    return;
  }

  win.__ONLYOFFICE_REPORTER_BRIDGE__ = { install: installProxies };
  win.__ONLYOFFICE_REPORTER_HOOK__ = true;

  const nativeOpen = win.open.bind(win);
  win.open = function openReporter(
    url?: string | URL,
    target?: string,
    features?: string,
  ) {
    const popup = nativeOpen(url, target, features);
    const href = typeof url === "string" ? url : (url?.toString() ?? "");

    if (popup && href.includes(REPORTER_HTML)) {
      watchReporterWindow(popup, installProxies, win);
    }

    return popup;
  };
}

function watchReporterWindow(
  popup: Window,
  installProxies: (target: Window) => void,
  ownerWindow: Window,
) {
  const tryInstall = () => {
    if (popup.closed) {
      return true;
    }

    try {
      if (popup.location.href.includes(REPORTER_HTML)) {
        installProxies(popup);
        return true;
      }
    } catch {
      /**
       * @description 弹窗仍在导航过程中，继续轮询等待 reporter 页面可访问。
       */
    }

    return false;
  };

  if (tryInstall()) {
    return;
  }

  const interval = ownerWindow.setInterval(() => {
    if (tryInstall()) {
      ownerWindow.clearInterval(interval);
    }
  }, 1);

  popup.addEventListener(
    "load",
    () => {
      tryInstall();
      ownerWindow.clearInterval(interval);
    },
    { once: true },
  );
}

const BRIDGE_SOURCE = "onlyoffice-bridge";

type BridgeMessage = {
  source?: string;
  type?: string;
  frameEditorId?: string;
  bridgeInstanceId?: string;
  requestId?: string;
  command?: string;
  payload?: Record<string, unknown>;
  result?: unknown;
  error?: string;
  event?: string;
  args?: unknown[];
  method?: string;
  url?: string;
  headers?: Record<string, string>;
  body?: string | null;
  bodyEncoding?: "base64";
  readOnly?: boolean;
  status?: number;
  responseHeaders?: Record<string, string>;
  responseBody?: string | null;
  responseType?: XMLHttpRequestResponseType;
};

type BridgeSession = {
  frameEditorId: string;
  generation: number;
  owner: object;
  server: EditorServer;
  createIo: ScopedIoFactory;
  iframe: HTMLIFrameElement;
  iframeSrc: string;
  targetOrigin: string;
  bridgeInstanceId: string | null;
  socket: MockSocket | null;
  bridgeReady: boolean;
  handshakeSent: boolean;
  pendingReadOnly: boolean | null;
  pluginConfigAllowlist: Set<string>;
  ownerWindow: Window;
  runtime: BridgeRuntimeState;
};

type PendingBridgeRequest = {
  frameEditorId: string;
  generation: number;
  resolve: (value: unknown) => void;
  reject: (reason?: unknown) => void;
  timer: number;
};

type BridgeReady = { session: BridgeSession; generation: number };
type BridgeReadyWaiter = {
  resolve: (ready: BridgeReady) => void;
  reject: (reason?: unknown) => void;
  timer: number;
};

type BridgeRuntimeState = {
  sessions: Map<string, BridgeSession>;
  pendingRequests: Map<string, PendingBridgeRequest>;
  pendingReadyWaiters: Map<string, Set<BridgeReadyWaiter>>;
  editorEventSubscribers: Map<
    string,
    Map<string, Set<(args: unknown[]) => void>>
  >;
  listenerInstalled: boolean;
};

const bridgeRuntimeStates = new WeakMap<Window, BridgeRuntimeState>();
const bridgeWindowsByFrameEditorId = new Map<string, Set<Window>>();

function resolveBridgeOwnerWindow(
  frameEditorId: string,
  ownerWindow?: Window,
  owner?: object,
) {
  if (ownerWindow) return ownerWindow;
  if (owner && "ownerDocument" in owner) {
    const defaultView = (owner as Node).ownerDocument?.defaultView;
    if (defaultView) return defaultView;
  }
  const registeredWindows = bridgeWindowsByFrameEditorId.get(frameEditorId);
  if (registeredWindows?.size === 1) {
    return registeredWindows.values().next().value as Window;
  }
  return typeof window === "undefined" ? undefined : window;
}

function getBridgeRuntimeState(ownerWindow?: Window): BridgeRuntimeState {
  const targetWindow =
    ownerWindow ?? (typeof window === "undefined" ? undefined : window);
  if (!targetWindow) {
    throw new Error("OnlyOffice cross-origin bridge requires a browser Window");
  }

  let runtime = bridgeRuntimeStates.get(targetWindow);
  if (!runtime) {
    runtime = {
      sessions: new Map(),
      pendingRequests: new Map<string, PendingBridgeRequest>(),
      pendingReadyWaiters: new Map(),
      editorEventSubscribers: new Map(),
      listenerInstalled: false,
    };
    bridgeRuntimeStates.set(targetWindow, runtime);
  }
  return runtime;
}

function isBridgeMessage(data: unknown): data is BridgeMessage {
  return (
    typeof data === "object" &&
    data !== null &&
    (data as BridgeMessage).source === BRIDGE_SOURCE
  );
}

function getTargetOrigin(iframe: HTMLIFrameElement): string {
  try {
    return new URL(iframe.src, iframe.ownerDocument.baseURI).origin;
  } catch {
    return "";
  }
}

function postToIframe(
  session: BridgeSession,
  message: BridgeMessage,
  targetWindow?: Window | null,
) {
  const target = targetWindow ?? session.iframe.contentWindow;
  if (!target || !session.targetOrigin) {
    return;
  }
  target.postMessage(
    {
      ...message,
      source: BRIDGE_SOURCE,
      bridgeInstanceId: session.bridgeInstanceId ?? undefined,
    },
    session.targetOrigin,
  );
}

function updateSessionIframe(
  session: BridgeSession,
  iframe: HTMLIFrameElement,
  force = false,
) {
  let iframeUrl: URL;
  try {
    iframeUrl = new URL(iframe.src, iframe.ownerDocument.baseURI);
  } catch {
    return;
  }
  const targetOrigin = iframeUrl.origin;
  if (
    !force &&
    session.iframe === iframe &&
    session.iframeSrc === iframeUrl.href &&
    session.targetOrigin === targetOrigin
  ) {
    return;
  }

  detachSocket(session);
  session.iframe = iframe;
  session.iframeSrc = iframeUrl.href;
  session.targetOrigin = targetOrigin;
  session.bridgeInstanceId = null;
  session.bridgeReady = false;
  session.handshakeSent = false;
}

function isMessageFromSession(event: MessageEvent, session: BridgeSession) {
  return (
    Boolean(session.targetOrigin) &&
    event.origin === session.targetOrigin &&
    event.source === session.iframe.contentWindow
  );
}

function rejectSessionPendingRequests(session: BridgeSession, error: Error) {
  const { pendingRequests } = session.runtime;
  for (const [requestId, pending] of pendingRequests) {
    if (
      pending.frameEditorId !== session.frameEditorId ||
      pending.generation !== session.generation
    ) {
      continue;
    }

    session.ownerWindow.clearTimeout(pending.timer);
    pending.reject(error);
    pendingRequests.delete(requestId);
  }
}

function isCurrentSession(session: BridgeSession, generation: number) {
  return (
    session.runtime.sessions.get(session.frameEditorId) === session &&
    session.generation === generation
  );
}

function attachSocket(session: BridgeSession) {
  if (session.socket) {
    return;
  }

  const socket = session.createIo();
  session.socket = socket;
  session.server.registerSocketTransport(socket);

  const nativeServerEmit = socket.server.emit.bind(socket.server);
  socket.server.emit = (event: string, ...args: unknown[]) => {
    postToIframe(session, {
      type: "socket:event",
      frameEditorId: session.frameEditorId,
      event,
      args,
    });
    return nativeServerEmit(event, ...args);
  };
}

function detachSocket(session: BridgeSession) {
  if (!session.socket) {
    return;
  }
  session.server.handleDisconnect({ socket: session.socket });
  session.socket = null;
  session.handshakeSent = false;
}

function createPluginConfigAllowlist(
  configUrls: readonly string[],
) {
  const allowlist = new Set<string>();
  for (const configUrl of configUrls) {
    const sourceUrl = new URL(configUrl);
    if (
      (sourceUrl.protocol !== "http:" && sourceUrl.protocol !== "https:") ||
      !sourceUrl.pathname.endsWith("/config.json") ||
      sourceUrl.username ||
      sourceUrl.password
    ) {
      throw new Error(
        `Invalid OnlyOffice plugin config source URL: ${configUrl}`,
      );
    }
    sourceUrl.hash = "";
    allowlist.add(sourceUrl.href);
  }
  return allowlist;
}

async function fetchPluginConfig(
  session: BridgeSession,
  sourceUrl: string,
  method: string,
) {
  return session.ownerWindow.fetch(sourceUrl, {
    method,
    credentials: "same-origin",
    redirect: "error",
    headers: { Accept: "application/json" },
  });
}

async function handleHttpRequest(
  session: BridgeSession,
  message: BridgeMessage,
) {
  const generation = session.generation;
  const requestId = message.requestId;
  if (!requestId || !message.url || !message.method) {
    return;
  }

  try {
    const url = new URL(message.url);
    const method = message.method.toUpperCase();
    const isTrustedOrigin =
      (url.protocol === "http:" || url.protocol === "https:") &&
      !url.username &&
      !url.password &&
      url.origin === session.targetOrigin;
    const isCacheRequest =
      method === "GET" &&
      /(?:^|\/)cache\/files\/data\/[^/]+\/.+$/.test(url.pathname);
    const isDownloadRequest =
      method === "POST" && /(?:^|\/)downloadas\/[^/]+$/.test(url.pathname);
    const isUploadRequest =
      method === "POST" && /(?:^|\/)upload\/[^/]+$/.test(url.pathname);
    const isPluginRequest =
      (method === "GET" || method === "HEAD") &&
      session.pluginConfigAllowlist.has(url.href);

    if (
      (!isTrustedOrigin && !isPluginRequest) ||
      (!isCacheRequest &&
        !isDownloadRequest &&
        !isUploadRequest &&
        !isPluginRequest)
    ) {
      if (isCurrentSession(session, generation)) {
        postToIframe(session, {
          type: "http:response",
          frameEditorId: session.frameEditorId,
          requestId,
          status: 403,
          responseBody: null,
        });
      }
      return;
    }

    const init: RequestInit = {
      method,
      headers: message.headers,
    };
    if (method !== "GET" && method !== "HEAD") {
      init.body = decodeRequestBody(message);
    }

    const request = new Request(url, init);
    let response = isPluginRequest
      ? await fetchPluginConfig(session, url.href, method)
      : await session.server.handleRequest(request);
    if (!response) {
      response = new Response("Not Found", { status: 404 });
    }

    const responseHeaders: Record<string, string> = {};
    response.headers.forEach((value, key) => {
      responseHeaders[key] = value;
    });

    let responseBody: string | null = null;
    if (
      message.responseType === "arraybuffer" ||
      message.responseType === "blob"
    ) {
      const buffer = await response.arrayBuffer();
      responseBody = arrayBufferToBase64(buffer);
    } else {
      responseBody = await response.text();
    }

    if (!isCurrentSession(session, generation)) {
      return;
    }

    postToIframe(session, {
      type: "http:response",
      frameEditorId: session.frameEditorId,
      requestId,
      status: response.status,
      responseHeaders,
      responseBody,
      responseType: message.responseType,
    });
  } catch (error) {
    if (isCurrentSession(session, generation)) {
      postToIframe(session, {
        type: "http:response",
        frameEditorId: session.frameEditorId,
        requestId,
        status: 0,
        responseBody: null,
      });
    }
    console.error("[OnlyOfficeBridge] HTTP proxy failed:", error);
  }
}

function arrayBufferToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]!);
  }
  return btoa(binary);
}

function base64ToArrayBuffer(base64: string): ArrayBuffer {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
}

function decodeRequestBody(
  message: BridgeMessage,
): BodyInit | null | undefined {
  if (message.body == null) {
    return null;
  }
  if (message.bodyEncoding === "base64") {
    return base64ToArrayBuffer(message.body);
  }
  return message.body;
}

function describeBridgeState(
  runtime: BridgeRuntimeState,
  frameEditorId: string,
) {
  const session = runtime.sessions.get(frameEditorId);
  if (!session) {
    return `OnlyOffice cross-origin bridge is not ready for ${frameEditorId}: session is not registered`;
  }

  return `OnlyOffice cross-origin bridge is not ready for ${frameEditorId}: iframe=${session.iframe.src || "(empty)"}`;
}

function resolveReadyWaiters(frameEditorId: string, session: BridgeSession) {
  const { pendingReadyWaiters } = session.runtime;
  const waiters = pendingReadyWaiters.get(frameEditorId);
  if (!waiters) {
    return;
  }

  pendingReadyWaiters.delete(frameEditorId);
  waiters.forEach((waiter) => {
    session.ownerWindow.clearTimeout(waiter.timer);
    waiter.resolve({ session, generation: session.generation });
  });
}

function rejectReadyWaiters(
  runtime: BridgeRuntimeState,
  ownerWindow: Window,
  frameEditorId: string,
  error: Error,
) {
  const { pendingReadyWaiters } = runtime;
  const waiters = pendingReadyWaiters.get(frameEditorId);
  if (!waiters) {
    return;
  }

  pendingReadyWaiters.delete(frameEditorId);
  waiters.forEach((waiter) => {
    ownerWindow.clearTimeout(waiter.timer);
    waiter.reject(error);
  });
}

function waitForBridgeReady(
  frameEditorId: string,
  timeout: number,
  ownerWindow: Window,
) {
  const runtime = getBridgeRuntimeState(ownerWindow);
  const { pendingReadyWaiters } = runtime;
  const session = runtime.sessions.get(frameEditorId);
  if (session?.bridgeReady) {
    return Promise.resolve({ session, generation: session.generation });
  }

  return new Promise<BridgeReady>((resolve, reject) => {
    const timer = ownerWindow.setTimeout(() => {
      const waiters = pendingReadyWaiters.get(frameEditorId);
      waiters?.delete(waiter);
      if (waiters?.size === 0) {
        pendingReadyWaiters.delete(frameEditorId);
      }
      reject(new Error(describeBridgeState(runtime, frameEditorId)));
    }, timeout);
    const waiter = { resolve, reject, timer };

    let waiters = pendingReadyWaiters.get(frameEditorId);
    if (!waiters) {
      waiters = new Set();
      pendingReadyWaiters.set(frameEditorId, waiters);
    }
    waiters.add(waiter);
  });
}

function handleBridgeMessage(
  runtime: BridgeRuntimeState,
  event: MessageEvent,
) {
  if (!isBridgeMessage(event.data)) {
    return;
  }

  const message = event.data;
  const frameEditorId = message.frameEditorId;
  if (!frameEditorId) {
    return;
  }

  const session = runtime.sessions.get(frameEditorId);
  if (!session || !isMessageFromSession(event, session)) {
    return;
  }

  const bridgeInstanceId = message.bridgeInstanceId;
  if (typeof bridgeInstanceId !== "string" || !bridgeInstanceId) {
    return;
  }

  if (message.type === "hello") {
    if (session.bridgeInstanceId !== bridgeInstanceId) {
      if (session.bridgeInstanceId !== null) {
        rejectSessionPendingRequests(
          session,
          new DOMException(
            `OnlyOffice cross-origin iframe reloaded: ${frameEditorId}`,
            "AbortError",
          ),
        );
        detachSocket(session);
        session.generation += 1;
      }
      session.bridgeInstanceId = bridgeInstanceId;
      session.bridgeReady = false;
      session.handshakeSent = false;
    }
  } else if (session.bridgeInstanceId !== bridgeInstanceId) {
    return;
  }

  if (message.type !== "hello" && !session.bridgeReady) {
    return;
  }

  if (
    message.type === CROSS_ORIGIN_BRIDGE_MESSAGE.EDITOR_RESPONSE &&
    message.requestId
  ) {
    const pending = runtime.pendingRequests.get(message.requestId);
    if (
      pending?.frameEditorId === frameEditorId &&
      pending.generation === session.generation
    ) {
      session.ownerWindow.clearTimeout(pending.timer);
      runtime.pendingRequests.delete(message.requestId);
      if (message.error) {
        pending.reject(new Error(message.error));
      } else {
        pending.resolve(message.result);
      }
    }
    return;
  }

  if (
    message.type === CROSS_ORIGIN_BRIDGE_MESSAGE.EDITOR_EVENT &&
    message.event
  ) {
    const frameSubscribers = runtime.editorEventSubscribers.get(frameEditorId);
    const subscribers = frameSubscribers?.get(message.event);
    subscribers?.forEach((handler) => {
      handler(message.args ?? []);
    });
    return;
  }

  switch (message.type) {
    case "hello": {
      session.bridgeReady = true;
      resolveReadyWaiters(frameEditorId, session);
      attachSocket(session);
      const targetWindow =
        event.source && "postMessage" in event.source
          ? (event.source as Window)
          : undefined;
      postToIframe(session, { type: "hello:ack", frameEditorId }, targetWindow);
      if (!session.handshakeSent) {
        session.handshakeSent = true;
        const generation = session.generation;
        const socket = session.socket;
        session.ownerWindow.setTimeout(() => {
          if (
            socket &&
            isCurrentSession(session, generation) &&
            session.socket === socket
          ) {
            session.server.sendCoAuthoringHandshake(socket);
          }
        }, 0);
      }
      if (session.pendingReadOnly !== null) {
        postToIframe(session, {
          type: CROSS_ORIGIN_BRIDGE_MESSAGE.EDITOR_SET_READONLY,
          frameEditorId,
          readOnly: session.pendingReadOnly,
        });
        session.pendingReadOnly = null;
      }
      break;
    }
    case "socket:emit": {
      if (!session.socket || !message.event) {
        break;
      }
      (
        session.socket as { emit: (event: string, ...args: unknown[]) => void }
      ).emit(message.event, ...(message.args ?? []));
      break;
    }
    case "http": {
      void handleHttpRequest(session, message);
      break;
    }
    default:
      break;
  }
}

export function registerCrossOriginBridge(
  frameEditorId: string,
  iframe: HTMLIFrameElement,
  server: EditorServer,
  createIo: ScopedIoFactory,
  owner: object = iframe,
  pluginConfigUrls: readonly string[] = [],
) {
  if (!iframe.isConnected) {
    return false;
  }
  const ownerWindow = iframe.ownerDocument.defaultView;
  if (!ownerWindow) {
    return false;
  }
  const runtime = getBridgeRuntimeState(ownerWindow);

  let iframeUrl: URL;
  try {
    iframeUrl = new URL(iframe.src, iframe.ownerDocument.baseURI);
  } catch {
    return false;
  }
  if (
    !iframeUrl.origin ||
    iframeUrl.searchParams.get("frameEditorId") !== frameEditorId
  ) {
    return false;
  }
  const pluginConfigAllowlist = createPluginConfigAllowlist(
    pluginConfigUrls,
  );
  let registeredWindows = bridgeWindowsByFrameEditorId.get(frameEditorId);
  if (!registeredWindows) {
    registeredWindows = new Set();
    bridgeWindowsByFrameEditorId.set(frameEditorId, registeredWindows);
  }
  registeredWindows.add(ownerWindow);

  const existing = runtime.sessions.get(frameEditorId);
  if (existing) {
    const changed =
      existing.owner !== owner ||
      existing.iframe !== iframe ||
      existing.iframeSrc !== iframeUrl.href ||
      existing.targetOrigin !== iframeUrl.origin;
    if (changed) {
      rejectSessionPendingRequests(
        existing,
        new DOMException(
          `OnlyOffice cross-origin iframe was replaced: ${frameEditorId}`,
          "AbortError",
        ),
      );
      rejectReadyWaiters(
        runtime,
        ownerWindow,
        frameEditorId,
        new DOMException(
          `OnlyOffice cross-origin iframe was replaced: ${frameEditorId}`,
          "AbortError",
        ),
      );
      existing.generation += 1;
    }
    updateSessionIframe(existing, iframe, changed);
    existing.server = server;
    existing.createIo = createIo;
    existing.owner = owner;
    existing.pluginConfigAllowlist = pluginConfigAllowlist;
    return true;
  }

  runtime.sessions.set(frameEditorId, {
    frameEditorId,
    generation: 1,
    owner,
    server,
    createIo,
    iframe,
    iframeSrc: iframeUrl.href,
    targetOrigin: iframeUrl.origin,
    bridgeInstanceId: null,
    socket: null,
    bridgeReady: false,
    handshakeSent: false,
    pendingReadOnly: null,
    pluginConfigAllowlist,
    ownerWindow,
    runtime,
  });

  if (!runtime.listenerInstalled) {
    ownerWindow.addEventListener("message", (event) => {
      handleBridgeMessage(runtime, event);
    });
    runtime.listenerInstalled = true;
  }
  return true;
}

export function unregisterCrossOriginBridge(
  frameEditorId: string,
  owner?: object,
  ownerWindow?: Window,
) {
  const targetWindow = resolveBridgeOwnerWindow(
    frameEditorId,
    ownerWindow,
    owner,
  );
  const runtime = getBridgeRuntimeState(targetWindow);
  const session = runtime.sessions.get(frameEditorId);
  if (owner && session?.owner !== owner) {
    return;
  }
  if (session) {
    rejectSessionPendingRequests(
      session,
      new DOMException(
        `OnlyOffice cross-origin bridge was unregistered: ${frameEditorId}`,
        "AbortError",
      ),
    );
    detachSocket(session);
    runtime.sessions.delete(frameEditorId);
    const registeredWindows = bridgeWindowsByFrameEditorId.get(frameEditorId);
    registeredWindows?.delete(session.ownerWindow);
    if (registeredWindows?.size === 0) {
      bridgeWindowsByFrameEditorId.delete(frameEditorId);
    }
    rejectReadyWaiters(
      runtime,
      session.ownerWindow,
      frameEditorId,
      new Error(
        `OnlyOffice cross-origin bridge was unregistered: ${frameEditorId}`,
      ),
    );
  }
  runtime.editorEventSubscribers.delete(frameEditorId);
}

export function setCrossOriginReadOnly(
  frameEditorId: string,
  readOnly: boolean,
  ownerWindow?: Window,
) {
  const runtime = getBridgeRuntimeState(
    resolveBridgeOwnerWindow(frameEditorId, ownerWindow),
  );
  const session = runtime.sessions.get(frameEditorId);
  if (!session) {
    return false;
  }

  session.pendingReadOnly = readOnly;
  if (!session.bridgeReady) {
    return true;
  }

  postToIframe(session, {
    type: CROSS_ORIGIN_BRIDGE_MESSAGE.EDITOR_SET_READONLY,
    frameEditorId,
    readOnly,
  });
  session.pendingReadOnly = null;
  return true;
}

export function callCrossOriginEditor(
  frameEditorId: string,
  command: string,
  payload: Record<string, unknown> = {},
  timeout = 5000,
  ownerWindow?: Window,
) {
  const targetWindow = resolveBridgeOwnerWindow(frameEditorId, ownerWindow);
  if (!targetWindow) {
    return Promise.reject(
      new Error("OnlyOffice cross-origin bridge requires a browser Window"),
    );
  }
  const runtime = getBridgeRuntimeState(targetWindow);
  const requestId = `editor-${Date.now()}-${Math.random()
    .toString(36)
    .slice(2)}`;

  return waitForBridgeReady(frameEditorId, timeout, targetWindow).then(
    ({ session, generation }) =>
      new Promise<unknown>((resolve, reject) => {
        if (
          runtime.sessions.get(frameEditorId) !== session ||
          session.generation !== generation ||
          !session.bridgeReady
        ) {
          reject(
            new DOMException(
              `OnlyOffice cross-origin bridge changed: ${frameEditorId}`,
              "AbortError",
            ),
          );
          return;
        }

        const timer = targetWindow.setTimeout(() => {
          runtime.pendingRequests.delete(requestId);
          reject(
            new Error(`OnlyOffice cross-origin command timed out: ${command}`),
          );
        }, timeout);

        runtime.pendingRequests.set(requestId, {
          frameEditorId,
          generation,
          resolve,
          reject,
          timer,
        });
        postToIframe(session, {
          type: CROSS_ORIGIN_BRIDGE_MESSAGE.EDITOR_COMMAND,
          frameEditorId,
          requestId,
          command,
          payload,
        });
      }),
  );
}

export function subscribeCrossOriginEditorEvent(
  frameEditorId: string,
  event: string,
  handler: (args: unknown[]) => void,
  ownerWindow?: Window,
) {
  const runtime = getBridgeRuntimeState(
    resolveBridgeOwnerWindow(frameEditorId, ownerWindow),
  );
  let frameSubscribers = runtime.editorEventSubscribers.get(frameEditorId);
  if (!frameSubscribers) {
    frameSubscribers = new Map();
    runtime.editorEventSubscribers.set(frameEditorId, frameSubscribers);
  }

  let subscribers = frameSubscribers.get(event);
  if (!subscribers) {
    subscribers = new Set();
    frameSubscribers.set(event, subscribers);
  }

  subscribers.add(handler);

  return () => {
    subscribers?.delete(handler);
    if (subscribers?.size === 0) {
      frameSubscribers?.delete(event);
    }
    if (frameSubscribers?.size === 0) {
      runtime.editorEventSubscribers.delete(frameEditorId);
    }
  };
}

export function canAccessIframeWindow(
  iframe: HTMLIFrameElement | null | undefined,
) {
  if (!iframe) {
    return false;
  }
  try {
    void iframe.contentWindow?.location.href;
    return true;
  } catch {
    return false;
  }
}

export function watchCrossOriginIframe(
  frameEditorId: string,
  getIframe: () => HTMLIFrameElement | null | undefined,
  server: EditorServer,
  createIo: ScopedIoFactory,
  ownerWindow: Window = window,
  pluginConfigUrls: readonly string[] = [],
) {
  const ownerDocument = ownerWindow.document;
  const owner = {};
  let registered = false;
  let retryTimer: number | null = null;

  const stopRetry = () => {
    if (retryTimer === null) return;
    ownerWindow.clearInterval(retryTimer);
    retryTimer = null;
  };

  const tryRegister = () => {
    const iframe = getIframe();
    if (!iframe?.src || canAccessIframeWindow(iframe)) {
      return false;
    }
    if (registered) {
      return registerCrossOriginBridge(
        frameEditorId,
        iframe,
        server,
        createIo,
        owner,
        pluginConfigUrls,
      );
    }
    registered = registerCrossOriginBridge(
      frameEditorId,
      iframe,
      server,
      createIo,
      owner,
      pluginConfigUrls,
    );
    return registered;
  };

  const retryRegistration = () => {
    if (tryRegister()) {
      stopRetry();
      return;
    }
    if (retryTimer === null) {
      retryTimer = ownerWindow.setInterval(() => {
        if (tryRegister()) stopRetry();
      }, 10);
    }
  };

  retryRegistration();

  const MutationObserverConstructor = (
    ownerWindow as Window & { MutationObserver: typeof MutationObserver }
  ).MutationObserver;
  const observer = new MutationObserverConstructor(() => {
    retryRegistration();
  });
  observer.observe(ownerDocument.body, {
    attributes: true,
    attributeFilter: ["src"],
    childList: true,
    subtree: true,
  });

  return () => {
    stopRetry();
    observer.disconnect();
    if (registered) {
      unregisterCrossOriginBridge(frameEditorId, owner, ownerWindow);
    }
  };
}
