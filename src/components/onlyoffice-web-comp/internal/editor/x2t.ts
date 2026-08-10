/**
 * @description x2t 主线程代理，将重型文档转换转交给 Web Worker，避免阻塞界面线程。
 */

import { X2tConvertParams, X2tConvertResult } from "./types";
import {
  getStaticResource,
  resolveSiteUrl,
  type StaticResource,
} from "../../const";
import type { EditorLogger } from "./logger";

const DEFAULT_WORKER_READY_TIMEOUT_MS = 30_000;
const DEFAULT_REQUEST_TIMEOUT_MS = 5 * 60_000;

type X2tConversionErrorCode =
  | "worker-create"
  | "worker-ready-timeout"
  | "worker-error"
  | "worker-message-error"
  | "worker-post-message"
  | "worker-request-timeout"
  | "worker-response"
  | "worker-protocol"
  | "worker-terminated";

export class X2tConversionError extends Error {
  readonly code: X2tConversionErrorCode;
  readonly details?: Record<string, unknown>;

  constructor(
    message: string,
    options: {
      code: X2tConversionErrorCode;
      details?: Record<string, unknown>;
      cause?: unknown;
    },
  ) {
    super(message);
    this.name = "X2tConversionError";
    this.code = options.code;
    this.details = options.details;
    if (options.cause !== undefined) {
      this.cause = options.cause;
    }
  }
}

interface PendingMessage {
  expectedType: string;
  resolve: (value: any) => void;
  reject: (error: Error) => void;
  timeoutId: ReturnType<typeof setTimeout>;
}

interface WorkerResponse {
  id?: number;
  type: string;
  payload?: any;
  error?: string;
  errorName?: string;
  errorStack?: string;
  errorDetails?: unknown;
}

interface WorkerReadyState {
  worker: Worker;
  resolve: () => void;
  reject: (error: Error) => void;
  timeoutId: ReturnType<typeof setTimeout>;
}

export interface X2tConverterOptions {
  /** @internal 测试或定制运行时可替换 Worker 构造方式。 */
  workerFactory?: () => Worker;
  workerReadyTimeoutMs?: number;
  requestTimeoutMs?: number;
}

function toError(value: unknown, fallbackMessage: string): Error {
  if (value instanceof Error) return value;
  return new Error(value == null ? fallbackMessage : String(value));
}

function addTransferable(
  transferables: Set<ArrayBuffer>,
  value: ArrayBuffer | ArrayBufferView | undefined,
) {
  if (value instanceof ArrayBuffer) {
    transferables.add(value);
    return;
  }

  if (
    value &&
    ArrayBuffer.isView(value) &&
    value.buffer instanceof ArrayBuffer
  ) {
    transferables.add(value.buffer);
  }
}

function collectRequestTransferables(
  payload?: X2tConvertParams,
): Transferable[] {
  if (!payload) return [];

  const transferables = new Set<ArrayBuffer>();
  addTransferable(transferables, payload.data);
  addTransferable(transferables, payload.pdfBin);

  for (const fileMap of [payload.media, payload.fonts, payload.themes]) {
    Object.values(fileMap ?? {}).forEach((value) =>
      addTransferable(transferables, value),
    );
  }

  return Array.from(transferables);
}

export class X2tConverter {
  private worker: Worker | null = null;
  private initPromise: Promise<void> | null = null;
  private workerReadyState: WorkerReadyState | null = null;
  private workerIsReady = false;
  private requestQueue: Promise<void> = Promise.resolve();
  private requestGeneration = 0;
  private messageId = 0;
  private pendingMessages = new Map<number, PendingMessage>();
  private resourceKey = "";
  private logger?: EditorLogger;
  private readonly workerFactory: () => Worker;
  private readonly workerReadyTimeoutMs: number;
  private readonly requestTimeoutMs: number;

  constructor(options: X2tConverterOptions = {}) {
    this.workerFactory =
      options.workerFactory ??
      (() =>
        new Worker(new URL("./x2t.worker.ts", import.meta.url), {
          type: "module",
        }));
    this.workerReadyTimeoutMs =
      options.workerReadyTimeoutMs ?? DEFAULT_WORKER_READY_TIMEOUT_MS;
    this.requestTimeoutMs =
      options.requestTimeoutMs ?? DEFAULT_REQUEST_TIMEOUT_MS;
  }

  private getWorkerStaticResource(): StaticResource {
    const staticResource = getStaticResource();
    if (typeof window === "undefined") {
      return staticResource;
    }

    const origin = window.location.origin;
    return {
      ...staticResource,
      version: { ...staticResource.version },
      onlyoffice: { ...staticResource.onlyoffice },
      x2t: {
        ...staticResource.x2t,
        root: resolveSiteUrl(origin, staticResource.x2t.root),
        script: resolveSiteUrl(origin, staticResource.x2t.script),
        wasm: resolveSiteUrl(origin, staticResource.x2t.wasm),
        pdfFonts: {
          root: resolveSiteUrl(origin, staticResource.x2t.pdfFonts.root),
          default: resolveSiteUrl(origin, staticResource.x2t.pdfFonts.default),
        },
      },
    };
  }

  /**
   * @description 生成递增的 worker 消息 ID。
   */
  private getNextId(): number {
    return ++this.messageId;
  }

  private logRaw(
    level: "log" | "error",
    message: string,
    consoleArgs: unknown[],
  ) {
    if (this.logger) {
      this.logger.raw(level, "worker", message, consoleArgs);
      return;
    }
    console[level](...consoleArgs);
  }

  private rejectPendingMessages(error: Error) {
    for (const pending of this.pendingMessages.values()) {
      clearTimeout(pending.timeoutId);
      pending.reject(error);
    }
    this.pendingMessages.clear();
  }

  /**
   * @description 让损坏或超时的 Worker 完整退役；后续 convert 会创建干净实例重试。
   */
  private resetWorker(error: Error, worker = this.worker) {
    if (!worker || worker !== this.worker) return;

    const readyState = this.workerReadyState;
    if (readyState?.worker === worker) {
      clearTimeout(readyState.timeoutId);
      this.workerReadyState = null;
      readyState.reject(error);
    }

    this.rejectPendingMessages(error);
    worker.onmessage = null;
    worker.onerror = null;
    worker.onmessageerror = null;
    worker.terminate();

    this.worker = null;
    this.initPromise = null;
    this.workerIsReady = false;
    this.resourceKey = "";
  }

  /**
   * @description 向 worker 发送请求并等待对应响应。
   */
  private sendMessage<T>(type: string, payload?: X2tConvertParams): Promise<T> {
    return new Promise((resolve, reject) => {
      const worker = this.worker;
      if (!worker || !this.workerIsReady) {
        reject(
          new X2tConversionError("x2t worker is not ready", {
            code: "worker-protocol",
          }),
        );
        return;
      }

      const id = this.getNextId();
      const timeoutId = setTimeout(() => {
        if (!this.pendingMessages.has(id)) return;

        const timeoutError = new X2tConversionError(
          `x2t worker request timed out after ${this.requestTimeoutMs}ms`,
          {
            code: "worker-request-timeout",
            details: { id, type, timeoutMs: this.requestTimeoutMs },
          },
        );
        this.resetWorker(timeoutError, worker);
      }, this.requestTimeoutMs);

      this.pendingMessages.set(id, {
        expectedType: `${type}:done`,
        resolve,
        reject,
        timeoutId,
      });

      try {
        worker.postMessage(
          { id, type, payload },
          collectRequestTransferables(payload),
        );
      } catch (error) {
        clearTimeout(timeoutId);
        this.pendingMessages.delete(id);
        reject(
          new X2tConversionError("Failed to send x2t worker request", {
            code: "worker-post-message",
            cause: error,
          }),
        );
      }
    });
  }

  /**
   * @description 处理 worker 返回的响应消息。
   */
  private handleWorkerMessage(
    worker: Worker,
    event: MessageEvent<WorkerResponse>,
  ) {
    if (worker !== this.worker) return;

    const { id, type, payload, error, errorName, errorStack, errorDetails } =
      event.data;

    if (type === "ready") {
      const readyState = this.workerReadyState;
      if (!readyState || readyState.worker !== worker) return;

      clearTimeout(readyState.timeoutId);
      this.workerReadyState = null;
      this.workerIsReady = true;
      this.logRaw("log", "worker ready", ["[X2tConverter] Worker ready"]);
      readyState.resolve();
      return;
    }

    if (typeof id !== "number") {
      const protocolError = new X2tConversionError(
        "x2t worker response is missing a request id",
        {
          code: "worker-protocol",
          details: { responseType: type },
        },
      );
      this.logRaw("error", "worker protocol error", [
        "[X2tConverter] Worker response is missing a request id:",
        event.data,
      ]);
      this.resetWorker(protocolError, worker);
      return;
    }

    const pending = this.pendingMessages.get(id);
    if (!pending) return;

    clearTimeout(pending.timeoutId);
    this.pendingMessages.delete(id);

    if (type === "error") {
      const errorMessage = error || "Unknown worker error";
      const details = {
        ...(errorDetails && typeof errorDetails === "object"
          ? (errorDetails as Record<string, unknown>)
          : {}),
        workerErrorName: errorName,
        workerStack: errorStack,
      };
      if (this.logger) {
        this.logger.error("worker", "worker request failed", {
          message: errorMessage,
          ...details,
        });
      } else {
        console.error("[X2tConverter] Worker request failed:", {
          message: errorMessage,
          ...details,
        });
      }
      const responseError = new X2tConversionError(errorMessage, {
        code: "worker-response",
        details,
      });
      pending.reject(responseError);
      this.resetWorker(responseError, worker);
      return;
    }

    if (type !== pending.expectedType) {
      const protocolError = new X2tConversionError(
        `Unexpected x2t worker response: ${type || "(empty)"}`,
        {
          code: "worker-protocol",
          details: { id, expectedType: pending.expectedType, actualType: type },
        },
      );
      pending.reject(protocolError);
      this.resetWorker(protocolError, worker);
      return;
    }

    pending.resolve(payload);
  }

  /**
   * @description 处理 worker 运行错误，并让所有等待中的请求失败。
   */
  private handleWorkerError(worker: Worker, error: ErrorEvent) {
    const workerError = new X2tConversionError(
      `x2t worker failed: ${error.message || "Unknown worker error"}`,
      {
        code: "worker-error",
        cause: error.error,
        details: {
          filename: error.filename,
          lineno: error.lineno,
          colno: error.colno,
        },
      },
    );
    this.logRaw("error", "worker error", [
      "[X2tConverter] Worker error:",
      error,
    ]);
    this.resetWorker(workerError, worker);
  }

  private handleWorkerMessageError(worker: Worker, event: MessageEvent) {
    const error = new X2tConversionError(
      "x2t worker returned an unreadable message",
      {
        code: "worker-message-error",
        details: { data: event.data },
      },
    );
    this.logRaw("error", "worker message error", [
      "[X2tConverter] Worker message error:",
      event,
    ]);
    this.resetWorker(error, worker);
  }

  /**
   * @description 初始化 x2t worker；只有收到 ready 握手后才视为可用。
   */
  public init(logger?: EditorLogger): Promise<void> {
    if (logger) {
      this.logger = logger;
    }
    if (this.initPromise) {
      return this.initPromise;
    }

    let worker: Worker;
    try {
      worker = this.workerFactory();
    } catch (error) {
      return Promise.reject(
        new X2tConversionError("Failed to create x2t worker", {
          code: "worker-create",
          cause: toError(error, "Unknown worker creation error"),
        }),
      );
    }

    this.worker = worker;
    this.workerIsReady = false;
    this.initPromise = new Promise<void>((resolve, reject) => {
      const timeoutId = setTimeout(() => {
        const timeoutError = new X2tConversionError(
          `x2t worker did not become ready within ${this.workerReadyTimeoutMs}ms`,
          {
            code: "worker-ready-timeout",
            details: { timeoutMs: this.workerReadyTimeoutMs },
          },
        );
        this.resetWorker(timeoutError, worker);
      }, this.workerReadyTimeoutMs);

      this.workerReadyState = { worker, resolve, reject, timeoutId };
      worker.onmessage = (event) => this.handleWorkerMessage(worker, event);
      worker.onerror = (event) => this.handleWorkerError(worker, event);
      worker.onmessageerror = (event) =>
        this.handleWorkerMessageError(worker, event);
    });

    this.logRaw("log", "worker created", ["[X2tConverter] Worker created"]);
    return this.initPromise;
  }

  /**
   * @description 将文档从一种格式转换为另一种格式。
   */
  public async convert(
    params: X2tConvertParams,
    logger?: EditorLogger,
  ): Promise<X2tConvertResult> {
    const requestGeneration = this.requestGeneration;
    const request = this.requestQueue.then(() => {
      if (requestGeneration !== this.requestGeneration) {
        throw new X2tConversionError("x2t request was cancelled", {
          code: "worker-terminated",
        });
      }
      return this.convertNow(params, logger);
    });
    this.requestQueue = request.then(
      () => undefined,
      () => undefined,
    );
    return request;
  }

  private async convertNow(
    {
      data,
      fileFrom,
      fileTo,
      formatFrom,
      formatTo,
      media,
      pdfBin,
      fonts,
      fontAliases,
      fontExportAliases,
      themes,
      csvEncoding,
      csvDelimiter,
      csvDelimiterChar,
    }: X2tConvertParams,
    logger?: EditorLogger,
  ): Promise<X2tConvertResult> {
    if (logger) {
      this.logger = logger;
    }
    const staticResource = this.getWorkerStaticResource();
    const resourceKey = JSON.stringify(staticResource.x2t);
    if (this.worker && this.resourceKey && this.resourceKey !== resourceKey) {
      this.resetWorker(
        new X2tConversionError("x2t static resource changed", {
          code: "worker-terminated",
        }),
      );
    }
    this.resourceKey = resourceKey;

    await this.init(logger);

    const cloneMap = (map?: { [key: string]: Uint8Array }) => {
      if (!map) return undefined;
      return Object.fromEntries(
        Object.entries(map).map(([key, value]) => [key, value.slice(0)]),
      );
    };

    /**
     * @description 发送给 worker 前复制数据，转移全部副本而不分离调用方持有的二进制。
     */
    const payload: X2tConvertParams = {
      data: data.slice(0),
      fileFrom,
      fileTo,
      formatFrom,
      formatTo,
      media: cloneMap(media),
      pdfBin: pdfBin?.slice(0),
      fonts: cloneMap(fonts),
      fontAliases,
      fontExportAliases,
      themes: cloneMap(themes),
      csvEncoding,
      csvDelimiter,
      csvDelimiterChar,
      staticResource,
    };
    this.logger?.worker("convert", {
      fileFrom,
      fileTo,
      formatFrom,
      formatTo,
    });
    return this.sendMessage<X2tConvertResult>("convert", payload);
  }

  /**
   * @description 终止 worker 并释放关联资源。
   */
  public terminate(logger?: EditorLogger): void {
    if (logger) {
      this.logger = logger;
    }
    this.requestGeneration += 1;

    const worker = this.worker;
    if (!worker) {
      this.initPromise = null;
      this.workerIsReady = false;
      this.resourceKey = "";
      return;
    }

    this.resetWorker(
      new X2tConversionError("x2t worker was terminated", {
        code: "worker-terminated",
      }),
      worker,
    );
    this.logRaw("log", "worker terminated", [
      "[X2tConverter] Worker terminated",
    ]);
  }

  /**
   * @description 判断 worker 是否已完成 ready 握手。
   */
  public get isInitialized(): boolean {
    return this.worker !== null && this.workerIsReady;
  }
}

/** @description 浏览器端 Office 格式转换器。 */
export const converter = new X2tConverter();
