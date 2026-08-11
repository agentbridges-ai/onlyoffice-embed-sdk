const OFFICE_PLUGIN_PROTOCOL = "onlyoffice-browser-plugin/v1" as const;
const PLUGIN_REQUEST_TIMEOUT_MS = 45_000;

type OfficePluginWindowMessage = {
  protocol: typeof OFFICE_PLUGIN_PROTOCOL;
  type: "READY" | "RESULT";
  pluginGuid: string;
  pluginInstanceId: string;
  editorType?: string;
  requestId?: string;
  ok?: boolean;
  result?: unknown;
  error?: string;
};

type OfficePluginRuntime = {
  pluginGuid: string;
  pluginInstanceId: string;
  origin: string;
  source: WindowProxy;
};

type PendingPluginRequest = {
  runtime: OfficePluginRuntime;
  resolve: (result: unknown) => void;
  reject: (error: Error) => void;
  timeoutId: number;
};

export type OfficePluginBridgeOptions = {
  pluginGuids: readonly string[];
  /** Exact origins discovered from each configured plugin manifest. */
  pluginOrigins?: ReadonlyMap<
    string,
    ReadonlySet<string> | readonly string[]
  >;
  /** Proves that a READY source belongs to this editor instance. */
  isAllowedSource: (source: WindowProxy) => boolean;
  onReady: (pluginGuid: string, editorType: string) => void;
};

function normalizePluginOrigin(value: string) {
  try {
    const url = new URL(value);
    if (url.protocol !== "http:" && url.protocol !== "https:") return null;
    const origin = url.origin;
    return origin === "null" ? null : origin;
  } catch {
    return null;
  }
}

function getOwnerOrigin(ownerWindow: Window) {
  const inheritedOrigin = normalizePluginOrigin(
    (ownerWindow as Window & { origin?: string }).origin || "",
  );
  return inheritedOrigin ?? normalizePluginOrigin(ownerWindow.location.href);
}

/**
 * Direct-embed equivalent of onlyoffice-browser's host-side plugin bridge.
 *
 * READY is accepted only from an exact origin associated with the configured
 * GUID. RESULT additionally has to come from the exact origin, WindowProxy and
 * pluginInstanceId registered by READY, preventing another configured plugin
 * or sibling frame from completing an in-flight request.
 */
export class OfficePluginBridge {
  private readonly ownerWindow: Window;
  private readonly configuredPluginGuids: Set<string>;
  private readonly configuredPluginOrigins = new Map<string, Set<string>>();
  private readonly onReady: OfficePluginBridgeOptions["onReady"];
  private readonly isAllowedSource: OfficePluginBridgeOptions["isAllowedSource"];
  private readonly runtimes = new Map<string, OfficePluginRuntime>();
  private readonly pending = new Map<string, PendingPluginRequest>();
  private requestSequence = 0;
  private destroyed = false;

  constructor(ownerWindow: Window, options: OfficePluginBridgeOptions) {
    this.ownerWindow = ownerWindow;
    this.configuredPluginGuids = new Set(options.pluginGuids);
    const ownerOrigin = getOwnerOrigin(ownerWindow);
    for (const [pluginGuid, origins] of options.pluginOrigins ?? []) {
      this.configuredPluginGuids.add(pluginGuid);
      let configuredOrigins = this.configuredPluginOrigins.get(pluginGuid);
      if (!configuredOrigins) {
        configuredOrigins = new Set();
        this.configuredPluginOrigins.set(pluginGuid, configuredOrigins);
      }
      for (const value of origins) {
        const origin = normalizePluginOrigin(value);
        if (origin) configuredOrigins.add(origin);
      }
    }
    for (const pluginGuid of options.pluginGuids) {
      if (!ownerOrigin || this.configuredPluginOrigins.has(pluginGuid)) {
        continue;
      }
      this.configuredPluginOrigins.set(pluginGuid, new Set([ownerOrigin]));
    }
    this.onReady = options.onReady;
    this.isAllowedSource = options.isAllowedSource;
    ownerWindow.addEventListener("message", this.handleMessage);
  }

  invoke(pluginGuid: string, payload: unknown): Promise<unknown> {
    if (this.destroyed) {
      return Promise.reject(new Error("Editor is not open"));
    }
    if (!pluginGuid.trim()) {
      return Promise.reject(new Error("A plugin GUID is required"));
    }

    const runtime = this.runtimes.get(pluginGuid);
    if (!runtime) {
      return Promise.reject(
        new Error(`Office plugin is not ready: ${pluginGuid}`),
      );
    }

    const requestId = `${pluginGuid}-request-${Date.now()}-${++this.requestSequence}`;
    return new Promise((resolve, reject) => {
      const timeoutId = this.ownerWindow.setTimeout(() => {
        this.pending.delete(requestId);
        reject(new Error(`Office plugin operation timed out: ${pluginGuid}`));
      }, PLUGIN_REQUEST_TIMEOUT_MS);
      this.pending.set(requestId, {
        runtime,
        resolve,
        reject,
        timeoutId,
      });
      runtime.source.postMessage(
        {
          protocol: OFFICE_PLUGIN_PROTOCOL,
          type: "INVOKE",
          pluginGuid,
          pluginInstanceId: runtime.pluginInstanceId,
          requestId,
          payload,
        },
        runtime.origin,
      );
    });
  }

  destroy() {
    if (this.destroyed) return;
    this.destroyed = true;
    this.ownerWindow.removeEventListener("message", this.handleMessage);
    for (const request of this.pending.values()) {
      this.ownerWindow.clearTimeout(request.timeoutId);
      request.reject(
        new Error("Editor was destroyed before plugin operation completed"),
      );
    }
    this.pending.clear();
    this.runtimes.clear();
  }

  private readonly handleMessage = (
    event: MessageEvent<OfficePluginWindowMessage>,
  ) => {
    const message = event.data;
    if (
      this.destroyed ||
      !message ||
      message.protocol !== OFFICE_PLUGIN_PROTOCOL ||
      !this.configuredPluginGuids.has(message.pluginGuid) ||
      !this.configuredPluginOrigins
        .get(message.pluginGuid)
        ?.has(event.origin) ||
      typeof message.pluginInstanceId !== "string" ||
      !message.pluginInstanceId ||
      !event.source ||
      !this.isAllowedSource(event.source as WindowProxy)
    ) {
      return;
    }

    if (message.type === "READY") {
      const incoming: OfficePluginRuntime = {
        pluginGuid: message.pluginGuid,
        pluginInstanceId: message.pluginInstanceId,
        origin: event.origin,
        source: event.source as WindowProxy,
      };
      const current = this.runtimes.get(message.pluginGuid);
      if (current?.pluginInstanceId === incoming.pluginInstanceId) {
        // A duplicate from the registered source is harmless; the same
        // instance id from any other source is a spoof and is ignored.
        if (
          current.source !== incoming.source ||
          current.origin !== incoming.origin
        ) {
          return;
        }
        return;
      }

      if (current) {
        this.rejectPendingForPlugin(
          message.pluginGuid,
          `Office plugin reloaded before operation completed: ${message.pluginGuid}`,
        );
      }
      this.runtimes.set(message.pluginGuid, incoming);
      this.onReady(message.pluginGuid, message.editorType || "");
      return;
    }

    if (message.type !== "RESULT" || !message.requestId) return;
    const request = this.pending.get(message.requestId);
    if (
      !request ||
      request.runtime.pluginGuid !== message.pluginGuid ||
      request.runtime.pluginInstanceId !== message.pluginInstanceId ||
      request.runtime.origin !== event.origin ||
      request.runtime.source !== event.source
    ) {
      return;
    }

    this.pending.delete(message.requestId);
    this.ownerWindow.clearTimeout(request.timeoutId);
    if (message.ok === true) request.resolve(message.result);
    else request.reject(new Error(message.error || "Office plugin operation failed"));
  };

  private rejectPendingForPlugin(pluginGuid: string, reason: string) {
    for (const [requestId, request] of this.pending) {
      if (request.runtime.pluginGuid !== pluginGuid) continue;
      this.pending.delete(requestId);
      this.ownerWindow.clearTimeout(request.timeoutId);
      request.reject(new Error(reason));
    }
  }
}
