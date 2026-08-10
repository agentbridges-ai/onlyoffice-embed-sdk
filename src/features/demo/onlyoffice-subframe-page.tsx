"use client";

import { useEffect, useRef } from "react";
import {
  ONLYOFFICE_EVENT_KEYS,
  OnlyOfficeManager,
  onlyOfficeManagerFactory,
  onlyofficeEventbus,
  type OnlyOfficeConnector,
  type OfficeTheme,
} from "@/components/onlyoffice-web-comp";
import type { CommentItem, RevisionItem } from "@/components/onlyoffice-web-comp";
import {
  SUBFRAME_EDITOR_CONTAINER_ID,
  SUBFRAME_MESSAGE_SOURCE,
  type SubframeConnectorOperation,
  type SubframeEditorAction,
  type SubframeEventMessage,
  type SubframeEventName,
  type SubframeMessage,
  type SubframeOpenPayload,
  type SubframeRequest,
} from "./subframe-protocol";

type ConnectorPayload = {
  operation: SubframeConnectorOperation;
  command?: string;
  method?: string;
  args?: unknown[];
  recalculate?: boolean;
};

function getInstanceId() {
  if (typeof window === "undefined") return "";
  return new URLSearchParams(window.location.search).get("instance") || "subframe";
}

function getParentOrigin() {
  if (typeof document === "undefined") return "*";
  try {
    return document.referrer ? new URL(document.referrer).origin : "*";
  } catch {
    return "*";
  }
}

function toSerializable(value: unknown, seen = new WeakSet<object>()): unknown {
  if (value === null || value === undefined) return value;
  if (
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  ) {
    return value;
  }
  if (typeof value === "bigint") return String(value);
  if (typeof value === "function" || typeof value === "symbol") return undefined;
  if (value instanceof Date) return value.toISOString();
  if (value instanceof Blob || value instanceof ArrayBuffer) return value;
  if (value instanceof Uint8Array) return Array.from(value);
  if (typeof value !== "object") return String(value);
  if (seen.has(value)) return undefined;
  seen.add(value);

  if (Array.isArray(value)) {
    return value
      .map((item) => toSerializable(item, seen))
      .filter((item) => item !== undefined);
  }

  const result: Record<string, unknown> = {};
  for (const key of Object.keys(value)) {
    const item = toSerializable(
      (value as Record<string, unknown>)[key],
      seen,
    );
    if (item !== undefined) result[key] = item;
  }
  return result;
}

function serializeComments(items: CommentItem[]) {
  return items.map((item) => ({
    Id: String(item.Id),
    Data: (toSerializable(item.Data) || {}) as Record<string, unknown>,
  }));
}

function serializeRevisions(items: RevisionItem[]) {
  // `Raw` contains SDK instances and methods, so it must never cross the
  // structured-clone boundary. The parent only needs Id/Index/Data for the
  // demo; mutation commands send the stable Id back to this subframe.
  return items.map((item) => ({
    Id: String(item.Id),
    Index: item.Index,
    Data: (toSerializable(item.Data) || {}) as Record<string, unknown>,
  }));
}

function isRequest(value: unknown): value is SubframeRequest {
  return (
    !!value &&
    typeof value === "object" &&
    (value as { source?: unknown }).source === SUBFRAME_MESSAGE_SOURCE &&
    (value as { type?: unknown }).type === "request"
  );
}

function sendMessage(
  parentOrigin: string,
  message: SubframeMessage | SubframeEventMessage,
) {
  if (typeof window === "undefined" || window.parent === window) return;
  window.parent.postMessage(message, parentOrigin === "*" ? "*" : parentOrigin);
}

export function OnlyOfficeSubframePage() {
  const managerRef = useRef<Awaited<
    ReturnType<typeof onlyOfficeManagerFactory.open>
  > | null>(null);
  const connectorRef = useRef<OnlyOfficeConnector | null>(null);
  const instanceIdRef = useRef("");
  const parentOriginRef = useRef("*");
  const operationChainRef = useRef(Promise.resolve());

  useEffect(() => {
    const instanceId = getInstanceId();
    instanceIdRef.current = instanceId;
    parentOriginRef.current = getParentOrigin();

    const postEvent = (event: SubframeEventName, payload: unknown) => {
      sendMessage(parentOriginRef.current, {
        source: SUBFRAME_MESSAGE_SOURCE,
        type: "event",
        instanceId,
        event,
        payload,
      });
    };

    const onLoadingChange = (payload: { loading: boolean }) => {
      postEvent("loading-change", payload);
    };
    const onDocumentReady = (payload: unknown) => {
      postEvent("document-ready", payload);
    };
    const onOfficeXmlLimit = (payload: unknown) => {
      postEvent("office-xml-size-limit-exceeded", payload);
    };

    onlyofficeEventbus.on(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, onLoadingChange);
    onlyofficeEventbus.on(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, onDocumentReady);
    onlyofficeEventbus.on(
      ONLYOFFICE_EVENT_KEYS.OFFICE_XML_SIZE_LIMIT_EXCEEDED,
      onOfficeXmlLimit,
    );

    const configureStaticResource = () => {
      const search = new URLSearchParams(window.location.search);
      const mode = search.get("resourceMode");
      const cdnOrigin = search.get("cdnOrigin")?.trim();
      if (mode === "cdn" && cdnOrigin) {
        OnlyOfficeManager.registerStaticResource({ cdnOrigin });
      } else {
        OnlyOfficeManager.resetStaticResource();
      }
    };

    const getManager = () => {
      const manager = managerRef.current;
      if (!manager) throw new Error("OnlyOffice subframe editor is not open");
      return manager;
    };

    const getConnector = () => {
      const manager = getManager();
      if (!connectorRef.current) {
        connectorRef.current = manager.getEditor().createConnector({
          autoconnect: false,
        });
      }
      return connectorRef.current;
    };

    const runEditorAction = async (
      action: SubframeEditorAction,
      payload: unknown,
    ) => {
      const editor = getManager().getEditor();
      const data = (payload || {}) as Record<string, unknown>;

      switch (action) {
        case "get-all-comments":
          return serializeComments(await editor.getAllComments());
        case "add-comment":
          return await editor.addComment(data.input as string);
        case "update-comment":
          await editor.updateComment(
            String(data.id),
            data.data as Record<string, unknown>,
          );
          return null;
        case "remove-comment":
          await editor.removeComment(String(data.id));
          return null;
        case "get-all-revisions":
          return serializeRevisions(await editor.getAllRevisions());
        case "set-track-revisions":
          await editor.setTrackRevisions(!!data.enabled);
          return null;
        case "add-demo-revision":
          return serializeRevisions(
            await editor.addDemoRevision(
              typeof data.text === "string" ? data.text : undefined,
            ),
          );
        case "go-to-revision":
          await editor.goToRevision(String(data.id));
          return null;
        case "accept-revision":
          await editor.acceptRevision(String(data.id));
          return null;
        case "reject-revision":
          await editor.rejectRevision(String(data.id));
          return null;
        case "accept-all-revisions":
          await editor.acceptAllRevisions();
          return null;
        case "reject-all-revisions":
          await editor.rejectAllRevisions();
          return null;
      }
    };

    const runConnectorAction = async (payload: ConnectorPayload) => {
      const connector = getConnector();
      switch (payload.operation) {
        case "connect":
          if (!connector.isConnected) connector.connect();
          return null;
        case "disconnect":
          if (connector.isConnected) connector.disconnect();
          return null;
        case "call-command":
          return await new Promise<unknown>((resolve) => {
            connector.callCommand(
              payload.command || "",
              (result) => resolve(toSerializable(result)),
              payload.recalculate,
            );
          });
        case "execute-method":
          return await new Promise<unknown>((resolve) => {
            connector.executeMethod(
              payload.method || "",
              payload.args || [],
              (result) => resolve(toSerializable(result)),
            );
          });
      }
    };

    const openEditor = async (payload: SubframeOpenPayload) => {
      configureStaticResource();
      const manager = await onlyOfficeManagerFactory.open(
        {
          containerId: SUBFRAME_EDITOR_CONTAINER_ID,
          fileType: payload.fileType,
          defaultFileName: payload.defaultFileName,
          readOnly: payload.readOnly,
          theme: payload.theme,
          officeXmlEvent: payload.officeXmlEvent,
        },
        payload.document,
      );
      managerRef.current = manager;
      connectorRef.current = null;
      return {
        fileName: payload.document.fileName,
        fileType: payload.fileType,
        readOnly: manager.getReadOnly(),
      };
    };

    const runAction = async (request: SubframeRequest) => {
      switch (request.action) {
        case "open":
          return await openEditor(request.payload as SubframeOpenPayload);
        case "set-read-only": {
          const payload = (request.payload || {}) as { readOnly?: boolean };
          await getManager().setReadOnly(!!payload.readOnly);
          return null;
        }
        case "set-theme": {
          const payload = (request.payload || {}) as { theme?: OfficeTheme };
          await getManager().setTheme(payload.theme || "theme-white");
          return null;
        }
        case "toggle-language":
          return await getManager().toggleLanguage();
        case "download":
          return await getManager().exportAsBlob();
        case "print-logs":
          getManager().printLogs();
          return null;
        case "editor": {
          const payload = (request.payload || {}) as {
            action?: SubframeEditorAction;
            payload?: unknown;
          };
          if (!payload.action) throw new Error("Missing subframe editor action");
          return await runEditorAction(payload.action, payload.payload);
        }
        case "connector":
          return await runConnectorAction(request.payload as ConnectorPayload);
        case "destroy":
          onlyOfficeManagerFactory.destroy(SUBFRAME_EDITOR_CONTAINER_ID);
          managerRef.current = null;
          connectorRef.current = null;
          return null;
      }
    };

    const processRequest = (request: SubframeRequest) => {
      const operation = operationChainRef.current.then(
        () => runAction(request),
        () => runAction(request),
      );
      operationChainRef.current = operation.then(
        () => undefined,
        () => undefined,
      );

      void operation.then(
        (result) => {
          sendMessage(parentOriginRef.current, {
            source: SUBFRAME_MESSAGE_SOURCE,
            type: "response",
            instanceId,
            requestId: request.requestId,
            ok: true,
            result: toSerializable(result),
          });
        },
        (error) => {
          sendMessage(parentOriginRef.current, {
            source: SUBFRAME_MESSAGE_SOURCE,
            type: "response",
            instanceId,
            requestId: request.requestId,
            ok: false,
            error: error instanceof Error ? error.message : String(error),
          });
        },
      );
    };

    const onMessage = (event: MessageEvent) => {
      if (event.source !== window.parent || !isRequest(event.data)) return;
      if (
        parentOriginRef.current !== "*" &&
        event.origin !== parentOriginRef.current
      ) {
        return;
      }
      if (parentOriginRef.current === "*") {
        parentOriginRef.current = event.origin;
      }

      const request = event.data;
      if (request.instanceId !== instanceId) return;
      processRequest(request);
    };

    window.addEventListener("message", onMessage);
    sendMessage(parentOriginRef.current, {
      source: SUBFRAME_MESSAGE_SOURCE,
      type: "ready",
      instanceId,
    });

    return () => {
      window.removeEventListener("message", onMessage);
      onlyofficeEventbus.off(
        ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE,
        onLoadingChange,
      );
      onlyofficeEventbus.off(
        ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY,
        onDocumentReady,
      );
      onlyofficeEventbus.off(
        ONLYOFFICE_EVENT_KEYS.OFFICE_XML_SIZE_LIMIT_EXCEEDED,
        onOfficeXmlLimit,
      );
      onlyOfficeManagerFactory.destroy(SUBFRAME_EDITOR_CONTAINER_ID);
    };
  }, []);

  return (
    <main className="h-screen w-screen overflow-hidden bg-white">
      <div className="onlyoffice-container relative h-full w-full">
        <div id={SUBFRAME_EDITOR_CONTAINER_ID} className="absolute inset-0" />
      </div>
    </main>
  );
}
