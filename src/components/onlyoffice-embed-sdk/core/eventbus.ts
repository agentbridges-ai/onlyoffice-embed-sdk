import { EventEmitter, type EventListener } from "../util/event-emitter";
import { ONLYOFFICE_EVENT_KEYS } from "../const";
import type { OfficeXmlSizeLimitExceededPayload } from "../internal/editor/types";

export type DocumentReadyData = {
  fileName: string;
  fileType: string;
  instanceId?: string;
  /** Internal lifecycle owner; disambiguates equal container ids in other Documents. */
  manager?: object;
};

export type SaveDocumentData = {
  fileName: string;
  fileType: string;
  binData: Uint8Array;
  instanceId: string;
  media?: Record<string, Uint8Array>;
  themes?: Record<string, Uint8Array>;
};

export type OnSaveData = {
  fileName: string;
  instanceId: string;
};

export type OfficeResourceLoadStatus =
  | "checking"
  | "cache-hit"
  | "downloading"
  | "downloaded"
  | "not-observed";

export type OfficeLoadingPhase =
  | "host-loading"
  | "runtime-loading"
  | "static-resources"
  | "operation"
  | "ready"
  | "error"
  | "destroyed";

export type LoadingChangeData = {
  loading: boolean;
  phase: OfficeLoadingPhase;
  resourceStatus: OfficeResourceLoadStatus;
  /** True only after the browser reports real transferred bytes. */
  resourceDownload: boolean;
  /** Encoded body bytes transferred from the network, excluding cache hits. */
  transferredBytes: number;
  /** Canonical static resources observed from cache or network. */
  resourceCount: number;
  instanceId?: string;
  manager?: object;
};

export type OfficeXmlSizeLimitExceededData =
  OfficeXmlSizeLimitExceededPayload & {
    instanceId: string;
    containerId: string;
  };

type OnlyOfficeEventPayloads = {
  [ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY]: DocumentReadyData;
  [ONLYOFFICE_EVENT_KEYS.SAVE_DOCUMENT]: SaveDocumentData;
  [ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE]: LoadingChangeData;
  [ONLYOFFICE_EVENT_KEYS.ONSAVE]: OnSaveData;
  [ONLYOFFICE_EVENT_KEYS.OFFICE_XML_SIZE_LIMIT_EXCEEDED]: OfficeXmlSizeLimitExceededData;
};

class OnlyOfficeEventBus {
  private emitter = new EventEmitter();

  on<Key extends keyof OnlyOfficeEventPayloads>(
    key: Key,
    handler: (data: OnlyOfficeEventPayloads[Key]) => void,
  ) {
    this.emitter.on(key, handler as EventListener);
  }

  off<Key extends keyof OnlyOfficeEventPayloads>(
    key: Key,
    handler: (data: OnlyOfficeEventPayloads[Key]) => void,
  ) {
    this.emitter.off(key, handler as EventListener);
  }

  emit<Key extends keyof OnlyOfficeEventPayloads>(
    key: Key,
    data: OnlyOfficeEventPayloads[Key],
  ) {
    this.emitter.emit(key, data);
  }

  waitFor<Key extends keyof OnlyOfficeEventPayloads>(
    key: Key,
    timeout = 30000,
  ): Promise<OnlyOfficeEventPayloads[Key]> {
    return new Promise((resolve, reject) => {
      const timer = window.setTimeout(() => {
        this.off(key, handler);
        reject(new Error(`Timed out waiting for ${key}`));
      }, timeout);

      const handler = (data: OnlyOfficeEventPayloads[Key]) => {
        window.clearTimeout(timer);
        this.off(key, handler);
        resolve(data);
      };

      this.on(key, handler);
    });
  }
}

export const onlyofficeEventbus = new OnlyOfficeEventBus();
