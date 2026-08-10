"use client";

import { useEffect, useState } from "react";
import {
  DEFAULT_OFFICE_THEME,
  FILE_TYPE,
  ONLYOFFICE_CONTAINER_CONFIG,
  ONLYOFFICE_EVENT_KEYS,
  OFFICE_THEME,
  STATIC_RESOURCE,
  OnlyOfficeManager,
  editorManagerFactory,
  isOnlyOfficeCdnMode,
  onlyOfficeManagerFactory,
  onlyofficeEventbus,
} from "@/components/onlyoffice-web-comp";
import { mountOfficeEditor } from "@/components/onlyoffice-web-comp/compat/editor";
import { converter } from "@/components/onlyoffice-web-comp/internal/editor/x2t";
import { getScopedIoRegistry } from "@/components/onlyoffice-web-comp/internal/editor/runtime-bridge";
import { getX2tConvertFormats } from "@/components/onlyoffice-web-comp/internal/editor/utils";
import type { FileType, OfficeTheme } from "@/components/onlyoffice-web-comp";
import type {
  ResourceMode,
  ScenarioResult,
  StepResult,
} from "./onlyoffice-factory.contract";

export const CONTAINER_IDS = {
  plugin: "e2e-plugin-editor",
  factory: "e2e-factory-editor",
  concurrent: "e2e-concurrent-editor",
  create: "e2e-create-editor",
  file: "e2e-file-editor",
  textFallback: "e2e-text-fallback-editor",
  fromEditor: "e2e-from-editor",
  fixture: "e2e-fixture-editor",
} as const;

const DOCUMENT_READY_TIMEOUT_MS = 30_000;
const NEXOLYRA_PLUGIN_GUID =
  "asc.{E2E4D0B6-6F1E-4B80-9A4D-8F6B1C2D3E40}";

declare const Api: {
  GetActiveSheet: () => {
    GetRange: (address: string) => {
      SetValue: (value: string) => void;
    };
  };
  CreateRGBColor: (red: number, green: number, blue: number) => unknown;
  CreateSolidFill: (color: unknown) => unknown;
  CreateStroke: (width: number, fill: unknown) => unknown;
  CreateShape: (
    type: string,
    width: number,
    height: number,
    fill: unknown,
    stroke: unknown,
  ) => {
    SetPosition: (x: number, y: number) => void;
    GetDocContent: () => {
      GetContent: () => Array<{ AddText: (text: string) => void }>;
    };
  };
  GetPresentation: () => {
    GetCurrentSlide: () => {
      AddObject: (shape: unknown) => void;
    };
  };
};

function waitForDocumentReady(instanceId?: string) {
  let timeoutId: number | undefined;
  let handler: ((data: { instanceId?: string }) => void) | undefined;

  const promise = new Promise<void>((resolve, reject) => {
    handler = (data) => {
      if (instanceId && data.instanceId !== instanceId) return;
      if (timeoutId !== undefined) {
        window.clearTimeout(timeoutId);
      }
      if (handler) {
        onlyofficeEventbus.off(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, handler);
      }
      resolve();
    };

    timeoutId = window.setTimeout(() => {
      onlyofficeEventbus.off(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, handler);
      reject(new Error("Timed out waiting for OnlyOffice documentReady"));
    }, DOCUMENT_READY_TIMEOUT_MS);

    onlyofficeEventbus.on(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, handler);
  });

  return {
    promise,
    cancel() {
      if (timeoutId !== undefined) {
        window.clearTimeout(timeoutId);
      }
      if (handler) {
        onlyofficeEventbus.off(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, handler);
      }
    },
  };
}

async function withDocumentReady<T>(
  action: () => Promise<T>,
  operation = "OnlyOffice operation",
  instanceId?: string,
) {
  const ready = waitForDocumentReady(instanceId);
  try {
    const value = await action();
    await ready.promise.catch((error) => {
      throw new Error(
        `${operation} did not emit documentReady`,
        error instanceof Error ? { cause: error } : undefined,
      );
    });
    return value;
  } catch (error) {
    ready.cancel();
    throw error;
  }
}

function assert(condition: unknown, message: string): asserts condition {
  if (!condition) {
    throw new Error(message);
  }
}

async function fetchPublicFile(url: string, fileName: string) {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Failed to fetch ${url}`);
  }
  const blob = await response.blob();
  return new File([blob], fileName, { type: blob.type });
}

function frameOrigin(containerId: string) {
  const frames = Array.from(
    document.querySelectorAll<HTMLIFrameElement>('iframe[name="frameEditor"]'),
  );
  const frame =
    frames.find((item) => {
      try {
        return (
          new URL(item.src, window.location.href).searchParams.get(
            "frameEditorId",
          ) === containerId
        );
      } catch {
        return false;
      }
    }) ?? frames[0];

  assert(frame?.src, `Missing frameEditor iframe for ${containerId}`);
  return new URL(frame.src, window.location.href).origin;
}

async function assertExport(
  manager: OnlyOfficeManager,
  expectedFileType: FileType,
) {
  const data = await manager.exportDocument();
  assert(
    data.fileType === expectedFileType.toLowerCase(),
    "Unexpected export type",
  );
  assert(data.binData.byteLength > 0, "Expected exported bin data");
}

async function assertX2tImport(fileName: string, fileType: FileType) {
  const file = await fetchPublicFile(`/e2e/fixtures/${fileName}`, fileName);
  const data = await file.arrayBuffer();
  const { formatFrom, formatTo } = getX2tConvertFormats(fileType);
  const result = await converter.convert({
    data,
    fileFrom: `doc.${fileType.toLowerCase()}`,
    fileTo: "Editor.bin",
    formatFrom,
    formatTo,
  });

  assert(
    result.output && result.output.byteLength > 0,
    `Expected x2t to import ${fileName}`,
  );
}

async function assertX2tRejectsTraversalAndRecovers() {
  const fileName = "edge-invalid-bookmark.docx";
  const file = await fetchPublicFile(`/e2e/fixtures/${fileName}`, fileName);
  const data = await file.arrayBuffer();
  const { formatFrom, formatTo } = getX2tConvertFormats(FILE_TYPE.DOCX);
  const params = {
    data,
    fileFrom: "doc.docx",
    fileTo: "Editor.bin",
    formatFrom,
    formatTo,
  };

  const unsafeResult = await converter
    .convert({
      ...params,
      media: { "../params.xml": Uint8Array.from([1]) },
    })
    .then(
      () => null,
      (error) => error,
    );
  assert(unsafeResult instanceof Error, "x2t accepted a traversal media path");

  const recovered = await converter.convert(params);
  assert(
    recovered.output?.byteLength,
    "x2t did not recover with a clean Worker after rejecting traversal",
  );

  const isolated = await converter.convert({
    ...params,
    media: {
      "params.xml": Uint8Array.from([11]),
      "media/nested/sentinel.bin": Uint8Array.from([12]),
    },
    themes: {
      "doc.docx": Uint8Array.from([13]),
      "themes/nested/sentinel.xml": new TextEncoder().encode("<theme />"),
    },
  });
  assert(
    isolated.media["params.xml"]?.[0] === 11 &&
      isolated.media["nested/sentinel.bin"]?.[0] === 12,
    "x2t did not keep media inputs inside the media root",
  );
  assert(
    isolated.themes?.["themes/doc.docx"]?.[0] === 13 &&
      isolated.themes?.["themes/nested/sentinel.xml"],
    "x2t did not keep theme inputs inside the themes root",
  );

  const cleaned = await converter.convert(params);
  assert(
    !("params.xml" in cleaned.media) &&
      !("nested/sentinel.bin" in cleaned.media) &&
      !("themes/doc.docx" in (cleaned.themes ?? {})) &&
      !("themes/nested/sentinel.xml" in (cleaned.themes ?? {})),
    "x2t leaked nested MEMFS inputs into the next successful conversion",
  );
}

export function resetAll() {
  onlyOfficeManagerFactory.destroyAll();
  editorManagerFactory.destroyAll();
  converter.terminate();
}

export async function runScenario(
  mode: ResourceMode,
  cdnOrigin: string,
  onStepsChange?: (steps: StepResult[]) => void,
) {
  const steps: StepResult[] = [];

  const runStep = async (
    name: string,
    action: () => Promise<string | void>,
  ) => {
    const stepIndex =
      steps.push({ name, status: "running", detail: "running" }) - 1;
    onStepsChange?.([...steps]);

    try {
      const detail = await action();
      steps[stepIndex] = {
        name,
        status: "passed",
        detail: detail || undefined,
      };
      onStepsChange?.([...steps]);
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error);
      steps[stepIndex] = { name, status: "failed", detail };
      onStepsChange?.([...steps]);
      throw new Error(`${name}: ${detail}`);
    }
  };

  resetAll();

  await runStep("resource mode", async () => {
    if (mode === "cdn") {
      OnlyOfficeManager.registerStaticResource({ cdnOrigin });
    } else {
      OnlyOfficeManager.resetStaticResource();
    }

    assert(isOnlyOfficeCdnMode() === (mode === "cdn"), "CDN mode mismatch");
    const expectedRoot =
      mode === "cdn"
        ? `${cdnOrigin}/onlyoffice/9.4.0-develop`
        : "/packages/onlyoffice/9.4.0-develop";
    assert(
      STATIC_RESOURCE.onlyoffice.root === expectedRoot,
      `Unexpected SDK root: ${STATIC_RESOURCE.onlyoffice.root}`,
    );
    return mode === "cdn" ? cdnOrigin : "local packages";
  });

  await runStep("CDN Nexolyra plugin READY", async () => {
    if (mode !== "cdn") return "covered by the CDN scenario";

    const container = document.getElementById(CONTAINER_IDS.plugin);
    assert(container instanceof HTMLElement, "Missing plugin editor container");

    let resolvePluginReady!: (value: {
      pluginGuid: string;
      editorType: string;
    }) => void;
    const pluginReady = new Promise<{
      pluginGuid: string;
      editorType: string;
    }>((resolve) => {
      resolvePluginReady = resolve;
    });
    const mount = mountOfficeEditor(container, {
      hostUrl: "https://plugin-host.example.test/office-host.html",
      emptyType: "docx",
      plugins: {
        configUrls: ["/e2e/nexolyra-plugin/config.json"],
        autostart: [NEXOLYRA_PLUGIN_GUID],
      },
      onPluginReady(pluginGuid, editorType) {
        resolvePluginReady({ pluginGuid, editorType });
      },
    });

    let timeoutId = 0;
    try {
      const instance = await mount.activate();
      const ready = await Promise.race([
        pluginReady,
        new Promise<never>((_resolve, reject) => {
          timeoutId = window.setTimeout(
            () => reject(new Error("CDN background plugin did not post READY")),
            DOCUMENT_READY_TIMEOUT_MS,
          );
        }),
      ]);
      assert(
        ready.pluginGuid === NEXOLYRA_PLUGIN_GUID &&
          ready.editorType === "word",
        `Unexpected plugin READY: ${ready.pluginGuid}/${ready.editorType}`,
      );

      const result = (await instance.invokePlugin(NEXOLYRA_PLUGIN_GUID, {
        type: "ping",
      })) as {
        pong?: boolean;
        editorType?: string;
        bridgeHostIsTop?: boolean;
        entryPath?: string;
      };
      assert(result.pong === true, "Plugin INVOKE did not reach the fixture");
      assert(
        result.editorType === "word" && result.bridgeHostIsTop === true,
        "Plugin did not use the Nexolyra window.parent.parent bridge",
      );
      assert(
        result.entryPath === "/e2e/nexolyra-plugin/index.html",
        `Relative plugin entry was not resolved to the embed origin: ${result.entryPath}`,
      );
      return `${ready.editorType} READY + INVOKE/RESULT`;
    } finally {
      if (timeoutId) window.clearTimeout(timeoutId);
      await mount.destroy();
      editorManagerFactory.destroy(container);
    }
  });

  await runStep("manager factory concurrent open", async () => {
    const options = {
      containerId: CONTAINER_IDS.concurrent,
      fileType: FILE_TYPE.DOCX,
      defaultFileName: "Concurrent.docx",
      readOnly: false,
      theme: DEFAULT_OFFICE_THEME,
    };
    const results = await withDocumentReady(
      () =>
        Promise.allSettled([
          onlyOfficeManagerFactory.open(options, {
            fileName: "Concurrent-A.docx",
            isNew: true,
          }),
          onlyOfficeManagerFactory.open(options, {
            fileName: "Concurrent-B.docx",
            isNew: true,
          }),
        ]),
      "concurrent open",
      CONTAINER_IDS.concurrent,
    );
    const fulfilled = results.filter(
      (result): result is PromiseFulfilledResult<OnlyOfficeManager> =>
        result.status === "fulfilled",
    );
    const rejected = results.filter(
      (result): result is PromiseRejectedResult => result.status === "rejected",
    );
    assert(fulfilled.length === 1, "Expected one latest concurrent open");
    assert(rejected.length === 1, "Expected one superseded concurrent open");
    assert(
      rejected[0]?.reason instanceof DOMException &&
        rejected[0].reason.name === "AbortError",
      "Superseded concurrent open did not reject with AbortError",
    );

    const manager = onlyOfficeManagerFactory.get(CONTAINER_IDS.concurrent);
    assert(manager === fulfilled[0]?.value, "Factory cached a stale manager");
    assert(manager?.isReady(), "Latest concurrent manager was not ready");
    await assertExport(manager, FILE_TYPE.DOCX);
    onlyOfficeManagerFactory.destroy(CONTAINER_IDS.concurrent);
    assert(
      !onlyOfficeManagerFactory.get(CONTAINER_IDS.concurrent),
      "Factory destroy retained the concurrent manager",
    );

    let releaseSlowFile!: (value: ArrayBuffer) => void;
    let markSlowFileStarted!: () => void;
    const slowFileStarted = new Promise<void>((resolve) => {
      markSlowFileStarted = resolve;
    });
    const slowFileData = new Promise<ArrayBuffer>((resolve) => {
      releaseSlowFile = resolve;
    });
    const slowFile = new File(["slow"], "Slow.docx");
    Object.defineProperty(slowFile, "arrayBuffer", {
      value: () => {
        markSlowFileStarted();
        return slowFileData;
      },
    });

    const staleOpen = onlyOfficeManagerFactory.open(options, {
      fileName: slowFile.name,
      file: slowFile,
    });
    await slowFileStarted;
    onlyOfficeManagerFactory.destroy(CONTAINER_IDS.concurrent);

    const replacement = await withDocumentReady(
      () =>
        onlyOfficeManagerFactory.open(options, {
          fileName: "Replacement.docx",
          isNew: true,
        }),
      "replacement open",
      CONTAINER_IDS.concurrent,
    );
    const scopedIo = getScopedIoRegistry()[CONTAINER_IDS.concurrent];
    assert(scopedIo, "Replacement editor did not register scoped IO");

    releaseSlowFile(new Uint8Array([1, 2, 3]).buffer);
    const staleResult = await Promise.allSettled([staleOpen]);
    assert(
      staleResult[0]?.status === "rejected" &&
        staleResult[0].reason instanceof DOMException &&
        staleResult[0].reason.name === "AbortError",
      "Destroyed slow open did not reject with AbortError",
    );
    assert(
      getScopedIoRegistry()[CONTAINER_IDS.concurrent] === scopedIo,
      "Stale editor removed the replacement scoped IO",
    );
    await assertExport(replacement, FILE_TYPE.DOCX);
    onlyOfficeManagerFactory.destroy(CONTAINER_IDS.concurrent);
    return "one latest open, one AbortError";
  });

  await runStep("manager factory open/get", async () => {
    const manager = await withDocumentReady(
      () =>
        onlyOfficeManagerFactory.open(
          {
            containerId: CONTAINER_IDS.factory,
            fileType: FILE_TYPE.DOCX,
            defaultFileName: "Factory.docx",
            readOnly: false,
            theme: DEFAULT_OFFICE_THEME,
            user: { id: "factory-user", name: "Factory User" },
          },
          {
            fileName: "Factory.docx",
            isNew: true,
          },
        ),
      "factory open",
      CONTAINER_IDS.factory,
    );

    assert(manager.isReady(), "Factory manager was not ready");
    assert(
      onlyOfficeManagerFactory.get(CONTAINER_IDS.factory) === manager,
      "Factory get did not return the opened manager",
    );

    const origin = frameOrigin(CONTAINER_IDS.factory);
    if (mode === "cdn") {
      assert(
        origin === new URL(cdnOrigin).origin,
        "Expected CDN iframe origin",
      );
    } else {
      assert(origin === window.location.origin, "Expected local iframe origin");
    }
  });

  await runStep("manager facade from factory", async () => {
    const manager = onlyOfficeManagerFactory.get(CONTAINER_IDS.factory);
    assert(manager, "Factory manager is missing");

    manager.setUser({ id: "updated-user", name: "Updated User" });
    assert(manager.getUser().id === "updated-user", "setUser/getUser failed");

    await withDocumentReady(
      () => manager.setLanguage("en"),
      "setLanguage",
      CONTAINER_IDS.factory,
    );
    assert(manager.getLanguage() === "en", "setLanguage failed");

    await withDocumentReady(
      () => manager.setTheme(OFFICE_THEME.DARK),
      "setTheme",
      CONTAINER_IDS.factory,
    );
    assert(
      manager.getTheme() === (OFFICE_THEME.DARK as OfficeTheme),
      "setTheme failed",
    );

    await manager.setReadOnly(true);
    assert(manager.getReadOnly(), "setReadOnly(true) failed");
    await manager.setReadOnly(false);
    assert(!manager.getReadOnly(), "setReadOnly(false) failed");

    await withDocumentReady(
      () => manager.openNew("Factory-Reopened.docx"),
      "openNew",
      CONTAINER_IDS.factory,
    );
    await assertExport(manager, FILE_TYPE.DOCX);
  });

  await runStep("manager connector", async () => {
    const manager = onlyOfficeManagerFactory.get(CONTAINER_IDS.factory);
    assert(manager, "Factory manager is missing");

    const connector = manager.createConnector();
    try {
      await new Promise<void>((resolve, reject) => {
        const timeout = window.setTimeout(() => {
          reject(new Error("Connector did not return within 10 seconds"));
        }, 10_000);
        connector.executeMethod("GetEditorType", [], () => {
          window.clearTimeout(timeout);
          resolve();
        });
      });
    } finally {
      connector.disconnect();
    }

    return "GetEditorType callback received";
  });

  await runStep("manager factory destroy", async () => {
    onlyOfficeManagerFactory.destroy(CONTAINER_IDS.factory);
    assert(
      !onlyOfficeManagerFactory.get(CONTAINER_IDS.factory),
      "Factory destroy did not remove manager",
    );
  });

  await runStep("fixture manifest", async () => {
    const response = await fetch("/e2e/fixtures/manifest.json");
    assert(response.ok, "Fixture manifest is not available");
    const manifest = (await response.json()) as Array<{
      name: string;
      kind: "positive" | "negative";
      source: string;
    }>;
    const names = new Set(manifest.map((fixture) => fixture.name));
    assert(
      names.has("edge-invalid-bookmark.docx"),
      "Missing invalid bookmark fixture",
    );
    assert(names.has("xml-limit.docx"), "Missing XML limit fixture");
    assert(names.has("mismatch-xlsx-as-docx.docx"), "Missing mismatch fixture");
    return `${manifest.length} generated fixtures`;
  });

  await runStep("x2t edge imports", async () => {
    await assertX2tImport("edge-invalid-bookmark.docx", FILE_TYPE.DOCX);
    await assertX2tRejectsTraversalAndRecovers();
    return "DOCX converted; traversal rejected; clean Worker recovered";
  });

  await runStep("generated negative fixtures", async () => {
    const file = await fetchPublicFile(
      "/e2e/fixtures/xml-limit.docx",
      "xml-limit.docx",
    );
    const manager = await OnlyOfficeManager.createWithFile(
      {
        containerId: CONTAINER_IDS.fixture,
        fileType: FILE_TYPE.DOCX,
        defaultFileName: file.name,
        readOnly: false,
        theme: DEFAULT_OFFICE_THEME,
        officeXmlEvent: {
          isEnable: true,
          limitBytes: 1024,
        },
      },
      file,
    );

    assert(
      manager.getEditor().isOfficeXmlSizeLimitExceeded(),
      "XML size guard did not block oversized Office XML",
    );
    manager.destroy();
    editorManagerFactory.destroy(CONTAINER_IDS.fixture);

    const mismatch = await fetchPublicFile(
      "/e2e/fixtures/mismatch-xlsx-as-docx.docx",
      "mismatch-xlsx-as-docx.docx",
    );
    assert(mismatch.size > 0, "Mismatch negative fixture is empty");

    return "xml guard blocked, mismatch fixture available";
  });

  await runStep("manager create", async () => {
    const manager = await withDocumentReady(
      () =>
        OnlyOfficeManager.create({
          containerId: CONTAINER_IDS.create,
          fileType: FILE_TYPE.PPTX,
          defaultFileName: "FactoryDeck.pptx",
          readOnly: false,
          theme: DEFAULT_OFFICE_THEME,
        }),
      "manager create",
      CONTAINER_IDS.create,
    );

    assert(manager.isReady(), "OnlyOfficeManager.create was not ready");
    assert(manager.getEditor().exists(), "Created editor does not exist");
    const connector = manager.createConnector();
    try {
      await new Promise<void>((resolve, reject) => {
        const timeout = window.setTimeout(() => {
          reject(
            new Error(
              "Presentation connector did not return within 10 seconds",
            ),
          );
        }, 10_000);
        connector.callCommand(
          function () {
            const fill = Api.CreateSolidFill(Api.CreateRGBColor(230, 247, 255));
            const stroke = Api.CreateStroke(
              0,
              Api.CreateSolidFill(Api.CreateRGBColor(24, 144, 255)),
            );
            const shape = Api.CreateShape(
              "rect",
              7_200_000,
              900_000,
              fill,
              stroke,
            );
            shape.SetPosition(1_000_000, 1_000_000);
            shape
              .GetDocContent()
              .GetContent()[0]
              ?.AddText("[Connector] wrote this text box.");
            Api.GetPresentation().GetCurrentSlide().AddObject(shape);
          },
          () => {
            window.clearTimeout(timeout);
            resolve();
          },
        );
      });
    } finally {
      connector.disconnect();
    }
    await manager.toggleReadOnly();
    assert(manager.getReadOnly(), "toggleReadOnly failed");
    await assertExport(manager, FILE_TYPE.PPTX);
    manager.destroy();
    editorManagerFactory.destroy(CONTAINER_IDS.create);
  });

  await runStep("manager createWithFile", async () => {
    const file = await fetchPublicFile("/test.xlsx", "test.xlsx");
    const manager = await withDocumentReady(
      () =>
        OnlyOfficeManager.createWithFile(
          {
            containerId: CONTAINER_IDS.file,
            fileType: FILE_TYPE.XLSX,
            defaultFileName: "test.xlsx",
            readOnly: false,
            theme: DEFAULT_OFFICE_THEME,
          },
          file,
        ),
      "createWithFile",
      CONTAINER_IDS.file,
    );

    assert(manager.isReady(), "OnlyOfficeManager.createWithFile was not ready");
    await withDocumentReady(
      () => manager.openFile(file),
      "openFile",
      CONTAINER_IDS.file,
    );
    await assertExport(manager, FILE_TYPE.XLSX);
    manager.destroy();
    editorManagerFactory.destroy(CONTAINER_IDS.file);
  });

  await runStep("manager spreadsheet connector", async () => {
    const file = await fetchPublicFile("/test.xlsx", "test.xlsx");
    const manager = await withDocumentReady(
      () =>
        OnlyOfficeManager.createWithFile(
          {
            containerId: CONTAINER_IDS.file,
            fileType: FILE_TYPE.XLSX,
            defaultFileName: file.name,
            readOnly: false,
            theme: DEFAULT_OFFICE_THEME,
          },
          file,
        ),
      "spreadsheet createWithFile",
      CONTAINER_IDS.file,
    );

    const connector = manager.createConnector();
    try {
      await new Promise<void>((resolve, reject) => {
        const timeout = window.setTimeout(() => {
          reject(
            new Error("Spreadsheet connector did not return within 10 seconds"),
          );
        }, 10_000);
        connector.callCommand(
          function () {
            Api.GetActiveSheet()
              .GetRange("A1")
              .SetValue("[Connector] wrote this cell.");
          },
          () => {
            window.clearTimeout(timeout);
            resolve();
          },
        );
      });
      await assertExport(manager, FILE_TYPE.XLSX);
    } finally {
      connector.disconnect();
      manager.destroy();
      editorManagerFactory.destroy(CONTAINER_IDS.file);
    }

    return "Wrote [Connector] wrote this cell. to A1";
  });

  await runStep("text fallback files", async () => {
    const textDocx = await fetchPublicFile(
      "/e2e/fixtures/plain-text-as-docx.docx",
      "plain-text-as-docx.docx",
    );
    const docxManager = await withDocumentReady(
      () =>
        OnlyOfficeManager.createWithFile(
          {
            containerId: CONTAINER_IDS.textFallback,
            fileType: FILE_TYPE.DOCX,
            defaultFileName: textDocx.name,
            readOnly: false,
            theme: DEFAULT_OFFICE_THEME,
          },
          textDocx,
        ),
      "text DOCX fallback",
      CONTAINER_IDS.textFallback,
    );

    assert(docxManager.isReady(), "Text DOCX fallback was not ready");
    await assertExport(docxManager, FILE_TYPE.DOCX);
    docxManager.destroy();
    editorManagerFactory.destroy(CONTAINER_IDS.textFallback);

    const textXlsx = await fetchPublicFile(
      "/e2e/fixtures/plain-text-as-xlsx.xlsx",
      "plain-text-as-xlsx.xlsx",
    );
    const xlsxManager = await withDocumentReady(
      () =>
        OnlyOfficeManager.createWithFile(
          {
            containerId: CONTAINER_IDS.textFallback,
            fileType: FILE_TYPE.XLSX,
            defaultFileName: textXlsx.name,
            readOnly: false,
            theme: DEFAULT_OFFICE_THEME,
          },
          textXlsx,
        ),
      "text XLSX fallback",
      CONTAINER_IDS.textFallback,
    );

    assert(xlsxManager.isReady(), "Text XLSX fallback was not ready");
    await assertExport(xlsxManager, FILE_TYPE.XLSX);
    xlsxManager.destroy();
    editorManagerFactory.destroy(CONTAINER_IDS.textFallback);
  });

  await runStep("manager fromEditor", async () => {
    const editor = editorManagerFactory.get(CONTAINER_IDS.fromEditor);
    const manager = OnlyOfficeManager.fromEditor(editor, {
      containerId: CONTAINER_IDS.fromEditor,
      fileType: FILE_TYPE.DOCX,
      defaultFileName: "FromEditor.docx",
      readOnly: false,
      theme: DEFAULT_OFFICE_THEME,
    });

    await withDocumentReady(
      () =>
        manager.openDocument({
          fileName: "FromEditor.docx",
          isNew: true,
        }),
      "fromEditor open",
      CONTAINER_IDS.fromEditor,
    );

    assert(manager.isReady(), "OnlyOfficeManager.fromEditor was not ready");
    assert(manager.getEditor() === editor, "fromEditor did not keep editor");
    await assertExport(manager, FILE_TYPE.DOCX);
    manager.destroy();
    editorManagerFactory.destroy(CONTAINER_IDS.fromEditor);
  });

  await runStep("manager factory destroyAll", async () => {
    onlyOfficeManagerFactory.destroyAll();
    assert(
      !onlyOfficeManagerFactory.get(CONTAINER_IDS.factory),
      "Factory destroyAll left a manager behind",
    );
    editorManagerFactory.destroyAll();
  });

  return steps;
}

function OnlyOfficeTestEditor({ containerId }: { containerId: string }) {
  return (
    <div
      className={`${ONLYOFFICE_CONTAINER_CONFIG.PARENT_CLASS_NAME} relative h-[420px] min-h-[420px] border border-neutral-200 bg-white`}
      data-onlyoffice-container-id={containerId}
    >
      <div id={containerId} className="absolute inset-0" />
    </div>
  );
}

export function OnlyOfficeFactoryE2EPage() {
  const [params, setParams] = useState<{
    mode: ResourceMode;
    cdnOrigin: string;
  }>({ mode: "local", cdnOrigin: "" });
  const [paramsReady, setParamsReady] = useState(false);
  useEffect(() => {
    const search = new URLSearchParams(window.location.search);
    setParams({
      mode: search.get("mode") === "cdn" ? "cdn" : "local",
      cdnOrigin: search.get("cdnOrigin") || "http://127.0.0.1:3010",
    });
    setParamsReady(true);
  }, []);

  const [result, setResult] = useState<ScenarioResult>({
    mode: params.mode,
    status: "idle",
    steps: [],
  });

  useEffect(() => {
    if (!paramsReady) return;
    let disposed = false;

    window.requestAnimationFrame(() => {
      if (disposed) return;

      setResult({ mode: params.mode, status: "running", steps: [] });

      // 用例入口
      let latestSteps: StepResult[] = [];
      const updateSteps = (steps: StepResult[]) => {
        latestSteps = steps;
        if (!disposed) {
          setResult((current) => ({ ...current, steps }));
        }
      };

      runScenario(params.mode, params.cdnOrigin, updateSteps)
        .then((steps) => {
          if (!disposed) {
            setResult({ mode: params.mode, status: "passed", steps });
          }
        })
        .catch((error) => {
          if (!disposed) {
            setResult((current) => ({
              ...current,
              status: "failed",
              steps: latestSteps,
              error: error instanceof Error ? error.message : String(error),
            }));
          }
        });
    });

    return () => {
      disposed = true;
      resetAll();
      OnlyOfficeManager.resetStaticResource();
    };
  }, [params.cdnOrigin, params.mode, paramsReady]);

  return (
    <main className="min-h-screen bg-neutral-50 p-4 text-neutral-900">
      <section className="mb-4 border border-neutral-200 bg-white p-3">
        <h1 className="text-base font-semibold">OnlyOffice factory e2e</h1>
        <p className="text-sm text-neutral-600">
          <span data-testid="scenario-status">{result.status}</span>
          {" · "}
          <span>{result.mode}</span>
        </p>
        {result.error && (
          <p className="mt-2 text-sm text-red-600" data-testid="scenario-error">
            {result.error}
          </p>
        )}
        <pre
          className="mt-3 max-h-64 overflow-auto bg-neutral-950 p-3 text-xs text-neutral-50"
          data-testid="scenario-result"
        >
          {JSON.stringify(result, null, 2)}
        </pre>
      </section>

      <div className="grid gap-4">
        <OnlyOfficeTestEditor containerId={CONTAINER_IDS.plugin} />
        <OnlyOfficeTestEditor containerId={CONTAINER_IDS.factory} />
        <OnlyOfficeTestEditor containerId={CONTAINER_IDS.concurrent} />
        <OnlyOfficeTestEditor containerId={CONTAINER_IDS.create} />
        <OnlyOfficeTestEditor containerId={CONTAINER_IDS.file} />
        <OnlyOfficeTestEditor containerId={CONTAINER_IDS.textFallback} />
        <OnlyOfficeTestEditor containerId={CONTAINER_IDS.fromEditor} />
        <OnlyOfficeTestEditor containerId={CONTAINER_IDS.fixture} />
      </div>
    </main>
  );
}
