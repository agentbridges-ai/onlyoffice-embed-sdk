"use client";

/**
 * 单实例演示页：独立 Subframe + RPC 工具栏（上传/导出/主题/语言/只读）。
 * 文档说明见 `onlyoffice-web-comp/docs/单实例示例.md`。
 */
import { memo, useCallback, useEffect, useRef, useState } from "react";
import {
  ONLYOFFICE_LANG_KEY,
  OFFICE_XML_EVENT_CONFIG,
  OFFICE_THEME_OPTIONS,
  DEFAULT_OFFICE_THEME,
  type FileType,
  type OfficeTheme,
} from "@/components/onlyoffice-web-comp";

import {
  DemoButton,
  DemoField,
  DemoMenu,
  DemoMenuRow,
  DemoSelect,
  demoHeaderClass,
  demoHeaderInnerClass,
  demoTitleClass,
  demoToolbarClass,
} from "./demo-toolbar";
import { DocxCommentsCrud } from "./docx-comments-crud";
import { DocxRevisionsCrud } from "./docx-revisions-crud";
import { getFileExtension, OFFICE_UPLOAD_ACCEPT } from "./office-formats";
import {
  createConnectorDemo,
  type ConnectorDemo,
} from "./connector-demo";
import { CHINESE_ZODIAC_SLOTS } from "./chinese-zodiac-slots";
import {
  applyDemoResourceMode,
  getDemoResourceState,
  ResourceSwitcher,
  subscribeDemoResourceChange,
} from "./resource-switcher";
import { getSubframeOrigin, SubframeManager } from "./subframe-manager";
import { SUBFRAME_EDITOR_CONTAINER_ID as SUBFRAME_CONTAINER_ID } from "./subframe-protocol";

type OfficePreviewPageProps = {
  title: string;
  defaultFileName: string;
  fileType: FileType;
  accept?: string;
  newButtonLabel: string;
  /** public 目录下的默认文件路径，如 /test.xlsx */
  initialFileUrl?: string;
  /** 嵌入文档页或父容器时使用 h-full */
  embedded?: boolean;
};

type ConnectorMessage = {
  text: string;
  tone: "success" | "error";
};

function LoadingOverlay() {
  return (
    <div className="absolute inset-0 z-20 flex items-center justify-center bg-white/70 backdrop-blur-sm">
      <div className="rounded-lg border border-gray-200 bg-white px-4 py-3 text-sm text-gray-700 shadow">
        加载中...
      </div>
    </div>
  );
}

const OnlyOfficeSubframeHost = memo(function OnlyOfficeSubframeHost({
  src,
  title,
  onFrame,
  onLoad,
}: {
  src: string;
  title: string;
  onFrame: (frame: HTMLIFrameElement | null) => void;
  onLoad: () => void;
}) {
  return (
    <iframe
      ref={onFrame}
      src={src}
      title={title}
      data-onlyoffice-subframe="true"
      data-onlyoffice-instance-id="single-editor"
      data-onlyoffice-zodiac="rat"
      className="absolute inset-0 h-full w-full border-0"
      onLoad={onLoad}
    />
  );
});

async function fetchPublicFile(url: string, fileName: string) {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Failed to fetch ${url}`);
  }
  const blob = await response.blob();
  return new File([blob], fileName, { type: blob.type });
}

export function OfficePreviewPage({
  title,
  defaultFileName,
  fileType,
  accept = OFFICE_UPLOAD_ACCEPT,
  newButtonLabel,
  initialFileUrl,
  embedded = false,
}: OfficePreviewPageProps) {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const frameRef = useRef<HTMLIFrameElement | null>(null);
  const managerRef = useRef<SubframeManager | null>(null);
  const connectorRef = useRef<ConnectorDemo | null>(null);
  const [mounted, setMounted] = useState(false);
  const [frameReady, setFrameReady] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [readOnly, setReadOnly] = useState(false);
  const [currentLang, setCurrentLangState] = useState(
    ONLYOFFICE_LANG_KEY.ZH as string,
  );
  const [currentTheme, setCurrentThemeState] = useState<OfficeTheme>(
    DEFAULT_OFFICE_THEME,
  );
  const [officeXmlEventEnabled, setOfficeXmlEventEnabled] = useState<boolean>(
    OFFICE_XML_EVENT_CONFIG.default.isEnable,
  );
  const [officeXmlLimitMb, setOfficeXmlLimitMb] = useState(
    Math.round(OFFICE_XML_EVENT_CONFIG.default.limitBytes / 1024 / 1024),
  );
  const [editorReady, setEditorReady] = useState(false);
  const [connectorMessage, setConnectorMessage] = useState<ConnectorMessage | null>(
    null,
  );
  const connectorMessageTimerRef = useRef<number | null>(null);
  const [activeFileName, setActiveFileName] = useState(defaultFileName);
  const [cdnOrigin, setCdnOrigin] = useState(
    () => getDemoResourceState().cdnOrigin,
  );
  const [resourceMode, setResourceMode] = useState(
    () => getDemoResourceState().mode,
  );
  const [resourceRevision, setResourceRevision] = useState(
    () => getDemoResourceState().revision,
  );

  const replaceConnector = (manager: SubframeManager) => {
    if (connectorRef.current?.isConnected) {
      connectorRef.current.disconnect();
    }
    const connector = createConnectorDemo(() => manager);
    connectorRef.current = connector;
    return connector;
  };

  const handleFrame = useCallback((frame: HTMLIFrameElement | null) => {
    frameRef.current = frame;
    if (frame) {
      managerRef.current?.attachFrame(frame);
    } else {
      setFrameReady(false);
      managerRef.current?.detachFrame();
    }
  }, []);

  const handleFrameLoad = useCallback(() => {
    setFrameReady(true);
    const frame = frameRef.current;
    if (frame) managerRef.current?.attachFrame(frame);
  }, []);

  useEffect(() => {
    setMounted(true);
  }, []);

  useEffect(() => {
    const onMessage = (event: MessageEvent) => {
      managerRef.current?.handleMessage(event);
    };
    window.addEventListener("message", onMessage);
    return () => window.removeEventListener("message", onMessage);
  }, []);

  useEffect(
    () =>
      subscribeDemoResourceChange((state) => {
        setCdnOrigin(state.cdnOrigin);
        setResourceMode(state.mode);
        setLoading(true);
        if (connectorRef.current?.isConnected) {
          connectorRef.current.disconnect();
        }
        connectorRef.current = null;
        managerRef.current?.destroy();
        managerRef.current = null;
        setFrameReady(false);
        setActiveFileName(defaultFileName);
        setConnectorMessage(null);
        setResourceRevision(state.revision);
      }),
    [defaultFileName],
  );

  useEffect(() => {
    if (!frameReady || !frameRef.current) return;

    let disposed = false;
    const officeXmlEvent = {
      isEnable: officeXmlEventEnabled,
      limitBytes: officeXmlLimitMb * 1024 * 1024,
    };
    const manager = new SubframeManager({
      containerId: SUBFRAME_CONTAINER_ID,
      fileType,
      defaultFileName,
      readOnly,
      theme: currentTheme,
      officeXmlEvent,
      instanceId: "single-editor",
      targetOrigin: getSubframeOrigin(CHINESE_ZODIAC_SLOTS[0].id),
      frame: frameRef.current,
      onEvent: (event, payload) => {
        if (event === "loading-change") {
          setLoading(!!(payload as { loading?: boolean } | undefined)?.loading);
        }
        if (event === "office-xml-size-limit-exceeded") {
          setError("文件过大，不支持解析");
        }
      },
    });
    managerRef.current = manager;
    setEditorReady(false);
    setLoading(true);
    setError(null);

    const init = async () => {
      const document = initialFileUrl
        ? {
            fileName: defaultFileName,
            file: await fetchPublicFile(initialFileUrl, defaultFileName),
          }
        : {
            fileName: defaultFileName,
            isNew: true,
          };
      if (disposed) return;

      manager.configure({
        defaultFileName,
        readOnly,
        theme: currentTheme,
        officeXmlEvent,
      });
      await manager.openDocument(document);
      if (disposed) return;

      replaceConnector(manager);
      setCurrentLangState(ONLYOFFICE_LANG_KEY.ZH);
      setCurrentThemeState(manager.getTheme());
      setEditorReady(true);
      setLoading(false);
      setActiveFileName(document.fileName);
    };

    init().catch((err) => {
      if (disposed) return;
      setError("无法加载编辑器组件");
      setLoading(false);
      console.error("Failed to initialize OnlyOffice subframe:", err);
    });

    return () => {
      disposed = true;
      if (connectorRef.current?.isConnected) {
        connectorRef.current.disconnect();
      }
      connectorRef.current = null;
      manager.destroy();
      if (managerRef.current === manager) {
        managerRef.current = null;
      }
      setEditorReady(false);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [frameReady, resourceRevision]);

  useEffect(
    () => () => {
      if (connectorMessageTimerRef.current !== null) {
        window.clearTimeout(connectorMessageTimerRef.current);
      }
    },
    [],
  );

  const getSubframeSrc = () => {
    if (!mounted) return "";

    const url = new URL(
      "/subframe",
      getSubframeOrigin(CHINESE_ZODIAC_SLOTS[0].id),
    );
    url.searchParams.set("instance", "single-editor");
    url.searchParams.set("resourceMode", resourceMode);
    url.searchParams.set("resourceRevision", String(resourceRevision));
    if (resourceMode === "cdn") {
      url.searchParams.set("cdnOrigin", cdnOrigin);
    }
    return url.href;
  };

  const runAction = async (action: () => Promise<void>, message: string) => {
    try {
      setError(null);
      await action();
    } catch (err) {
      setError(message);
      console.error(message, err);
    }
  };

  const showConnectorMessage = (text: string, tone: ConnectorMessage["tone"]) => {
    if (connectorMessageTimerRef.current !== null) {
      window.clearTimeout(connectorMessageTimerRef.current);
    }
    setConnectorMessage({ text, tone });
    connectorMessageTimerRef.current = window.setTimeout(() => {
      setConnectorMessage(null);
      connectorMessageTimerRef.current = null;
    }, 4_000);
  };

  const handleOpenDocument = (
    fileName: string,
    file?: File,
    nextReadOnly = readOnly,
  ) =>
    runAction(async () => {
      const manager = managerRef.current;
      if (!manager) {
        throw new Error("Editor is not initialized");
      }
      manager.configure({
        defaultFileName,
        readOnly: nextReadOnly,
        theme: currentTheme,
        officeXmlEvent: {
          isEnable: officeXmlEventEnabled,
          limitBytes: officeXmlLimitMb * 1024 * 1024,
        },
      });
      await manager.openDocument({
        fileName,
        file,
        isNew: !file,
        readOnly: nextReadOnly,
      });
      replaceConnector(manager);
      setActiveFileName(fileName);
      setReadOnly(nextReadOnly);
      setConnectorMessage(null);
    }, "操作失败");

  const handleLanguageSwitch = () =>
    runAction(async () => {
      const manager = managerRef.current;
      if (!manager) {
        throw new Error("Editor is not initialized");
      }
      const nextLang = await manager.toggleLanguage();
      setCurrentLangState(nextLang);
    }, "切换语言失败");

  const handleThemeChange = (theme: OfficeTheme) =>
    runAction(async () => {
      const manager = managerRef.current;
      if (!manager) {
        throw new Error("Editor is not initialized");
      }
      await manager.setTheme(theme);
      setCurrentThemeState(manager.getTheme());
    }, "切换主题失败");

  const handleExport = () =>
    runAction(async () => {
      const manager = managerRef.current;
      if (!manager) {
        throw new Error("Editor is not initialized");
      }

      await manager.downloadExport();
    }, "导出失败");

  const handlePrintLogs = () => {
    managerRef.current?.printLogs();
  };

  const handleConnector = () =>
    runAction(async () => {
      const manager = managerRef.current;
      if (!manager) {
        throw new Error("Editor is not initialized");
      }

      const logger = manager.getLogger();
      try {
        logger.operation("Connector command started", {
          fileName: activeFileName,
          fileType,
        });
        const connector = connectorRef.current ?? replaceConnector(manager);
        const result = await connector.write(activeFileName, fileType);
        showConnectorMessage(result.message, "success");
        logger.operation("Connector command completed", {
          fileName: activeFileName,
          fileType: result.fileType,
        });
      } catch (error) {
        logger.error("operation", "Connector command failed", {
          fileName: activeFileName,
          fileType,
          error,
        });
        showConnectorMessage(
          error instanceof Error ? error.message : "Connector command failed",
          "error",
        );
        throw error;
      }
    }, "连接器调用失败");

  const handleToggleReadOnly = () =>
    runAction(async () => {
      const manager = managerRef.current;
      if (!manager) {
        throw new Error("Editor is not initialized");
      }
      await manager.setReadOnly(!manager.getReadOnly());
      setReadOnly(manager.getReadOnly());
    }, "切换模式失败");

  const handleResourceLoad = () =>
    runAction(async () => {
      applyDemoResourceMode("cdn", cdnOrigin);
    }, "切换资源失败");

  const isDocxFile = getFileExtension(activeFileName, fileType) === "docx";

  return (
    <div
      className={`flex flex-col bg-white ${
        embedded ? "h-full min-h-0" : "h-screen"
      }`}
    >
      <header className={demoHeaderClass}>
        <div className={demoHeaderInnerClass}>
          <div className="mr-auto flex min-w-0 items-baseline gap-2.5">
            <h1 className={demoTitleClass}>{title}</h1>
          </div>

          <div className={demoToolbarClass}>
            <DemoButton onClick={() => fileInputRef.current?.click()}>
              上传
            </DemoButton>
            <DemoButton onClick={() => handleOpenDocument(defaultFileName)}>
              {newButtonLabel}
            </DemoButton>
            {editorReady && (
              <>
                <DemoButton onClick={handleExport}>导出</DemoButton>
                <DemoButton onClick={handlePrintLogs}>打印日志</DemoButton>
                <DemoButton onClick={handleConnector}>连接器写入</DemoButton>
                <DemoButton active={readOnly} onClick={handleToggleReadOnly}>
                  {readOnly ? "只读" : "编辑"}
                </DemoButton>
              </>
            )}
            <DemoMenu label="更多" disabled={loading}>
              <DemoMenuRow>
                <ResourceSwitcher
                  cdnOrigin={cdnOrigin}
                  disabled={loading}
                  onCdnOriginChange={setCdnOrigin}
                  onLoad={handleResourceLoad}
                />
              </DemoMenuRow>
              <DemoMenuRow>
                <DemoButton onClick={handleLanguageSwitch}>
                  {currentLang === ONLYOFFICE_LANG_KEY.ZH ? "中文" : "English"}
                </DemoButton>
                <DemoField label="主题">
                  <DemoSelect
                    value={currentTheme}
                    onChange={(event) =>
                      handleThemeChange(event.target.value as OfficeTheme)
                    }
                    disabled={!editorReady}
                  >
                    {OFFICE_THEME_OPTIONS.map((option) => (
                      <option key={option.id} value={option.id}>
                        {option.label}
                      </option>
                    ))}
                  </DemoSelect>
                </DemoField>
              </DemoMenuRow>
              <DemoMenuRow>
                <DemoField label="XML 检测">
                  <input
                    type="checkbox"
                    checked={officeXmlEventEnabled}
                    onChange={(event) =>
                      setOfficeXmlEventEnabled(event.target.checked)
                    }
                    className="h-4 w-4"
                  />
                </DemoField>
                <DemoField label="阈值 MB">
                  <input
                    type="number"
                    min={1}
                    step={1}
                    value={officeXmlLimitMb}
                    onChange={(event) =>
                      setOfficeXmlLimitMb(
                        Math.max(1, Number(event.target.value) || 1),
                      )
                    }
                    className="h-6 w-20 border-0 bg-transparent py-0 pl-0.5 text-[13px] text-neutral-800 outline-none"
                  />
                </DemoField>
              </DemoMenuRow>
              {isDocxFile && (
                <>
                  <DocxCommentsCrud
                    disabled={!editorReady || loading || readOnly}
                    getManager={() => managerRef.current}
                    onError={(message, err) => {
                      setError(message);
                      console.error(message, err);
                    }}
                  />
                  <DocxRevisionsCrud
                    disabled={!editorReady || loading || readOnly}
                    getManager={() => managerRef.current}
                    onError={(message, err) => {
                      setError(message);
                      console.error(message, err);
                    }}
                  />
                </>
              )}
            </DemoMenu>
          </div>
        </div>
      </header>
      {connectorMessage && (
        <output
          aria-live="polite"
          className={`fixed top-4 left-1/2 z-50 -translate-x-1/2 rounded border px-4 py-2 text-sm shadow-lg ${
            connectorMessage.tone === "success"
              ? "border-emerald-200 bg-emerald-50 text-emerald-800"
              : "border-red-200 bg-red-50 text-red-800"
          }`}
          role="status"
        >
          {connectorMessage.text}
        </output>
      )}
      {error && (
        <div className="mx-4 mt-4 rounded border-l-4 border-red-500 bg-red-50 p-4 text-red-700">
          <p className="font-medium">错误：{error}</p>
        </div>
      )}

      <div className="relative min-h-0 flex-1">
        {mounted && (
          <OnlyOfficeSubframeHost
            src={getSubframeSrc()}
            title={`${title} · ${CHINESE_ZODIAC_SLOTS[0].emoji} ${CHINESE_ZODIAC_SLOTS[0].name}`}
            onFrame={handleFrame}
            onLoad={handleFrameLoad}
          />
        )}
        {loading && <LoadingOverlay />}
      </div>

      <input
        ref={fileInputRef}
        type="file"
        accept={accept}
        className="hidden"
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) {
            handleOpenDocument(file.name, file);
            if (fileInputRef.current) fileInputRef.current.value = "";
          }
        }}
      />
    </div>
  );
}
