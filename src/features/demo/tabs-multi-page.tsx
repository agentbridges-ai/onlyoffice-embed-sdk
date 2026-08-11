"use client";

/**
 * 多实例 Tab 演示：每个 Tab 运行在独立 origin 的 Subframe 中，切换 Tab 时只隐藏 iframe。
 * 文档与完整源码说明见 `onlyoffice-web-comp/docs/多实例示例.md`。
 */
import { memo, useCallback, useEffect, useRef, useState } from "react";
import { nanoid } from "nanoid";
import {
  DemoButton,
  DemoField,
  DemoMenu,
  DemoMenuRow,
  DemoSelect,
  demoHeaderClass,
  demoHeaderInnerClass,
  demoSubtitleClass,
  demoTitleClass,
  demoToolbarClass,
} from "./demo-toolbar";
import { DocxCommentsCrud } from "./docx-comments-crud";
import { DocxRevisionsCrud } from "./docx-revisions-crud";
import {
  createConnectorDemo,
  type ConnectorDemo,
} from "./connector-demo";
import { getFileExtension } from "./office-formats";
import {
  CHINESE_ZODIAC_SLOT_COUNT,
  CHINESE_ZODIAC_SLOTS,
  getChineseZodiacSlot,
  type ChineseZodiacSlot,
} from "./chinese-zodiac-slots";
import {
  getSubframeOrigin,
  SubframeManager,
} from "./subframe-manager";
import {
  applyDemoResourceMode,
  getDemoResourceState,
  ResourceSwitcher,
  subscribeDemoResourceChange,
} from "./resource-switcher";
import {
  DEFAULT_OFFICE_THEME,
  FILE_TYPE,
  OFFICE_XML_EVENT_CONFIG,
  OFFICE_THEME_OPTIONS,
  type FileType,
  type OfficeTheme,
} from "@/components/onlyoffice-web-comp";
import { SUBFRAME_EDITOR_CONTAINER_ID as SUBFRAME_CONTAINER_ID } from "./subframe-protocol";

type DocKind = "word" | "excel" | "ppt";

type DocPreset = {
  label: string;
  badge: string;
  fileType: FileType;
  defaultFileName: string;
  accept: string;
};

type TabItem = {
  id: string;
  label: string;
  containerId: string;
  subframeHost: string;
  chineseZodiac: ChineseZodiacSlot;
  fileName: string;
  readOnly: boolean;
  docKind: DocKind;
};

type ConnectorMessage = {
  text: string;
  tone: "success" | "error";
};

const DOC_PRESETS: Record<DocKind, DocPreset> = {
  word: {
    label: "Word",
    badge: "W",
    fileType: FILE_TYPE.DOCX,
    defaultFileName: "New_Document.docx",
    accept: ".docx,.doc,.docm,.odt,.rtf,.txt",
  },
  excel: {
    label: "Excel",
    badge: "E",
    fileType: FILE_TYPE.XLSX,
    defaultFileName: "New_Spreadsheet.xlsx",
    accept: ".xlsx,.xls,.ods,.csv",
  },
  ppt: {
    label: "PPT",
    badge: "P",
    fileType: FILE_TYPE.PPTX,
    defaultFileName: "New_Presentation.pptx",
    accept: ".pptx,.ppt,.odp",
  },
};

const SubframeTabHost = memo(function SubframeTabHost({
  tabId,
  zodiacId,
  src,
  title,
  onFrame,
  onLoad,
}: {
  tabId: string;
  zodiacId: string;
  src: string;
  title: string;
  onFrame: (tabId: string, frame: HTMLIFrameElement | null) => void;
  onLoad: (tabId: string) => void;
}) {
  const handleFrame = useCallback(
    (frame: HTMLIFrameElement | null) => onFrame(tabId, frame),
    [onFrame, tabId],
  );
  const handleLoad = useCallback(() => onLoad(tabId), [onLoad, tabId]);

  return (
    <iframe
      ref={handleFrame}
      src={src}
      title={title}
      data-onlyoffice-subframe="true"
      data-onlyoffice-instance-id={tabId}
      data-onlyoffice-zodiac={zodiacId}
      className="absolute inset-0 h-full w-full border-0"
      onLoad={handleLoad}
    />
  );
});

function getPreset(docKind: DocKind) {
  return DOC_PRESETS[docKind];
}

function isNewDocument(tab: TabItem) {
  return tab.fileName === getPreset(tab.docKind).defaultFileName;
}

function createTab(slotIndex: number, docKind: DocKind): TabItem {
  const id = nanoid(6);
  const preset = getPreset(docKind);
  const chineseZodiac = getChineseZodiacSlot(slotIndex);
  return {
    id,
    label: `${preset.label} ${slotIndex + 1}`,
    containerId: `tab-editor-${id}`,
    subframeHost: chineseZodiac.id,
    chineseZodiac,
    fileName: preset.defaultFileName,
    readOnly: false,
    docKind,
  };
}

function createInitialTabState() {
  const initialTab = createTab(0, "word");
  return { tabs: [initialTab], activeId: initialTab.id };
}

export function TabsMultiPage({ embedded = false }: { embedded?: boolean }) {
  const [tabs, setTabs] = useState<TabItem[]>([]);
  const [activeId, setActiveId] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [theme, setTheme] = useState<OfficeTheme>(DEFAULT_OFFICE_THEME);
  const [officeXmlEventEnabled, setOfficeXmlEventEnabled] = useState<boolean>(
    OFFICE_XML_EVENT_CONFIG.default.isEnable,
  );
  const [officeXmlLimitMb, setOfficeXmlLimitMb] = useState(
    Math.round(OFFICE_XML_EVENT_CONFIG.default.limitBytes / 1024 / 1024),
  );
  const [resourceState, setResourceState] = useState(() => getDemoResourceState());
  const [cdnOrigin, setCdnOrigin] = useState(resourceState.cdnOrigin);
  const initializedRef = useRef(new Set<string>());
  const connectorsRef = useRef(new Map<string, ConnectorDemo>());
  const openingTabsRef = useRef(new Map<string, Promise<void>>());
  const frameRefs = useRef(new Map<string, HTMLIFrameElement>());
  const managersRef = useRef(new Map<string, SubframeManager>());
  const activeIdRef = useRef(activeId);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const connectorMessageTimerRef = useRef<number | null>(null);
  const [connectorMessage, setConnectorMessage] =
    useState<ConnectorMessage | null>(null);

  activeIdRef.current = activeId;

  const handleSubframeEvent = useCallback(
    (tabId: string, event: string, payload: unknown) => {
      if (event === "loading-change" && activeIdRef.current === tabId) {
        setLoading(!!(payload as { loading?: boolean } | undefined)?.loading);
      }
      if (
        event === "office-xml-size-limit-exceeded" &&
        activeIdRef.current === tabId
      ) {
        setError("文件过大，不支持解析");
      }
    },
    [],
  );

  const handleFrameRef = useCallback(
    (tabId: string, frame: HTMLIFrameElement | null) => {
      if (frame) {
        frameRefs.current.set(tabId, frame);
        managersRef.current.get(tabId)?.attachFrame(frame);
        return;
      }

      frameRefs.current.delete(tabId);
      managersRef.current.get(tabId)?.detachFrame();
    },
    [],
  );

  const handleSubframeLoad = useCallback((tabId: string) => {
    const frame = frameRefs.current.get(tabId);
    if (frame) managersRef.current.get(tabId)?.attachFrame(frame);
  }, []);

  const getTabManager = (tab: TabItem) => {
    const existing = managersRef.current.get(tab.id);
    const frame = frameRefs.current.get(tab.id);
    if (existing) {
      if (frame) existing.attachFrame(frame);
      return existing;
    }
    if (!frame) {
      throw new Error("OnlyOffice subframe is not mounted");
    }

    const preset = getPreset(tab.docKind);
    const manager = new SubframeManager({
      containerId: SUBFRAME_CONTAINER_ID,
      fileType: preset.fileType,
      defaultFileName: preset.defaultFileName,
      readOnly: tab.readOnly,
      theme,
      officeXmlEvent: getOfficeXmlEventConfig(),
      instanceId: tab.id,
      targetOrigin: getSubframeOrigin(tab.subframeHost),
      frame,
      onEvent: (event, payload) =>
        handleSubframeEvent(tab.id, event, payload),
    });
    managersRef.current.set(tab.id, manager);
    return manager;
  };

  const removeTabConnector = (tabId: string) => {
    const connector = connectorsRef.current.get(tabId);
    if (connector?.isConnected) {
      connector.disconnect();
    }
    connectorsRef.current.delete(tabId);
  };

  const replaceTabConnector = (
    tabId: string,
    manager: SubframeManager,
  ) => {
    removeTabConnector(tabId);
    const connector = createConnectorDemo(() => manager);
    connectorsRef.current.set(tabId, connector);
    return connector;
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

  useEffect(
    () =>
      subscribeDemoResourceChange((state) => {
        setResourceState(state);
        setCdnOrigin(state.cdnOrigin);
        setLoading(true);
        connectorsRef.current.forEach((connector) => connector.disconnect());
        connectorsRef.current.clear();
        managersRef.current.forEach((manager) => manager.destroy());
        managersRef.current.clear();
        openingTabsRef.current.clear();
        initializedRef.current.clear();
      }),
    [],
  );

  useEffect(() => {
    const onMessage = (event: MessageEvent) => {
      managersRef.current.forEach((manager) => {
        manager.handleMessage(event);
      });
    };
    window.addEventListener("message", onMessage);
    return () => window.removeEventListener("message", onMessage);
  }, []);

  useEffect(() => {
    const { tabs: initialTabs, activeId: initialActiveId } = createInitialTabState();
    setTabs(initialTabs);
    setActiveId(initialActiveId);
  }, []);

  const activeTab = tabs.find((tab) => tab.id === activeId);
  const activePreset = activeTab ? getPreset(activeTab.docKind) : null;

  const updateTab = (tabId: string, patch: Partial<TabItem>) => {
    setTabs((prev) =>
      prev.map((tab) => (tab.id === tabId ? { ...tab, ...patch } : tab)),
    );
  };

  const getOfficeXmlEventConfig = () => ({
    isEnable: officeXmlEventEnabled,
    limitBytes: officeXmlLimitMb * 1024 * 1024,
  });

  const runAction = async (action: () => Promise<void>, message: string) => {
    try {
      setError(null);
      await action();
    } catch (err) {
      setError(message);
      console.error(message, err);
    }
  };

  const openTabEditor = (tab: TabItem) => {
    const existing = openingTabsRef.current.get(tab.id);
    if (existing) return existing;

    const opening = (async () => {
      const preset = getPreset(tab.docKind);

      const manager = getTabManager(tab);
      manager.configure({
        defaultFileName: preset.defaultFileName,
        readOnly: tab.readOnly,
        theme,
        officeXmlEvent: getOfficeXmlEventConfig(),
      });
      await manager.openDocument({
        fileName: tab.fileName,
        isNew: isNewDocument(tab),
        readOnly: tab.readOnly,
      });

      replaceTabConnector(tab.id, manager);
      initializedRef.current.add(tab.id);
    })();

    openingTabsRef.current.set(tab.id, opening);
    const clearOpening = () => {
      if (openingTabsRef.current.get(tab.id) === opening) {
        openingTabsRef.current.delete(tab.id);
      }
    };
    void opening.then(clearOpening, clearOpening);
    return opening;
  };

  useEffect(() => {
    if (!activeId) return;

    const tab = tabs.find((item) => item.id === activeId);
    if (!tab || initializedRef.current.has(tab.id)) return;

    let cancelled = false;

    openTabEditor(tab)
      .then(() => {
        if (cancelled) return;
        setLoading(false);
      })
      .catch((err) => {
        if (cancelled) return;
        setError("无法加载编辑器");
        setLoading(false);
        console.error("Failed to open tab editor:", err);
      });

    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeId, tabs, resourceState.revision]);

  useEffect(() => {
    return () => {
      connectorsRef.current.forEach((connector) => connector.disconnect());
      connectorsRef.current.clear();
      managersRef.current.forEach((manager) => manager.destroy());
      managersRef.current.clear();
      openingTabsRef.current.clear();
      frameRefs.current.clear();
      initializedRef.current.clear();
      if (connectorMessageTimerRef.current !== null) {
        window.clearTimeout(connectorMessageTimerRef.current);
      }
    };
  }, []);

  const addTab = (docKind: DocKind) => {
    const nextSlotIndex = CHINESE_ZODIAC_SLOTS.findIndex(
      (slot) => !tabs.some((tab) => tab.chineseZodiac.id === slot.id),
    );
    if (nextSlotIndex < 0) {
      setError(`多实例最多支持 ${CHINESE_ZODIAC_SLOT_COUNT} 个固定生肖槽位`);
      return;
    }

    const nextTab = createTab(nextSlotIndex, docKind);
    setTabs((prev) => [...prev, nextTab]);
    setActiveId(nextTab.id);
  };

  const closeTab = (tabId: string) => {
    if (tabs.length <= 1) return;

    const tab = tabs.find((item) => item.id === tabId);
    if (tab) {
      removeTabConnector(tab.id);
      managersRef.current.get(tab.id)?.destroy();
      managersRef.current.delete(tab.id);
      openingTabsRef.current.delete(tab.id);
      frameRefs.current.delete(tab.id);
      initializedRef.current.delete(tab.id);
    }

    setTabs((prev) => {
      const next = prev.filter((item) => item.id !== tabId);
      if (activeId === tabId) {
        setActiveId(next[0]?.id ?? "");
      }
      return next;
    });
  };

  const ensureActiveManager = async () => {
    if (!activeTab) throw new Error("No active tab");

    if (!initializedRef.current.has(activeTab.id)) {
      await openTabEditor(activeTab);
    }

    return getTabManager(activeTab);
  };

  const uploadFile = (file: File) =>
    runAction(async () => {
      if (!activeTab) return;

      const preset = getPreset(activeTab.docKind);

      const manager = getTabManager(activeTab);
      manager.configure({
        defaultFileName: preset.defaultFileName,
        readOnly: activeTab.readOnly,
        theme,
        officeXmlEvent: getOfficeXmlEventConfig(),
      });
      await manager.openDocument({
        fileName: file.name,
        file,
        readOnly: activeTab.readOnly,
      });

      replaceTabConnector(activeTab.id, manager);
      initializedRef.current.add(activeTab.id);
      updateTab(activeTab.id, { fileName: file.name });
    }, "上传失败");

  const newDocument = () =>
    runAction(async () => {
      if (!activeTab) return;

      const preset = getPreset(activeTab.docKind);

      const manager = getTabManager(activeTab);
      manager.configure({
        defaultFileName: preset.defaultFileName,
        readOnly: activeTab.readOnly,
        theme,
        officeXmlEvent: getOfficeXmlEventConfig(),
      });
      await manager.openDocument({
        fileName: preset.defaultFileName,
        isNew: true,
        readOnly: activeTab.readOnly,
      });

      replaceTabConnector(activeTab.id, manager);
      initializedRef.current.add(activeTab.id);
      updateTab(activeTab.id, { fileName: preset.defaultFileName });
    }, "新建失败");

  const exportDocument = () =>
    runAction(async () => {
      const manager = await ensureActiveManager();
      await manager.downloadExport();
    }, "导出失败");

  const printActiveLogs = () =>
    runAction(async () => {
      const manager = await ensureActiveManager();
      manager.printLogs();
    }, "打印日志失败");

  const writeWithConnector = () =>
    runAction(async () => {
      if (!activeTab) throw new Error("No active tab");

      const manager = await ensureActiveManager();
      const logger = manager.getLogger();
      let connector = connectorsRef.current.get(activeTab.id);
      if (!connector) {
        connector = replaceTabConnector(activeTab.id, manager);
      }

      try {
        logger.operation("Connector command started", {
          tabId: activeTab.id,
          fileName: activeTab.fileName,
          fileType: getPreset(activeTab.docKind).fileType,
        });
        const result = await connector.write(
          activeTab.fileName,
          getPreset(activeTab.docKind).fileType,
        );
        logger.operation("Connector command completed", {
          tabId: activeTab.id,
          fileName: activeTab.fileName,
          fileType: result.fileType,
        });
        showConnectorMessage(result.message, "success");
      } catch (error) {
        logger.error("operation", "Connector command failed", {
          tabId: activeTab.id,
          fileName: activeTab.fileName,
          fileType: getPreset(activeTab.docKind).fileType,
          error,
        });
        showConnectorMessage(
          error instanceof Error ? error.message : "Connector command failed",
          "error",
        );
        throw error;
      }
    }, "连接器调用失败");

  const toggleReadOnly = () =>
    runAction(async () => {
      if (!activeTab) return;

      const manager = await ensureActiveManager();
      const nextReadOnly = !activeTab.readOnly;

      await manager.setReadOnly(nextReadOnly);

      updateTab(activeTab.id, { readOnly: nextReadOnly });
    }, "切换模式失败");

  const applyTheme = (nextTheme: OfficeTheme) =>
    runAction(async () => {
      setTheme(nextTheme);

      await Promise.all(
        tabs.map(async (tab) => {
          if (!initializedRef.current.has(tab.id)) return;

          const manager = managersRef.current.get(tab.id);
          if (manager?.isReady()) {
            await manager.setTheme(nextTheme);
          }
        }),
      );
    }, "切换主题失败");

  const loadResource = () =>
    runAction(async () => {
      applyDemoResourceMode("cdn", cdnOrigin);
    }, "切换资源失败");

  const getSubframeSrc = (tab: TabItem) => {
    const url = new URL(
      "/subframe",
      getSubframeOrigin(tab.subframeHost),
    );
    url.searchParams.set("instance", tab.id);
    url.searchParams.set("resourceMode", resourceState.mode);
    url.searchParams.set("resourceRevision", String(resourceState.revision));
    if (resourceState.mode === "cdn") {
      url.searchParams.set("cdnOrigin", resourceState.cdnOrigin);
    }
    return url.href;
  };

  const isActiveDocx = activeTab
    ? getFileExtension(activeTab.fileName, activePreset?.fileType) === "docx"
    : false;

  return (
    <div
      className={`flex flex-col bg-neutral-100 ${
        embedded ? "h-full min-h-0" : "h-screen"
      }`}
    >
      <header className={demoHeaderClass}>
        <div className={demoHeaderInnerClass}>
          <div className="mr-auto min-w-0">
            <h1 className={demoTitleClass}>{embedded ? "示例" : "多实例"}</h1>
            <p className={demoSubtitleClass}>
              {activeTab
                ? `${activePreset?.label} · ${activeTab.fileName}`
                : "切换标签页，实例状态会保留"}
            </p>
          </div>

          <div className={demoToolbarClass}>
            <DemoButton onClick={() => fileInputRef.current?.click()}>
              上传
            </DemoButton>
            <DemoButton onClick={newDocument}>
              新建{activePreset?.label ?? "文档"}
            </DemoButton>
            <DemoButton onClick={exportDocument}>导出</DemoButton>
            <DemoButton onClick={printActiveLogs}>打印日志</DemoButton>
            <DemoButton onClick={writeWithConnector}>连接器写入</DemoButton>
            <DemoButton active={!!activeTab?.readOnly} onClick={toggleReadOnly}>
              {activeTab?.readOnly ? "只读" : "编辑"}
            </DemoButton>
            <DemoMenu label="更多" disabled={loading}>
              <DemoMenuRow>
                <ResourceSwitcher
                  cdnOrigin={cdnOrigin}
                  disabled={loading}
                  onCdnOriginChange={setCdnOrigin}
                  onLoad={loadResource}
                />
              </DemoMenuRow>
              <DemoMenuRow>
                <DemoField label="主题">
                  <DemoSelect
                    value={theme}
                    onChange={(event) =>
                      applyTheme(event.target.value as OfficeTheme)
                    }
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
              {isActiveDocx && (
                <>
                  <DocxCommentsCrud
                    disabled={loading || !!activeTab?.readOnly}
                    getManager={ensureActiveManager}
                    onError={(message, err) => {
                      setError(message);
                      console.error(message, err);
                    }}
                  />
                  <DocxRevisionsCrud
                    disabled={loading || !!activeTab?.readOnly}
                    getManager={ensureActiveManager}
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

        <div className="border-t border-neutral-200/80 bg-[#f5f4f3] px-2 py-1">
          <div className="flex items-end gap-0.5 overflow-x-auto">
            {tabs.map((tab) => {
              const preset = getPreset(tab.docKind);
              const isActive = activeId === tab.id;
              return (
                <div
                  key={tab.id}
                  className={`group relative flex max-w-[200px] min-w-[88px] shrink-0 items-stretch border border-b-0 ${
                    isActive
                      ? "z-10 -mb-px border-neutral-300 bg-white"
                      : "border-transparent hover:bg-white/50"
                  }`}
                >
                  <button
                    type="button"
                    onClick={() => setActiveId(tab.id)}
                    data-onlyoffice-zodiac={tab.chineseZodiac.id}
                    className={`flex min-w-0 flex-1 items-center gap-1.5 px-2.5 py-1.5 text-[12px] ${
                      isActive ? "text-neutral-900" : "text-neutral-500"
                    }`}
                    title={`${tab.chineseZodiac.name} · ${tab.fileName}`}
                  >
                    <span
                      className="shrink-0 text-[16px] leading-none"
                      role="img"
                      aria-label={tab.chineseZodiac.name}
                    >
                      {tab.chineseZodiac.emoji}
                    </span>
                    <span className="sr-only">{tab.chineseZodiac.name}</span>
                    <span className="shrink-0 text-[11px] text-neutral-400">
                      {preset.badge}
                    </span>
                    <span className="truncate">{tab.label}</span>
                  </button>
                  {tabs.length > 1 && (
                    <button
                      type="button"
                      onClick={() => closeTab(tab.id)}
                      className={`mr-1 self-center rounded px-1 text-[10px] transition-colors ${
                        isActive
                          ? "text-gray-400 hover:text-gray-600"
                          : "text-gray-300 opacity-0 hover:text-gray-500 group-hover:opacity-100"
                      }`}
                      aria-label={`关闭 ${tab.label}`}
                    >
                      ×
                    </button>
                  )}
                </div>
              );
            })}

            <div className="mb-px flex shrink-0 items-center gap-1 pl-1.5">
              {(Object.keys(DOC_PRESETS) as DocKind[]).map((kind) => {
                const preset = getPreset(kind);
                return (
                  <button
                    key={kind}
                    type="button"
                    onClick={() => addTab(kind)}
                    disabled={tabs.length >= CHINESE_ZODIAC_SLOT_COUNT}
                    className="inline-flex h-7 items-center border border-dashed border-neutral-300 bg-transparent px-2 text-[12px] text-neutral-600 hover:border-neutral-400 hover:bg-white"
                    title={
                      tabs.length >= CHINESE_ZODIAC_SLOT_COUNT
                        ? `已达到 ${CHINESE_ZODIAC_SLOT_COUNT} 个生肖槽位上限`
                        : `新建 ${preset.label} 标签页`
                    }
                  >
                    + {preset.label}
                  </button>
                );
              })}
            </div>
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
        <div className="mx-4 mt-4 rounded border-l-4 border-red-500 bg-red-50 p-3 text-sm text-red-700">
          {error}
        </div>
      )}

      <div className="relative min-h-0 flex-1 bg-white">
        {tabs.length === 0 && (
          <div className="absolute inset-0 flex items-center justify-center text-sm text-gray-400">
            加载中...
          </div>
        )}
        {tabs.map((tab) => (
          <div
            key={tab.id}
            className={`absolute inset-0 ${
              activeId === tab.id ? "visible z-10" : "invisible z-0"
            }`}
          >
            <SubframeTabHost
              tabId={tab.id}
              zodiacId={tab.chineseZodiac.id}
              src={getSubframeSrc(tab)}
              title={`${getPreset(tab.docKind).label} ${tab.label} 编辑器`}
              onFrame={handleFrameRef}
              onLoad={handleSubframeLoad}
            />
          </div>
        ))}

        {loading && (
          <div className="absolute inset-0 z-20 flex items-center justify-center bg-white/70 backdrop-blur-sm">
            <div className="rounded-lg border border-gray-200 bg-white px-4 py-3 text-sm text-gray-700 shadow">
              加载中...
            </div>
          </div>
        )}
      </div>

      <input
        ref={fileInputRef}
        type="file"
        accept={activePreset?.accept ?? ".docx,.doc,.odt,.rtf,.txt"}
        className="hidden"
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) {
            uploadFile(file);
            if (fileInputRef.current) fileInputRef.current.value = "";
          }
        }}
      />
    </div>
  );
}
