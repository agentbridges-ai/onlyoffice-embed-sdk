import type { FileType, OnlyOfficeConnector } from "@/components/onlyoffice-embed-sdk";
import { getFileExtension } from "./office-formats";
import type { MultiInstanceManager } from "./subframe-manager";

export type ConnectorDemo = {
  readonly isConnected: boolean;
  disconnect: () => void;
  write: (fileName: string, fallback: FileType) => Promise<{
    fileType: FileType;
    message: string;
  }>;
};

function getActiveConnectorFileType(fileName: string, fallback: FileType) {
  const extension = getFileExtension(fileName, fallback.toLowerCase());
  if (["xlsx", "xls", "ods", "csv"].includes(extension)) return "XLSX";
  if (["docx", "doc", "docm", "odt", "rtf", "txt"].includes(extension)) {
    return "DOCX";
  }
  if (["pptx", "ppt", "odp"].includes(extension)) return "PPTX";
  return fallback;
}

function getConnectorCommand(fileType: FileType) {
  if (fileType === "XLSX") {
    return new Function(
      'Api.GetActiveSheet().GetRange("A1").SetValue("[Connector] wrote this cell.");',
    ) as () => void;
  }
  if (fileType === "DOCX") {
    return new Function(
      'const paragraph = Api.CreateParagraph(); paragraph.AddText("[Connector] wrote this paragraph."); Api.GetDocument().Push(paragraph);',
    ) as () => void;
  }
  return new Function(
    'const fill = Api.CreateSolidFill(Api.CreateRGBColor(230, 247, 255)); const stroke = Api.CreateStroke(0, Api.CreateSolidFill(Api.CreateRGBColor(24, 144, 255))); const shape = Api.CreateShape("rect", 7200000, 900000, fill, stroke); shape.SetPosition(1000000, 1000000); shape.GetDocContent().GetContent()[0]?.AddText("[Connector] wrote this text box."); Api.GetPresentation().GetCurrentSlide().AddObject(shape);',
  ) as () => void;
}

function getConnectorSuccessMessage(fileType: FileType) {
  if (fileType === "XLSX") return "Connector: wrote to A1";
  if (fileType === "PPTX") return "Connector: added a text box";
  return "Connector: inserted a paragraph";
}

/**
 * 与评注/修订组件一致：通过业务侧提供的 getManager 获取当前编辑器实例。
 * Automation API 仅存在于 callCommand 传入的函数源码中，不污染宿主页面全局作用域。
 */
export function createConnectorDemo(
  getManager: () => MultiInstanceManager,
): ConnectorDemo {
  const manager = getManager();
  const connector: OnlyOfficeConnector = manager
    .getEditor()
    .createConnector({ autoconnect: false });

  return {
    get isConnected() {
      return connector.isConnected;
    },
    disconnect() {
      if (connector.isConnected) connector.disconnect();
    },
    async write(fileName, fallback) {
      const fileType = getActiveConnectorFileType(fileName, fallback);
      if (!connector.isConnected) connector.connect();
      await new Promise<void>((resolve) => {
        connector.callCommand(getConnectorCommand(fileType), () => {
          resolve();
        });
      });
      return { fileType, message: getConnectorSuccessMessage(fileType) };
    },
  };
}
