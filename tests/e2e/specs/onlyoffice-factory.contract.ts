export type ResourceMode = "local" | "cdn";

export type StepResult = {
  name: string;
  status: "running" | "passed" | "failed";
  detail?: string;
};

export type ScenarioResult = {
  mode: ResourceMode;
  status: "idle" | "running" | "passed" | "failed";
  steps: StepResult[];
  error?: string;
};

export const ONLYOFFICE_FACTORY_EXPECTED_STEPS = [
  "resource mode",
  "CDN Nexolyra plugin READY",
  "manager factory concurrent open",
  "manager factory open/get",
  "manager facade from factory",
  "manager connector",
  "manager factory destroy",
  "fixture manifest",
  "x2t edge imports",
  "x2t legacy DOC and Pivot/Slicer",
  "generated negative fixtures",
  "manager create",
  "manager createWithFile",
  "manager legacy DOC OLE preview",
  "manager spreadsheet connector",
  "manager PivotTable and Slicer model",
  "text fallback files",
  "manager fromEditor",
  "manager factory destroyAll",
] as const;
