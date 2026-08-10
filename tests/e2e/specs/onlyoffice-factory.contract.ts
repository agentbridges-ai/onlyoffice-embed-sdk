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
  "manager factory concurrent open",
  "manager factory open/get",
  "manager facade from factory",
  "manager connector",
  "manager factory destroy",
  "fixture manifest",
  "x2t edge imports",
  "generated negative fixtures",
  "manager create",
  "manager createWithFile",
  "manager spreadsheet connector",
  "text fallback files",
  "manager fromEditor",
  "manager factory destroyAll",
] as const;
