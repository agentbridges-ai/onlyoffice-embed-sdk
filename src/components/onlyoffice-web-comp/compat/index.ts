/**
 * onlyoffice-browser-compatible migration surface.
 *
 * This entry intentionally lives apart from the native manager API. It keeps
 * the 0.3.34 editor contract while using embed-sdk's hardened direct editor,
 * x2t, lifecycle, cross-origin bridge, and static-resource implementations.
 */
export * from "./editor";
export * from "./runtime-resources";
export * from "./version";
