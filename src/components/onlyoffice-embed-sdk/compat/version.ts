/** Version of the independently published compatibility package. */
export const ONLYOFFICE_EMBED_SDK_VERSION = "0.4.3";

/** Identifies the direct-embed host model used by the compatibility facade. */
export const ONLYOFFICE_EMBED_HOST_BUILD_ID =
  "onlyoffice-embed-sdk-hosted-v9";

/** Immutable, independently attested converter used by the hosted runtime. */
export const ONLYOFFICE_X2T_RELEASE = Object.freeze({
  repository: "agentbridges-ai/onlyoffice-x2t-wasm",
  tag: "v9.3.0+4",
  sourceCommit: "5790fda684ac1a837f1ab92fcaaf4a0de9ec4ec1",
  releaseUrl:
    "https://github.com/agentbridges-ai/onlyoffice-x2t-wasm/releases/tag/v9.3.0+4",
  toolchain: Object.freeze({
    onlyofficeCore: "v9.3.0.140",
    emscripten: "4.0.11",
    linkerFlags: Object.freeze([
      "-sALLOW_MEMORY_GROWTH",
      "-sEMULATE_FUNCTION_POINTER_CASTS=1",
      "-sEXPORTED_FUNCTIONS=_main1",
    ]),
  }),
  files: Object.freeze({
    scriptSha256:
      "e2f79cf6b71d10a7b5db695dcb4043afb1c5e0412520201bf71d6f12f97648d7",
    scriptBrotliSha256:
      "e8f1ebbe466ffc2553c93a0a9c59a93099d505b8e477c543a28e8af8f1c702a7",
    wasmRawSha256:
      "68d225de76693d0341531b44fc24a2f9a8e0f06aa4c3d903399cb32a8008d5e7",
    wasmBrotliSha256:
      "b0701d6c3e7708ab297649586adf7a52230aed0fef5998eaa8b35abb5c7424b5",
  }),
  regressions: Object.freeze({
    legacyDoc: Object.freeze({
      inputSha256:
        "d85e44ae5368ccbbe57ded8533ced05a250c30cfa15da10f19fdaf63f080238c",
      outputSha256:
        "074a9b350ff6a6e1ee32866c03416a0682c05635cdb8f3f60b6e4a02eaad9a2a",
      outputSize: 132_030,
      header: "DOCY;v5;",
    }),
    pivotSlicer: Object.freeze({
      inputSha256:
        "ffecc0a33c9e41b392fbee30127a97f3e5c3577c717be103471460bd07c2ec58",
      outputSha256:
        "c40fb3f4f67311426110d4786eb4684981aec9ed05b4f13c6367c8470de4d89e",
      outputSize: 85_138,
      header: "XLSY;v2;",
    }),
  }),
});

/** Canonical, deployment-visible coordinates of the hosted compat runtime. */
export const ONLYOFFICE_EMBED_HOST_MANIFEST = Object.freeze({
  packageVersion: ONLYOFFICE_EMBED_SDK_VERSION,
  hostBuildId: ONLYOFFICE_EMBED_HOST_BUILD_ID,
  compatSubframeProtocol: 1,
  compatSubframePath: "/subframe?runtime=compat",
  canonicalResourceOrigin: "https://onlyoffice.agent-bridges.com",
  editorFrameOrigin: "zodiac-slot",
  staticAssetOrigin: "canonical",
  onlyofficeResourcePath:
    "/onlyoffice/runtime/onlyoffice-embed-sdk-hosted-v9",
  onlyofficeVersion: "9.4.0-develop",
  interfaceThemes: Object.freeze(["theme-white", "theme-night"]),
  interfaceThemeUpdate: "native-in-place",
  printTarget: "single-pdf-frame",
  x2t: ONLYOFFICE_X2T_RELEASE,
});

/** SHA-256 of JSON.stringify(ONLYOFFICE_EMBED_HOST_MANIFEST). */
export const ONLYOFFICE_EMBED_HOST_ASSET_DIGEST =
  "f23dca89bd7828cede0885b38b9e79fb6fdba91e93cef4c79de9dbc08153ca11";
