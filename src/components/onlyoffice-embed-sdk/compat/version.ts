/** Version of the independently published compatibility package. */
export const ONLYOFFICE_EMBED_SDK_VERSION = "0.2.0";

/** Identifies the direct-embed host model used by the compatibility facade. */
export const ONLYOFFICE_EMBED_HOST_BUILD_ID =
  "onlyoffice-embed-sdk-direct-v1";

/** Canonical, deployment-visible coordinates of the hosted compat runtime. */
export const ONLYOFFICE_EMBED_HOST_MANIFEST = Object.freeze({
  packageVersion: ONLYOFFICE_EMBED_SDK_VERSION,
  hostBuildId: ONLYOFFICE_EMBED_HOST_BUILD_ID,
  compatSubframeProtocol: 1,
  compatSubframePath: "/subframe?runtime=compat",
  onlyofficeVersion: "9.4.0-develop",
});

/** SHA-256 of JSON.stringify(ONLYOFFICE_EMBED_HOST_MANIFEST). */
export const ONLYOFFICE_EMBED_HOST_ASSET_DIGEST =
  "08d22b63478f418488c67356842455ea7bcf040ddecbac9c6b0c3d72db4b0dbe";
