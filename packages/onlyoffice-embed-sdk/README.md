# @agentbridges-ai/onlyoffice-embed-sdk

ESM-only browser SDK for embedding ONLYOFFICE editors, running x2t WebAssembly
conversion, and managing editor lifecycle without a Document Server.

## Install

```sh
pnpm add @agentbridges-ai/onlyoffice-embed-sdk
```

The npm package contains JavaScript, TypeScript declarations, and a
self-contained x2t Worker. It deliberately does not contain the large
ONLYOFFICE SDK, x2t WASM, or font files.

## Native API

Register runtime assets before creating the first editor, then use the native
manager API:

```ts
import {
  FILE_TYPE,
  OnlyOfficeManager,
  registerOnlyOfficeStaticResource,
} from "@agentbridges-ai/onlyoffice-embed-sdk";

registerOnlyOfficeStaticResource({
  cdnOrigin: "https://onlyoffice-embed-resource.pages.dev",
  onlyofficeVersion: "9.4.0-develop",
  // SHA-256 of your immutable deployment manifest (recommended).
  assetManifestDigest: "<64 lowercase hex characters>",
});

const manager = await OnlyOfficeManager.createWithFile(
  {
    containerId: "office-editor",
    fileType: FILE_TYPE.DOCX,
    defaultFileName: "document.docx",
  },
  file,
);
```

Without `cdnOrigin`, editor UI assets are read from
`/packages/onlyoffice/9.4.0-develop` and the converter from the immutable
`/packages/onlyoffice/x2t/v9.3.0+4` path on the consuming application's origin.
`onlyofficeVersion` never selects the converter build. Import
`ONLYOFFICE_X2T_RELEASE` from `/compat` to inspect the signed source commit,
toolchain, raw/compressed asset hashes, and real DOC/Pivot regression identity.

The x2t Worker is inlined so installed packages work after both Vite and Bun
browser builds. A restrictive Content Security Policy must allow
`worker-src blob:` (or the equivalent fallback through `child-src`).

## onlyoffice-browser compatibility API

Existing applications can migrate imports to the explicit compatibility
entry:

```ts
import {
  createOfficeEditor,
  mountOfficeEditor,
  createOfficeRuntimeResourceManager,
  type OfficeEditorInstance,
  type OfficeEditorMount,
  type OfficeHostIdentity,
  type OfficeHostUrlContext,
  type OfficeHostUrlResolver,
} from "@agentbridges-ai/onlyoffice-embed-sdk/compat";
```

The facade covers the `onlyoffice-browser@0.3.34` editor contract: File, Blob,
buffer, URL, and empty-document inputs; mount/activate/destroy; save and x2t
conversion; read-only and theme changes; state callbacks; plugin invocation;
and host identity checks. Plugin messages keep the
`onlyoffice-browser-plugin/v1` protocol and are bound to the exact editor
frame, origin, source window, plugin instance, GUID, and request.

This is a direct-embed adapter, not the old independent `office-host.html`
runtime. The following differences are intentional:

- `hostUrl` is resolved for source compatibility and mount state, but no outer
  host is navigated. Static assets come from `registerOnlyOfficeStaticResource`.
- `getHostIdentity()` reports the direct-embed package/resource identity. An
  identity captured from `onlyoffice-browser` is rejected with
  `OfficeHostIdentityMismatchError` and must be regenerated. For a release
  process assertion, register the deployed manifest's SHA-256 as
  `assetManifestDigest`. This value is trusted caller input: the SDK validates
  its hex format and compares expected/actual identity values, but does not
  fetch that manifest or verify any remote asset bytes. Without it, the
  fallback digest covers only the package version and resource URL
  coordinates.
- `spellcheck` is forwarded to the native editor customization while runtime
  language changes remain instance-scoped.
- Native UI Save Copy As and Download As results are converted to `File`
  objects and delivered to `onSaveAs` and `onDownload`, respectively. When a
  callback is absent, the facade falls back to a browser download, matching
  the old host's persistence contract.
- In CDN mode, application-owned plugin manifests keep their exact,
  application-origin `.../config.json` URLs so ONLYOFFICE can derive and
  validate the plugin origin. The child forwards cross-origin config reads,
  but the parent fetches only the per-instance exact allowlist; relative
  plugin entry/icon URLs stay relative to that validated manifest directory,
  and the general HTTP bridge remains denied.
- The resource manager streams registered HTTP resources. Transactional
  offline install, per-font install, planning, repair, update, pause, and
  resume operations throw `OfficeRuntimeResourceCompatibilityError`.
- `confirmSaveToNewFormat()` uses the browser confirmation dialog.
- The SDK does not inject or port a custom preview/edit toggle into upstream
  ONLYOFFICE bundles. `mode` and `setReadonly()` only select the native
  desktop/embedded configuration; host applications own any mode controls.

Integrations that require host-origin isolation must use
`@agentbridges-ai/onlyoffice-embed-sdk/compat/subframe`. The root and `/compat`
entries are direct-embed APIs and intentionally do not create the outer
cross-origin iframe.

### Isolated compatibility subframes

Applications that cannot run the editor in their own window should use the
formal cross-origin entry. It preserves the `0.3.34` mount/instance shape while
placing every editor in one of the hosted zodiac origins:

```ts
import {
  createOfficeEditor,
  getOfficeSubframeOrigin,
  HOSTED_COMPAT_SUBFRAME_IDENTITY,
} from "@agentbridges-ai/onlyoffice-embed-sdk/compat/subframe";

const editor = await createOfficeEditor(container, {
  hostUrl: getOfficeSubframeOrigin("rat"),
  expectedHostIdentity: HOSTED_COMPAT_SUBFRAME_IDENTITY,
  file,
  fileName: file.name,
  onSave: async (output) => {
    await persist(output);
    return true;
  },
});
```

The fixed slots are `rat`, `ox`, `tiger`, `rabbit`, `dragon`, `snake`,
`horse`, `goat`, `monkey`, `rooster`, `dog`, and `pig`. Local development uses
`http://<slot>.onlyoffice.localhost:<port>`; production uses
`https://<slot>.onlyoffice.agent-bridges.com`. The package creates
`/subframe?runtime=compat` and binds every message to the exact iframe window,
origin, instance ID, and random session token. It never transfers an
`EditorManager`, `DocsAPI`, or `WindowProxy` across that boundary.

The outer iframe remains on its zodiac origin, while static editor and x2t
requests use one canonical resource origin so all slots share the browser HTTP
cache. Production defaults to `https://onlyoffice.agent-bridges.com`; local
development derives `http://onlyoffice.localhost:<port>`. `resourceOrigin`
may override that value only with one of those canonical roots. Hashed bundles,
the hosted build path, and tagged x2t files are immutable for one year;
`/subframe` and `/api/version` remain `no-store`.

Default initialization loads `api.js` directly and does not create a hidden
`iframe[data-onlyoffice-preload]` on the canonical origin. This keeps the
shared HTTP-cache benefit without adding one redundant canonical browsing
context inside every zodiac host. Applications that deliberately want the
upstream preload document may opt in by calling `preloadOnlyOffice()` before
initialization; it is not part of the hosted compatibility path.

File, Blob, ArrayBuffer, Uint8Array, URL, and empty-document inputs are
supported. URL inputs are fetched in the parent window with `fetchOptions`
before a copied buffer is transferred. The proxy keeps the existing
`invokePlugin`, `save`, `confirmSaveToNewFormat`, read-only/theme controls,
callbacks, state, identity, and destroy APIs. It additionally exposes
`setLanguage`, `saveAs`, `download`, and `print`; print synchronously reserves
the parent-window print surface, exports a PDF without marking the current
document persisted, invokes the browser print flow in that parent surface, and
returns the same `File`. Native menu output still arrives through `onSaveAs`
and `onDownload`.

Each active editor needs a different slot origin. A slot can be reused only
after its previous mount has been destroyed. Hosted subframes and all static
resource responses include `Origin-Agent-Cluster: ?1`.

The hosted `HOSTED_COMPAT_SUBFRAME_IDENTITY` is published by
`https://onlyoffice.agent-bridges.com/api/version`. Its digest is the SHA-256
of the canonical `runtimeManifest` JSON returned by that endpoint, so a
consumer can verify the package version, host build, protocol, route, and
ONLYOFFICE resource version as one release coordinate. This hosted digest is
not a byte-by-byte audit of every CDN asset.

### Migrating the release identity

An `onlyoffice-browser@0.3.34` release identity is deliberately incompatible
with this direct-embed runtime. Do not copy its `packageVersion`,
`hostBuildId`, or digest into `expectedHostIdentity`.

If the release process uses an immutable runtime-asset manifest, produce one
canonical JSON manifest containing the path, byte size, and SHA-256 of every
deployed ONLYOFFICE SDK and font asset. The hosted SDK identity additionally
binds the independently signed x2t release tag, source commit, and compressed
and raw WASM digests. Serialize that manifest
once, deploy the exact bytes beside the assets, and hash those bytes without
parsing or reformatting them. The SDK does not fetch or interpret this file;
the deployment pipeline is responsible for validating its entries and for
preventing assets from changing under a retained digest:

```sh
node -e 'const fs=require("node:fs"),c=require("node:crypto");const b=fs.readFileSync(process.argv[1]);process.stdout.write(c.createHash("sha256").update(b).digest("hex")+"\n")' ./onlyoffice-runtime-manifest.json
```

Pin the resulting lowercase 64-character digest in the consuming
application's release manifest. For SDK `0.3.2`, the expected identity is:

```json
{
  "packageVersion": "0.3.2",
  "hostBuildId": "onlyoffice-embed-sdk-hosted-v4",
  "assetManifestDigest": "<sha256-of-the-exact-deployed-manifest-bytes>"
}
```

Register the same digest before creating any resource manager, editor, or x2t
converter, and pass the pinned object to the compatibility facade:

```ts
import {
  createOfficeEditor,
  createOfficeRuntimeResourceManager,
} from "@agentbridges-ai/onlyoffice-embed-sdk/compat";
import releaseIdentity from "./onlyoffice-release-identity.json";

await createOfficeRuntimeResourceManager({
  cdnOrigin: "https://onlyoffice-embed-resource.pages.dev",
  onlyofficeVersion: "9.4.0-develop",
  assetManifestDigest: releaseIdentity.assetManifestDigest,
});

const container = document.querySelector<HTMLElement>("#office-editor")!;
const editor = await createOfficeEditor(container, {
  // Retained for the 0.3.34 call shape; direct embed does not navigate it.
  hostUrl: "https://host.example/office-host.html",
  expectedHostIdentity: releaseIdentity,
  // file/buffer/url/emptyType and callbacks...
});
```

`resolveOfficeEmbedHostIdentity()` must then equal the pinned object. This
detects a stale or mismatched application release identity; it does not prove
that the manifest exists or that the CDN serves the bytes described by it.
Changing an asset URL or its contents while reusing the registered digest is
not detected by the SDK, so the release pipeline and immutable hosting policy
must enforce that binding. While an application continues to pass its old
0.3.34 identity as `expectedHostIdentity`, activation fails with
`OfficeHostIdentityMismatchError`; omitting `expectedHostIdentity` disables
this equality check.

## Build and verify

From the repository root:

```sh
pnpm typecheck:sdk
pnpm build:sdk
pnpm verify:sdk
```

`verify:sdk` performs two clean SDK build-and-pack runs, checks byte-for-byte
tarball reproducibility and the publish allowlist, installs the actual tarball
outside the workspace, and validates TypeScript, Node SSR, Vite browser/SSR,
and Bun imports/builds.

Release tags use stable versions only: `sdk-v<package-version>`, for example
`sdk-v0.3.2`. The protected release workflow requires a GitHub-verified signed
annotated tag on `main`, approval through the `npm-production` environment,
and npm trusted publishing. It publishes the exact verified tarball with OIDC
provenance; direct manual publication is not a supported release path.

Prerelease versions are intentionally rejected; add an explicit dist-tag
policy before enabling them.

License: AGPL-3.0-only.
