# @agentbridges-ai/onlyoffice-embed-sdk

ESM-only browser SDK for embedding ONLYOFFICE editors, running x2t WebAssembly
conversion, and managing editor lifecycle without a Document Server.

## Install

After the first protected npm release:

```sh
pnpm add @agentbridges-ai/onlyoffice-embed-sdk
```

Until then, use the repository's verified `.tgz` build artifact; the nested
workspace package is not a supported Git dependency.

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

Without `cdnOrigin`, assets are read from
`/packages/onlyoffice/9.4.0-develop` on the consuming application's origin.

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
- `spellcheck` is accepted for source compatibility while the hardened embed
  runtime keeps its own spellcheck policy.
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

New integrations should prefer the native root entry. The `/compat` entry is
for controlled migration.

### Migrating the release identity

An `onlyoffice-browser@0.3.34` release identity is deliberately incompatible
with this direct-embed runtime. Do not copy its `packageVersion`,
`hostBuildId`, or digest into `expectedHostIdentity`.

If the release process uses an immutable runtime-asset manifest, produce one
canonical JSON manifest containing the path, byte size, and SHA-256 of every
deployed ONLYOFFICE SDK, x2t, WASM, and font asset. Serialize that manifest
once, deploy the exact bytes beside the assets, and hash those bytes without
parsing or reformatting them. The SDK does not fetch or interpret this file;
the deployment pipeline is responsible for validating its entries and for
preventing assets from changing under a retained digest:

```sh
node -e 'const fs=require("node:fs"),c=require("node:crypto");const b=fs.readFileSync(process.argv[1]);process.stdout.write(c.createHash("sha256").update(b).digest("hex")+"\n")' ./onlyoffice-runtime-manifest.json
```

Pin the resulting lowercase 64-character digest in the consuming
application's release manifest. For SDK `0.1.2`, the expected identity is:

```json
{
  "packageVersion": "0.1.2",
  "hostBuildId": "onlyoffice-embed-sdk-direct-v1",
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
`sdk-v0.1.2`. Before the first real release, an organization owner must:

1. Publish a minimal `@agentbridges-ai/onlyoffice-embed-sdk@0.0.0` placeholder
   with `npm publish --access public --tag bootstrap`. Do not manually publish
   the real release version.
2. Configure npm trusted publishing for this repository, the
   `.github/workflows/publish-npm.yml` workflow, and the `npm-production`
   environment.
3. Create that GitHub environment with required reviewers, and protect `main`
   plus `sdk-v*` tag creation with repository rules.
4. Push a GitHub-verified signed annotated `sdk-v0.1.2` tag. The workflow then
   verifies and publishes the exact checked tarball with npm OIDC provenance.

Prerelease versions are intentionally rejected; add an explicit dist-tag
policy before enabling them.

License: AGPL-3.0-only.
