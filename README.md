# OnlyOffice Embed SDK

> 📖 English | [中文](README.zh.md)

🌐 **Live Demo**: https://onlyoffice.agent-bridges.com/

A browser-based document solution built on the OnlyOffice static SDK. View, edit, and convert Word, Excel, and PowerPoint entirely on the client—**no Document Server required**.

This repository has two parts:

| Part | Path | Description |
|------|------|-------------|
| **Publishable SDK** | [`packages/onlyoffice-embed-sdk/`](packages/onlyoffice-embed-sdk/) | `@agentbridges-ai/onlyoffice-embed-sdk` package, build, and migration guide |
| **Shared implementation** | [`src/components/onlyoffice-embed-sdk/`](src/components/onlyoffice-embed-sdk/) | Editor, x2t, bridge, compatibility facade, and Markdown docs |
| **Demo site** | [`src/app/`](src/app/) + [`src/features/`](src/features/) | TanStack Start file routes on a Cloudflare Worker |

## Project Positioning

The repository contains a private demo/deployment application and a separate, publishable ESM package named `@agentbridges-ai/onlyoffice-embed-sdk`. The package bundles the editor lifecycle, hardened bridge, and x2t Worker code, while the large ONLYOFFICE SDK, x2t WASM, and font files remain separately deployed runtime assets.

Use this project when you want to embed OnlyOffice editing into your own web app without running OnlyOffice Document Server. The demo site is intentionally part of the repository so you can copy a working integration instead of reconstructing the editor lifecycle from scattered snippets.

## Integrating into Your Project

The package integration path is:

1. Install `@agentbridges-ai/onlyoffice-embed-sdk` and import its native manager API. Applications migrating from the independent-host API use `/compat/subframe`, which mounts the formal isolated runtime on one of the 12 fixed origins. The direct `/compat` entry is reserved for controlled same-window integrations.
2. Host [`public/packages/onlyoffice/`](public/packages/onlyoffice/) on the application origin or a CDN, then call `registerOnlyOfficeStaticResource` before creating an editor.
3. Build your UI by following [`src/features/demo/office-preview-page.tsx`](src/features/demo/office-preview-page.tsx): create an editor container, keep an `OnlyOfficeManager` instance, call `openDocument`, `downloadExport`, `toggleReadOnly`, and destroy the manager on unmount. To call the editor Automation API from the parent page, get a Developer Edition Connector with `createConnector()`.

The hosted compatibility entry serves its own runtime resources; the parent application does not copy `public/packages`. Package-specific installation, CSP, compatibility, and release details are in [`packages/onlyoffice-embed-sdk/README.md`](packages/onlyoffice-embed-sdk/README.md).

Static resource resolution is centralized in [`src/components/onlyoffice-embed-sdk/const/index.ts`](src/components/onlyoffice-embed-sdk/const/index.ts). The editor UI uses the Developer Edition Docker-exported 9.4 SDK at `/packages/onlyoffice/9.4.0-develop`; the converter is independently pinned to a signed `agentbridges-ai/onlyoffice-x2t-wasm` release under `/packages/onlyoffice/x2t/<tag>`. `onlyofficeVersion` changes only the editor UI resource directory.

## Core Advantages

- **Local processing**: Documents stay in the browser
- **Format support**: Word, Excel, PowerPoint, CSV, DOCM, and more
- **No backend**: Host static SDK assets only
- **Engineering APIs**: Read-only/edit toggle, theme, language, multi-instance isolation
- **Connector support**: Call the editor Automation API from the parent page with the Developer Edition Connector; works in both local and cross-origin CDN modes

## Quick Try

1. Visit the [live demo](https://onlyoffice.agent-bridges.com/) or run locally:

```bash
git clone <repository-url>
cd onlyoffice-embed-sdk
pnpm install
pnpm dev
# http://localhost:3001
```

2. Open a route:

| Route | Description |
|-------|-------------|
| `/` | Product landing page |
| `/docs` | OnlyOffice Embed SDK documentation (rendered from Markdown) |
| `/docs/demos/single` | Single-instance editor demo |
| `/docs/demos/multi` | Multi-instance Tab demo |
| `/api/version` | Current installable SDK version (JSON) |

3. Upload a file → edit → export

Legacy route `/examples` redirects to the single-instance demo; `/multi` redirects to the multi-instance demo.

## Version API

The landing-page version badge reads the public endpoint directly instead of
embedding a separate display value:

```bash
curl https://onlyoffice.agent-bridges.com/api/version
```

```json
{
  "name": "@agentbridges-ai/onlyoffice-embed-sdk",
  "version": "0.3.0",
  "release": "sdk-v0.3.0",
  "hostIdentity": {
    "packageVersion": "0.3.0",
    "hostBuildId": "onlyoffice-embed-sdk-hosted-v2",
    "assetManifestDigest": "becf9ed215cab2269745d330ad42e2a9e556845e7702d46a64a3de6c565c239b"
  },
  "runtimeManifest": {
    "packageVersion": "0.3.0",
    "hostBuildId": "onlyoffice-embed-sdk-hosted-v2",
    "compatSubframeProtocol": 1,
    "compatSubframePath": "/subframe?runtime=compat",
    "onlyofficeVersion": "9.4.0-develop",
    "x2t": {
      "repository": "agentbridges-ai/onlyoffice-x2t-wasm",
      "tag": "v9.3.0+4",
      "sourceCommit": "5790fda684ac1a837f1ab92fcaaf4a0de9ec4ec1"
    }
  }
}
```

The endpoint supports `GET`, `HEAD`, and cross-origin queries. It sends
`Cache-Control: no-store` so consumers receive the currently deployed version.
The `x2t` object above shows its identity coordinates; the live response also
includes the pinned toolchain, raw/compressed asset hashes, and real DOC/Pivot
regression digests.
`hostIdentity.assetManifestDigest` is the SHA-256 of the canonical
`runtimeManifest` JSON bytes, allowing consumers to pin the deployed protocol
coordinates returned by the same endpoint. It is not a byte-by-byte audit of
the CDN assets.

## OnlyOffice Embed SDK Docs

**API details live in the OnlyOffice Embed SDK docs**, not duplicated here.

- **Entry**: [OnlyOffice Embed SDK README (English)](src/components/onlyoffice-embed-sdk/readme.md)
- **Overview**: [docs/概述.md](src/components/onlyoffice-embed-sdk/docs/概述.md)

| Doc | Topic |
|-----|-------|
| [Quick Start](src/components/onlyoffice-embed-sdk/docs/快速开始.md) | Init and container mount |
| [Core API](src/components/onlyoffice-embed-sdk/docs/核心API.md) | `OnlyOfficeManager`, multi-instance |
| [Events](src/components/onlyoffice-embed-sdk/docs/事件系统.md) | EventBus |
| [Examples](src/components/onlyoffice-embed-sdk/docs/完整示例.md) | React integration patterns |
| [Reference](src/components/onlyoffice-embed-sdk/docs/API参考.md) | Constants and types |
| [Notes & Formats](src/components/onlyoffice-embed-sdk/docs/注意事项与支持格式.md) | Prerequisites and formats |
| [Fonts](src/components/onlyoffice-embed-sdk/docs/字体配置.md) | Custom font registration |
| [Comments & Revisions](src/components/onlyoffice-embed-sdk/docs/批注修订与-Word-API.md) | Comments and revisions |
| [Single-instance Demo](src/components/onlyoffice-embed-sdk/docs/单实例示例.md) | Single editor demo + source |
| [Multi-instance Demo](src/components/onlyoffice-embed-sdk/docs/多实例示例.md) | Full Tab demo source |

```typescript
import {
  OnlyOfficeManager,
  FILE_TYPE,
  ONLYOFFICE_ID,
} from "@agentbridges-ai/onlyoffice-embed-sdk";
```

## Project Structure

```
onlyoffice-embed-sdk/
├── packages/
│   └── onlyoffice-embed-sdk/             # Publishable ESM package
├── src/
│   ├── app/                              # TanStack Start file routes
│   │   ├── __root.tsx                    # HTML shell and global metadata
│   │   ├── index.tsx                     # Landing page
│   │   ├── docs/                         # Documentation routes
│   │   │   ├── index.tsx                  # /docs overview
│   │   │   ├── $slug.tsx                  # /docs/* Markdown pages
│   │   │   └── demos/                    # /docs/demos/single|multi
│   │   └── examples.tsx                  # → redirect to single demo
│   ├── features/
│   │   ├── docs/                         # Docs shell, markdown renderer, site-map
│   │   ├── demo/                         # Live demo components
│   │   ├── marketing/                    # Landing page
│   │   └── shell/                        # Site header / footer / layout
│   └── components/
│       └── onlyoffice-embed-sdk/          # OnlyOffice Embed SDK wrapper + docs/*.md source
├── public/                               # OnlyOffice SDK static assets
└── ...
```

Docs pages read Markdown directly from `src/components/onlyoffice-embed-sdk/docs/`. Demo tabs embed live editors from the OnlyOffice Embed SDK in `src/features/demo/`.

## Tech Stack

- **OnlyOffice SDK**: Core editing
- **x2t + WebAssembly**: Format conversion
- **TanStack Start + React 19**: SSR and file-based routing
- **Cloudflare Workers**: production runtime, custom domains, and response headers

## Deployment

```bash
pnpm install
pnpm build
pnpm run deploy
```

The application deploys as a TanStack Start SSR Worker through Wrangler. Production routes are configured in `wrangler.jsonc`, including `onlyoffice.agent-bridges.com` and twelve fixed Chinese-zodiac subframe hosts: `rat`, `ox`, `tiger`, `rabbit`, `dragon`, `snake`, `horse`, `goat`, `monkey`, `rooster`, `dog`, and `pig`. Vercel is not used.

### Deploy SDK Assets to Cloudflare Pages CDN

OnlyOffice SDK assets can be hosted separately from the app. Deploy the contents of `public/packages` to Cloudflare Pages, then register the Pages URL before creating an editor.

```bash
# one-time project creation
npx wrangler pages project create onlyoffice-embed-resource

# upload public/packages as the CDN root
npx wrangler pages deploy public/packages \
  --project-name onlyoffice-embed-resource \
  --commit-dirty=true
```

After deployment, use the fixed production Pages domain:

```text
https://onlyoffice-embed-resource.pages.dev/onlyoffice/9.4.0-develop/web-apps/apps/api/documents/api.js
```

Use that Pages origin as the runtime resource root:

```typescript
import { OnlyOfficeManager } from "@/components/onlyoffice-embed-sdk";

OnlyOfficeManager.registerStaticResource({
  cdnOrigin: "https://onlyoffice-embed-resource.pages.dev",
});
```

`cdnOrigin` points to the uploaded `public/packages` root, so do not append `/packages`. Configure this resource origin in [`src/components/onlyoffice-embed-sdk/const/index.ts`](src/components/onlyoffice-embed-sdk/const/index.ts) by editing the `cdnOrigin` logic in `buildStaticResource`. Cloudflare Pages Direct Upload supports Wrangler folder uploads; dashboard drag-and-drop is less suitable for this repository because the SDK contains many files.

### GitHub Actions production deployment

Pushes to `main` run [`.github/workflows/deploy-cloudflare.yml`](.github/workflows/deploy-cloudflare.yml). After type checking and building, it deploys `public/packages` to the `onlyoffice-embed-resource` Pages project and publishes the TanStack Start Worker serving `onlyoffice.agent-bridges.com` and its twelve Subframe hosts.

Configure these repository secrets before merging:

- `CLOUDFLARE_API_TOKEN`: a scoped Cloudflare API token with Pages and Workers deployment permissions
- `CLOUDFLARE_ACCOUNT_ID`: `40a503f1c86c028edfcd6f113c562b5b`

The application uses the stable Pages origin `https://onlyoffice-embed-resource.pages.dev`; deployment-specific Pages URLs are not used by the app.

## Fonts

Custom fonts are registered via **`__custom_font_registry__`**, with **`ttf-to-catalog-font.mjs`** producing OnlyOffice catalog wire-format files. See **[Fonts](src/components/onlyoffice-embed-sdk/docs/字体配置.md)** in the component docs for the full guide.

Quick outline:

1. Run `ttf-to-catalog-font.mjs --id <id> --verify` to produce `fonts/{id}` catalog files
2. Register the id and aliases in `window["__custom_font_registry__"]` inside `AllFonts.js`
3. Ensure aliases cover every font name used in your documents

Ensure all font files comply with applicable licenses.

## Related Resources

- [OnlyOffice API docs](https://api.onlyoffice.com/docs/docs-api/usage-api/config/document/)
- [OnlyOffice Web Apps](https://github.com/ONLYOFFICE/web-apps)
- [OnlyOffice SDK](https://github.com/ONLYOFFICE/sdkjs)
- [agentbridges-ai/onlyoffice-x2t-wasm](https://github.com/agentbridges-ai/onlyoffice-x2t-wasm)

## Contributing

Issues and Pull Requests are welcome.

## License

See [LICENSE](LICENSE).
