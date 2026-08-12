# OnlyOffice Embed SDK

> 📖 [English](README.md) | 中文

🌐 **在线演示**: [https://onlyoffice.agent-bridges.com/](https://onlyoffice.agent-bridges.com/)

基于 OnlyOffice 静态 SDK 的浏览器端文档处理方案：在客户端完成 Word / Excel / PowerPoint 的查看、编辑与转换，**无需 Document Server**。

本仓库包含两部分：

| 部分 | 路径 | 说明 |
|------|------|------|
| **可发布 SDK** | [`packages/onlyoffice-embed-sdk/`](packages/onlyoffice-embed-sdk/) | `@agentbridges-ai/onlyoffice-embed-sdk` 包、构建与迁移说明 |
| **共享实现** | [`src/components/onlyoffice-embed-sdk/`](src/components/onlyoffice-embed-sdk/) | 编辑器、x2t、bridge、兼容层与 Markdown 文档源 |
| **演示站点** | [`src/app/`](src/app/) + [`src/features/`](src/features/) | TanStack Start 文件路由 + Cloudflare Worker |

## 项目定位

仓库同时包含私有演示/部署应用，以及独立可发布的 ESM 包 `@agentbridges-ai/onlyoffice-embed-sdk`。npm 包内包含编辑器生命周期、加固后的跨域 bridge 与自包含 x2t Worker；体积较大的 ONLYOFFICE SDK、x2t WASM 和字体仍作为独立运行时资源部署。

如果你希望在自己的 Web 项目里接入 OnlyOffice，并且不想部署 OnlyOffice Document Server，可以把这里当作一套可复制的工程实现。演示站点也是项目的一部分，目的是让你直接参考一个已经跑通的编辑器生命周期，而不是只看零散 API 片段。

## 怎么集成进你的项目

实际接入可以按这个路径做：

1. 安装 `@agentbridges-ai/onlyoffice-embed-sdk` 并使用原生 manager API；从独立 host API 迁移的应用使用 `/compat/subframe`，由正式入口挂载到 12 个固定隔离 origin。直接 `/compat` 仅用于受控的同窗口集成。
2. 将 [`public/packages/onlyoffice/`](public/packages/onlyoffice/) 部署在应用同源路径或 CDN，并在首次创建编辑器前调用 `registerOnlyOfficeStaticResource`。
3. 参考 [`src/features/demo/office-preview-page.tsx`](src/features/demo/office-preview-page.tsx) 构造自己的界面：准备编辑器容器，维护一个 `OnlyOfficeManager` 实例，按需调用 `openDocument`、`downloadExport`、`toggleReadOnly`，并在页面卸载时销毁 manager。需要从父页面调用编辑器 Automation API 时，可通过 `createConnector()` 获取 Developer Edition Connector。

托管 compat subframe 自行提供运行时资源，父应用无需复制 `public/packages`。安装、CSP、兼容差异和发布约定详见 [`packages/onlyoffice-embed-sdk/README.md`](packages/onlyoffice-embed-sdk/README.md)。

静态资源读取统一在 [`src/components/onlyoffice-embed-sdk/const/index.ts`](src/components/onlyoffice-embed-sdk/const/index.ts) 配置。编辑器 UI 读取 Developer Edition Docker 导出的 9.4 SDK：`/packages/onlyoffice/9.4.0-develop`；转换器独立锁定到签名发布的 `agentbridges-ai/onlyoffice-x2t-wasm`，路径为 `/packages/onlyoffice/x2t/<tag>`。`onlyofficeVersion` 只切换编辑器 UI 资源目录。

## 核心优势

- **数据留在本地**：文档处理在浏览器内完成
- **格式兼容**：Word、Excel、PowerPoint、CSV、DOCM 等
- **零后端**：托管静态 SDK 即可使用
- **工程化 API**：只读/编辑、主题/语言、多实例容器隔离
- **连接器支持**：通过 Developer Edition Connector 从父页面调用编辑器 Automation API，本地和 CDN 模式均可用

## 快速体验

1. 访问 [在线演示](https://onlyoffice.agent-bridges.com/) 或本地启动：

```bash
git clone <repository-url>
cd onlyoffice-embed-sdk
pnpm install
pnpm dev
# http://localhost:3001
```

2. 打开路由：

| 路由 | 说明 |
|------|------|
| `/` | 产品主页 |
| `/docs` | OnlyOffice Embed SDK 文档（直接渲染 Markdown） |
| `/docs/demos/single` | 单实例在线示例 |
| `/docs/demos/multi` | 多实例 Tab 在线示例 |
| `/api/version` | 当前可安装 SDK 版本（JSON） |

3. 上传本地文件 → 编辑 → 导出

旧路由 `/examples` 会重定向到单实例示例；`/multi` 会重定向到多实例示例。

## 版本接口

首页版本徽标直接读取公开接口，不依赖页面内硬编码：

```bash
curl https://onlyoffice.agent-bridges.com/api/version
```

```json
{
  "name": "@agentbridges-ai/onlyoffice-embed-sdk",
  "version": "0.3.1",
  "release": "sdk-v0.3.1",
  "hostIdentity": {
    "packageVersion": "0.3.1",
    "hostBuildId": "onlyoffice-embed-sdk-hosted-v3",
    "assetManifestDigest": "1ba560ce80aa78dbb6a9b55a7fec42b98dc8a6cdcf18f43c44d5a8d57aa2a096"
  },
  "runtimeManifest": {
    "packageVersion": "0.3.1",
    "hostBuildId": "onlyoffice-embed-sdk-hosted-v3",
    "compatSubframeProtocol": 1,
    "compatSubframePath": "/subframe?runtime=compat",
    "canonicalResourceOrigin": "https://onlyoffice.agent-bridges.com",
    "onlyofficeResourcePath": "/onlyoffice/runtime/onlyoffice-embed-sdk-hosted-v3",
    "onlyofficeVersion": "9.4.0-develop",
    "x2t": {
      "repository": "agentbridges-ai/onlyoffice-x2t-wasm",
      "tag": "v9.3.0+4",
      "sourceCommit": "5790fda684ac1a837f1ab92fcaaf4a0de9ec4ec1"
    }
  }
}
```

接口支持 `GET`、`HEAD` 和跨域查询，并使用 `Cache-Control: no-store`
确保返回当前部署版本。
上例中的 `x2t` 展示身份坐标；线上完整响应还包含锁定的工具链、压缩/解压
资产哈希以及真实 DOC/Pivot 回归摘要。
`hostIdentity.assetManifestDigest` 是同一接口返回的规范
`runtimeManifest` JSON 字节的 SHA-256，可用于锁定已部署的协议坐标；它不等同于
对 CDN 中每个资产逐字节审计。

## OnlyOffice Embed SDK 文档

**API 与接入说明不在本 README 重复**，请阅读组件库文档：

- **入口**：[OnlyOffice Embed SDK README（中文）](src/components/onlyoffice-embed-sdk/readme.zh.md)
- **概述**：[docs/概述.md](src/components/onlyoffice-embed-sdk/docs/概述.md)

| 文档 | 内容 |
|------|------|
| [快速开始](src/components/onlyoffice-embed-sdk/docs/快速开始.md) | 初始化与容器挂载 |
| [核心API](src/components/onlyoffice-embed-sdk/docs/核心API.md) | `OnlyOfficeManager`、多实例 |
| [事件系统](src/components/onlyoffice-embed-sdk/docs/事件系统.md) | EventBus |
| [完整示例](src/components/onlyoffice-embed-sdk/docs/完整示例.md) | React 集成模式 |
| [API参考](src/components/onlyoffice-embed-sdk/docs/API参考.md) | 常量与类型 |
| [注意事项与格式](src/components/onlyoffice-embed-sdk/docs/注意事项与支持格式.md) | 前置条件与格式 |
| [字体配置](src/components/onlyoffice-embed-sdk/docs/字体配置.md) | 自定义字体注册 |
| [批注修订](src/components/onlyoffice-embed-sdk/docs/批注修订与-Word-API.md) | 批注、修订 |
| [单实例示例](src/components/onlyoffice-embed-sdk/docs/单实例示例.md) | 单实例 Demo 与源码说明 |
| [多实例示例](src/components/onlyoffice-embed-sdk/docs/多实例示例.md) | Tab 多实例完整源码 |

```typescript
import {
  OnlyOfficeManager,
  FILE_TYPE,
  ONLYOFFICE_ID,
} from "@agentbridges-ai/onlyoffice-embed-sdk";
```

## 项目结构

```
onlyoffice-embed-sdk/
├── packages/
│   └── onlyoffice-embed-sdk/             # 可发布 ESM 包
├── src/
│   ├── app/                              # TanStack Start 文件路由
│   │   ├── __root.tsx                    # HTML 壳与全局元数据
│   │   ├── index.tsx                     # 主页
│   │   ├── docs/                         # 文档路由
│   │   │   ├── index.tsx                  # /docs 概述
│   │   │   ├── $slug.tsx                  # /docs/* Markdown 页面
│   │   │   └── demos/                    # /docs/demos/single|multi
│   │   └── examples.tsx                  # → 重定向至单实例示例
│   ├── features/
│   │   ├── docs/                         # 文档壳、Markdown 渲染、site-map
│   │   ├── demo/                         # 在线演示组件
│   │   ├── marketing/                    # 着陆页
│   │   └── shell/                        # 站点 Header / Footer / Layout
│   └── components/
│       └── onlyoffice-embed-sdk/          # OnlyOffice Embed SDK 封装 + docs/*.md 文档源
├── public/                               # OnlyOffice SDK 静态资源
└── ...
```

文档页直接读取 `src/components/onlyoffice-embed-sdk/docs/` 下的 Markdown；示例 Tab 内嵌 `src/features/demo/` 的可交互 OnlyOffice Embed SDK 编辑器。

## 技术栈

- **OnlyOffice SDK**：文档编辑核心
- **x2t + WebAssembly**：格式转换
- **TanStack Start + React 19**：SSR 与文件路由
- **Cloudflare Workers**：生产运行时、自定义域名和响应头

## 部署

```bash
pnpm install
pnpm build
pnpm run deploy
```

应用通过 Wrangler 以 TanStack Start SSR Worker 部署。生产路由写在 `wrangler.jsonc`，包括 `onlyoffice.agent-bridges.com` 和隔离子帧使用的 12 个固定生肖子域名：`rat`、`ox`、`tiger`、`rabbit`、`dragon`、`snake`、`horse`、`goat`、`monkey`、`rooster`、`dog`、`pig`。项目不再使用 Vercel。

### 将 SDK 静态资源部署到 Cloudflare Pages CDN

OnlyOffice SDK 资源可以和应用分开托管。把 `public/packages` 的内容部署到 Cloudflare Pages，然后在创建编辑器前注册 Pages 地址即可。

```bash
# 首次创建项目
npx wrangler pages project create onlyoffice-embed-resource

# 将 public/packages 作为 CDN 根目录上传
npx wrangler pages deploy public/packages \
  --project-name onlyoffice-embed-resource \
  --commit-dirty=true
```

部署后使用固定的生产 Pages 域名：

```text
https://onlyoffice-embed-resource.pages.dev/onlyoffice/9.4.0-develop/web-apps/apps/api/documents/api.js
```

在运行时把 Pages origin 注册为静态资源根地址：

```typescript
import { OnlyOfficeManager } from "@/components/onlyoffice-embed-sdk";

OnlyOfficeManager.registerStaticResource({
  cdnOrigin: "https://onlyoffice-embed-resource.pages.dev",
});
```

`cdnOrigin` 对应上传后的 `public/packages` 根目录，不需要再追加 `/packages`。在 [`src/components/onlyoffice-embed-sdk/const/index.ts`](src/components/onlyoffice-embed-sdk/const/index.ts) 中修改 `buildStaticResource` 的 `cdnOrigin` 逻辑即可固定资源来源。Cloudflare Pages Direct Upload 支持用 Wrangler 上传目录；由于 SDK 文件数量较多，Dashboard 拖拽上传不太适合本仓库。

### GitHub Actions 生产发布

推送到 `main` 会运行 [`.github/workflows/deploy-cloudflare.yml`](.github/workflows/deploy-cloudflare.yml)。通过类型检查和构建后，Workflow 会把 `public/packages` 发布到 `onlyoffice-embed-resource` Pages 项目，并发布提供 `onlyoffice.agent-bridges.com` 及 12 个 Subframe 子域名的 TanStack Start Worker。

合并前请在仓库中配置以下 Secrets：

- `CLOUDFLARE_API_TOKEN`：具备 Pages 和 Workers 发布权限的受限 Cloudflare API Token
- `CLOUDFLARE_ACCOUNT_ID`：`40a503f1c86c028edfcd6f113c562b5b`

应用使用固定的 Pages 资源地址 `https://onlyoffice-embed-resource.pages.dev`，不会引用部署专属的 ID 地址。

## 字体配置

自定义字体通过 **`__custom_font_registry__`** 注册，配合 **`ttf-to-catalog-font.mjs`** 生成 OnlyOffice catalog 线格式。完整步骤见组件库文档 **[字体配置](src/components/onlyoffice-embed-sdk/docs/字体配置.md)**。

简要流程：

1. 运行 `ttf-to-catalog-font.mjs --id <id> --verify` 生成 `fonts/{id}` catalog 文件
2. 在 `AllFonts.js` 的 `window["__custom_font_registry__"]` 中注册 id 与别名
3. 确保别名覆盖文档内实际使用的字体名

请确保所用字体文件符合相关许可协议。

## 相关资源

- [OnlyOffice API 文档](https://api.onlyoffice.com/zh-CN/docs/docs-api/usage-api/config/document/)
- [OnlyOffice Web Apps](https://github.com/ONLYOFFICE/web-apps)
- [OnlyOffice SDK](https://github.com/ONLYOFFICE/sdkjs)
- [agentbridges-ai/onlyoffice-x2t-wasm](https://github.com/agentbridges-ai/onlyoffice-x2t-wasm)

## 参与贡献

欢迎提交 Issue 和 Pull Request。

## 开源许可

详见 [LICENSE](LICENSE)。
