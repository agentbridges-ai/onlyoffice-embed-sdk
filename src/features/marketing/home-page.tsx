"use client";

/**
 * 产品主页：Hero、特性网格、快速集成步骤与代码片段。
 */
import { useState } from "react";
import { OfficeHeroVisual } from "./office-hero-visual";
import { SiteLinkButton, SiteSectionTitle, SiteCard, SITE_GITHUB } from "@/features/shell";

const FEATURES = [
  {
    title: "易于集成",
    description: "通过 OnlyOfficeManager 与 EditorManager，几行代码即可挂载编辑器。",
  },
  {
    title: "工程化支持",
    description:
      "提供只读/编辑切换、多实例容器隔离、主题与语言等运行时 API，便于接入业务流。",
  },
  {
    title: "安全可靠",
    description: "文档在浏览器本地处理，无需上传至 Document Server。",
  },
  {
    title: "多格式支持",
    description:
      "支持 DOCX、DOCM、XLSX、CSV、PPTX、ODT、PDF 等常见 Office 格式。",
  },
] as const;

const STEPS = [
  { id: 1, title: "安装", body: "将 OnlyOffice Embed SDK 引入你的前端项目。" },
  { id: 2, title: "初始化", body: "调用 initializeOnlyOffice 加载静态 SDK。" },
  { id: 3, title: "挂载", body: "创建容器并通过 OnlyOfficeManager 打开文档。" },
] as const;

const INSTALL_SNIPPET = `import {
  OnlyOfficeManager,
  FILE_TYPE,
  ONLYOFFICE_ID,
} from "@/components/onlyoffice-web-comp";

const manager = await OnlyOfficeManager.create({
  containerId: ONLYOFFICE_ID,
  fileType: FILE_TYPE.DOCX,
  defaultFileName: "New_Document.docx",
});`;

function CopyButton({ text }: { text: string }) {
  const [copied, setCopied] = useState(false);

  return (
    <button
      type="button"
      onClick={async () => {
        await navigator.clipboard.writeText(text);
        setCopied(true);
        window.setTimeout(() => setCopied(false), 1500);
      }}
      className="border border-neutral-700 px-2 py-1 text-[11px] text-neutral-400 hover:border-neutral-500 hover:text-neutral-200"
    >
      {copied ? "已复制" : "复制"}
    </button>
  );
}

export function HomePage() {
  return (
    <>
      <section className="site-hero relative overflow-hidden border-b border-neutral-200">
        <div className="site-hero-grid pointer-events-none absolute inset-0" />
        <div className="relative mx-auto grid max-w-7xl gap-12 px-4 py-16 sm:px-6 lg:grid-cols-[1.05fr_0.95fr] lg:items-center lg:py-24">
          <div>
            <p className="mb-5 inline-flex border border-neutral-300 px-2.5 py-1 text-[11px] font-medium uppercase tracking-[0.14em] text-neutral-600">
              Browser-native · No Document Server
            </p>
            <h1 className="max-w-xl text-[clamp(2.25rem,5.5vw,3.75rem)] font-semibold leading-[1.05] tracking-[-0.04em] text-neutral-950">
              将 OnlyOffice
              <br />
              集成到你的 Web 应用
            </h1>
            <p className="mt-6 max-w-lg text-[15px] leading-relaxed text-neutral-600">
              OnlyOffice Embed SDK 是一套面向 Web 的浏览器端编辑能力，不依赖后端即可帮助你嵌入
              Word、Excel、PPT 编辑器，实现查看、编辑与导出。
            </p>
            <div className="mt-9 flex flex-wrap gap-2">
              <SiteLinkButton href="/docs/demos/single">开始使用</SiteLinkButton>
              <SiteLinkButton href="/docs" variant="secondary">
                阅读文档
              </SiteLinkButton>
              <SiteLinkButton href={SITE_GITHUB} external variant="secondary">
                GitHub
              </SiteLinkButton>
            </div>
          </div>

          <OfficeHeroVisual />
        </div>
      </section>

      <section className="border-b border-neutral-200 py-16 sm:py-20">
        <div className="mx-auto max-w-7xl px-4 sm:px-6">
          <div className="mb-10 max-w-xl">
            <SiteSectionTitle>为什么选择 OnlyOffice Embed SDK</SiteSectionTitle>
            <p className="mt-2 text-[14px] text-neutral-600">
              面向前端工程化场景设计，API 清晰、实例隔离、可嵌入文档与示例页。
            </p>
          </div>
          <div className="grid gap-px border border-neutral-200 bg-neutral-200 sm:grid-cols-2 lg:grid-cols-4">
            {FEATURES.map((feature) => (
              <SiteCard key={feature.title} className="border-0">
                <h3 className="text-[15px] font-medium text-neutral-950">
                  {feature.title}
                </h3>
                <p className="mt-2 text-[13px] leading-relaxed text-neutral-600">
                  {feature.description}
                </p>
              </SiteCard>
            ))}
          </div>
        </div>
      </section>

      <section className="bg-neutral-50 py-16 sm:py-20">
        <div className="mx-auto max-w-7xl px-4 sm:px-6">
          <div className="mx-auto max-w-2xl text-center">
            <SiteSectionTitle>快速集成，只需几步</SiteSectionTitle>
            <p className="mt-2 text-[14px] text-neutral-600">
              几行代码即可嵌入功能完整的文档编辑器
            </p>
          </div>

          <div className="mt-12 grid gap-10 lg:grid-cols-[240px_1fr]">
            <ol className="space-y-5">
              {STEPS.map((step) => (
                <li key={step.id} className="flex gap-3">
                  <span className="flex h-8 w-8 shrink-0 items-center justify-center border border-neutral-900 text-[12px] font-medium">
                    {step.id}
                  </span>
                  <div>
                    <p className="text-[14px] font-medium text-neutral-950">
                      {step.title}
                    </p>
                    <p className="mt-1 text-[13px] text-neutral-600">
                      {step.body}
                    </p>
                  </div>
                </li>
              ))}
            </ol>

            <div className="border border-neutral-900 bg-neutral-950 shadow-[10px_10px_0_0_#e5e5e5]">
              <div className="flex items-center justify-between border-b border-neutral-800 px-4 py-2">
                <span className="text-[12px] text-neutral-500">example.ts</span>
                <CopyButton text={INSTALL_SNIPPET} />
              </div>
              <pre className="overflow-x-auto p-4 text-[12px] leading-relaxed text-neutral-300">
                <code>{INSTALL_SNIPPET}</code>
              </pre>
            </div>
          </div>
        </div>
      </section>
    </>
  );
}
