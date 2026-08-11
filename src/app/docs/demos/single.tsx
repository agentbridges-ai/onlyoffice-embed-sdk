import { createFileRoute } from "@tanstack/react-router";
import { FILE_TYPE } from "@/components/onlyoffice-embed-sdk";
import { DocsDemoPage } from "@/features/docs/components/docs-demo-page";
import { readCompDoc } from "@/features/docs/lib/server";
import { OfficePreviewPage } from "@/features/demo/office-preview-page";

export const Route = createFileRoute("/docs/demos/single")({
  head: () => ({
    meta: [{ title: "单实例示例 — OnlyOffice Embed SDK" }],
  }),
  component: DocsDemosSinglePage,
});

function DocsDemosSinglePage() {
  const content = readCompDoc("单实例示例.md");

  return (
    <DocsDemoPage content={content} marker="single">
      <OfficePreviewPage
        embedded
        title="单实例"
        defaultFileName="New_Document.docx"
        fileType={FILE_TYPE.DOCX}
        newButtonLabel="新建文档"
      />
    </DocsDemoPage>
  );
}
