import { createFileRoute } from "@tanstack/react-router";
import { DocsDemoPage } from "@/features/docs/components/docs-demo-page";
import { readCompDoc } from "@/features/docs/lib/server";
import { TabsMultiPage } from "@/features/demo/tabs-multi-page";

export const Route = createFileRoute("/docs/demos/multi")({
  head: () => ({
    meta: [{ title: "多实例示例 — OnlyOffice Web Comp" }],
  }),
  component: DocsDemosMultiPage,
});

function DocsDemosMultiPage() {
  const content = readCompDoc("多实例示例.md");

  return (
    <DocsDemoPage content={content} marker="multi">
      <TabsMultiPage embedded />
    </DocsDemoPage>
  );
}
