import { createFileRoute } from "@tanstack/react-router";
import {
  CompDocPage,
  getCompDocMetadata,
} from "@/features/docs/comp-doc-page";

export const Route = createFileRoute("/docs/")({
  head: () => ({
    meta: [{ title: getCompDocMetadata("").title }],
  }),
  component: DocsOverviewPage,
});

function DocsOverviewPage() {
  return <CompDocPage slug="" />;
}
