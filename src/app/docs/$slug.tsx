import { createFileRoute, notFound } from "@tanstack/react-router";
import {
  CompDocPage,
  getCompDocMetadata,
} from "@/features/docs/comp-doc-page";
import { getMarkdownDocBySlug } from "@/features/docs/config/site-map";

export const Route = createFileRoute("/docs/$slug")({
  loader: ({ params }) => {
    if (!getMarkdownDocBySlug(params.slug)) {
      throw notFound();
    }

    return {
      title: getCompDocMetadata(params.slug).title,
      slug: params.slug,
    };
  },
  head: ({ loaderData }) => ({
    meta: [{ title: loaderData?.title ?? "文档 — OnlyOffice Embed SDK" }],
  }),
  component: DocsSlugPage,
});

function DocsSlugPage() {
  const { slug } = Route.useLoaderData();
  return <CompDocPage slug={slug} />;
}
