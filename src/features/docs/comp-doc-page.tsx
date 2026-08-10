/**
 * 服务端 Markdown 文档页：读取 OnlyOffice Embed SDK 的 Markdown 文档并渲染。
 */
import {
  extractDocTitle,
  getMarkdownDocBySlug,
} from "./config/site-map";
import { readCompDoc } from "./lib/server";
import { MarkdownContent } from "./components/markdown-content";

type CompDocPageProps = {
  slug: string;
};

export function CompDocPage({ slug }: CompDocPageProps) {
  const doc = getMarkdownDocBySlug(slug);
  if (!doc) {
    return null;
  }

  const content = readCompDoc(doc.file);

  return <MarkdownContent content={content} />;
}

export function getCompDocMetadata(slug: string) {
  const doc = getMarkdownDocBySlug(slug);
  if (!doc) return { title: "文档" };

  const content = readCompDoc(doc.file);
  const title = extractDocTitle(content, doc.label);

  return {
    title: `${title} — OnlyOffice Embed SDK`,
  };
}
