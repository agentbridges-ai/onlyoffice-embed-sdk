/**
 * 读取 `onlyoffice-web-comp/docs/` 下的 Markdown 源文件。
 *
 * Vite bundles these documents as raw assets so the same code works in the
 * browser, during SSR, and inside a Cloudflare Worker.
 */
import { getMarkdownDocBySlug } from "../config/site-map";
import { getCompDocContent } from "./content";

export function readCompDoc(file: string): string {
  return getCompDocContent(file);
}

export function readMarkdownDocBySlug(slug: string): string | null {
  const doc = getMarkdownDocBySlug(slug);
  if (!doc) return null;
  return readCompDoc(doc.file);
}
