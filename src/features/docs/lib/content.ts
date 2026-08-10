import type { MarkdownDoc } from "../config/site-map";

const markdownModules = import.meta.glob(
  "../../../components/onlyoffice-web-comp/docs/*.md",
  {
    eager: true,
    import: "default",
    query: "?raw",
  },
) as Record<string, string>;

function findMarkdownModule(file: string): string | undefined {
  const suffix = `/docs/${file}`;
  const modulePath = Object.keys(markdownModules).find((key) =>
    key.endsWith(suffix),
  );

  return modulePath ? markdownModules[modulePath] : undefined;
}

export function getCompDocContent(file: string): string {
  const content = findMarkdownModule(file);
  if (content === undefined) {
    throw new Error(`Markdown document not found: ${file}`);
  }
  return content;
}

export function getMarkdownDocContent(doc: MarkdownDoc): string {
  return getCompDocContent(doc.file);
}
