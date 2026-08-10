"use client";

/**
 * 将组件库 Markdown 渲染为文档站 HTML，并改写 `./xx.md` 为站内路由。
 * 本页目录：右侧 sticky，由 h1/h2 自动生成；锚点：rehype-slug。
 */
import { Link } from "@tanstack/react-router";
import { useMemo } from "react";
import type { Components } from "react-markdown";
import ReactMarkdown from "react-markdown";
import rehypeSlug from "rehype-slug";
import remarkGfm from "remark-gfm";
import { resolveMarkdownHref } from "../config/site-map";
import { getTocHeadings } from "../lib/extract-headings";
import { DocToc } from "./doc-toc";

const components: Components = {
  h1: ({ id, children }) => (
    <h1 id={id} className="docs-md-h1 scroll-mt-24">
      {children}
    </h1>
  ),
  h2: ({ id, children }) => (
    <h2 id={id} className="docs-md-h2 scroll-mt-24">
      {children}
    </h2>
  ),
  h3: ({ id, children }) => (
    <h3 id={id} className="docs-md-h3 scroll-mt-24">
      {children}
    </h3>
  ),
  p: ({ children }) => <p className="docs-md-p">{children}</p>,
  ul: ({ children }) => <ul className="docs-md-ul">{children}</ul>,
  ol: ({ children }) => <ol className="docs-md-ol">{children}</ol>,
  li: ({ children }) => <li className="docs-md-li">{children}</li>,
  blockquote: ({ children }) => (
    <blockquote className="docs-md-blockquote">{children}</blockquote>
  ),
  hr: () => <hr className="docs-md-hr" />,
  table: ({ children }) => (
    <div className="docs-table-wrap">
      <table className="docs-table">{children}</table>
    </div>
  ),
  thead: ({ children }) => <thead>{children}</thead>,
  tbody: ({ children }) => <tbody>{children}</tbody>,
  tr: ({ children }) => <tr>{children}</tr>,
  th: ({ children }) => <th>{children}</th>,
  td: ({ children }) => <td>{children}</td>,
  pre: ({ children }) => <pre className="docs-code">{children}</pre>,
  code: ({ className, children, ...props }) => {
    const isBlock = Boolean(className);
    if (isBlock) {
      return (
        <code className={className} {...props}>
          {children}
        </code>
      );
    }
    return (
      <code className="docs-md-inline-code" {...props}>
        {children}
      </code>
    );
  },
  a: ({ href, children }) => {
    const resolved = resolveMarkdownHref(href);
    const isExternal =
      resolved?.startsWith("http") || resolved?.startsWith("mailto:");

    if (isExternal) {
      return (
        <a
          href={resolved}
          className="docs-md-link"
          target="_blank"
          rel="noopener noreferrer"
        >
          {children}
        </a>
      );
    }

    if (resolved?.startsWith("/")) {
      return (
        <Link to={resolved} className="docs-md-link">
          {children}
        </Link>
      );
    }

    return (
      <a href={resolved} className="docs-md-link">
        {children}
      </a>
    );
  },
};

type MarkdownContentProps = {
  content: string;
  showToc?: boolean;
};

export function MarkdownContent({ content, showToc = true }: MarkdownContentProps) {
  const tocHeadings = useMemo(() => getTocHeadings(content), [content]);

  const article = (
    <article className="docs-markdown w-full min-w-0 max-w-full">
      <ReactMarkdown
        remarkPlugins={[remarkGfm]}
        rehypePlugins={[rehypeSlug]}
        components={components}
      >
        {content}
      </ReactMarkdown>
    </article>
  );

  if (!showToc) {
    return article;
  }

  return (
    <div className="docs-article-layout">
      {article}
      <DocToc headings={tocHeadings} />
    </div>
  );
}
