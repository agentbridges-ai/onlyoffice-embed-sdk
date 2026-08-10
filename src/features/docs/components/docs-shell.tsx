"use client";

/**
 * 文档站布局：左侧目录 + 右侧 Markdown 内容区。
 * 目录数据来自 {@link ../config/site-map.ts}；示例项使用独立路由区分激活态。
 */
import { Link, useLocation } from "@tanstack/react-router";
import { getDocsNav, isDocsNavActive } from "../config/site-map";

export function DocsShell({ children }: { children: React.ReactNode }) {
  const pathname = useLocation().pathname;
  const navGroups = getDocsNav();

  return (
    <div className="docs-layout w-full overflow-x-clip">
      <div className="docs-layout-inner mx-auto w-full max-w-[90rem] px-4 sm:px-6 lg:px-8">
        <div className="grid w-full grid-cols-1 gap-8 py-6 sm:gap-10 sm:py-8 lg:grid-cols-[minmax(0,220px)_minmax(0,1fr)] lg:gap-14 lg:py-12">
          <aside className="w-full min-w-0 lg:sticky lg:top-20 lg:self-start">
            <p className="text-[11px] font-medium uppercase tracking-[0.14em] text-neutral-400">
              Documentation
            </p>
            <nav className="mt-4 space-y-6" aria-label="文档目录">
              {navGroups.map((group) => (
                <div key={group.title}>
                  <p className="mb-2 text-[12px] font-medium text-neutral-500">
                    {group.title}
                  </p>
                  <ul className="space-y-0.5 border-l border-neutral-200">
                    {group.items.map((item) => {
                      const active = isDocsNavActive(pathname, item.href);
                      return (
                        <li key={item.href}>
                          <Link
                            to={item.href}
                            className={`-ml-px block border-l py-1.5 pl-3 text-[13px] transition-colors ${
                              active
                                ? "border-neutral-900 font-medium text-neutral-950"
                                : "border-transparent text-neutral-600 hover:border-neutral-400 hover:text-neutral-900"
                            }`}
                          >
                            {item.label}
                          </Link>
                        </li>
                      );
                    })}
                  </ul>
                </div>
              ))}
            </nav>
          </aside>

          <div className="w-full min-w-0 max-w-full docs-prose">{children}</div>
        </div>
      </div>
    </div>
  );
}
