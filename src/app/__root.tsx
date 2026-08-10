import {
  HeadContent,
  Outlet,
  Scripts,
  createRootRoute,
} from "@tanstack/react-router";
import appCss from "./globals.css?url";
import ClientLayoutWrapper from "./-ClientLayoutWrapper";

export const Route = createRootRoute({
  head: () => ({
    meta: [
      { charSet: "utf-8" },
      { name: "viewport", content: "width=device-width, initial-scale=1" },
      { title: "OnlyOffice Embed SDK" },
      {
        name: "description",
        content: "浏览器端 OnlyOffice 编辑器组件，无需 Document Server",
      },
    ],
    links: [{ rel: "stylesheet", href: appCss }],
  }),
  component: RootDocument,
  notFoundComponent: NotFoundPage,
});

function RootDocument() {
  return (
    <html lang="zh-CN">
      <head>
        <HeadContent />
      </head>
      <body className="antialiased">
        <ClientLayoutWrapper>
          <Outlet />
        </ClientLayoutWrapper>
        <Scripts />
      </body>
    </html>
  );
}

function NotFoundPage() {
  return (
    <main className="mx-auto flex min-h-[50vh] max-w-3xl items-center justify-center px-6 py-20">
      <div>
        <p className="text-sm text-neutral-500">404</p>
        <h1 className="mt-2 text-2xl font-semibold text-neutral-950">
          页面不存在
        </h1>
      </div>
    </main>
  );
}
