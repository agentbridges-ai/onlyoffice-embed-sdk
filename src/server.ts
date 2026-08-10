import handler, { createServerEntry } from "@tanstack/react-start/server-entry";

const STATIC_RESOURCE_ORIGIN =
  "https://d69d942c.onlyoffice-embed-resource.pages.dev";

async function fetchStaticResource(request: Request): Promise<Response | null> {
  const url = new URL(request.url);
  const isPackagePath = url.pathname.startsWith("/packages/");
  const isLegacyResourcePath = url.pathname.startsWith("/onlyoffice/");
  if (!isPackagePath && !isLegacyResourcePath) return null;

  const resourcePrefix = isPackagePath ? "/packages" : "";
  const resourcePath = isPackagePath
    ? url.pathname.slice("/packages".length)
    : url.pathname;
  const targetPath = resourcePath.endsWith("/index.html")
    ? resourcePath.slice(0, -"index.html".length)
    : resourcePath;
  const target = new URL(
    `${targetPath}${url.search}`,
    `${STATIC_RESOURCE_ORIGIN}/`,
  );
  const response = await fetch(new Request(target, request));
  const location = response.headers.get("Location");

  if (!location) return response;

  const locationUrl = new URL(location, target);
  if (
    locationUrl.origin !== target.origin ||
    locationUrl.pathname.startsWith("/packages/")
  ) {
    return response;
  }

  const headers = new Headers(response.headers);
  headers.set("Cache-Control", "no-store");
  headers.set(
    "Location",
    `${resourcePrefix}${locationUrl.pathname}${locationUrl.search}${locationUrl.hash}`,
  );

  return new Response(response.body, {
    status: response.status,
    statusText: response.statusText,
    headers,
  });
}

export default createServerEntry({
  async fetch(request, options) {
    const response =
      (await fetchStaticResource(request)) ?? (await handler.fetch(request, options));
    const headers = new Headers(response.headers);
    headers.set("Origin-Agent-Cluster", "?1");

    return new Response(response.body, {
      status: response.status,
      statusText: response.statusText,
      headers,
    });
  },
});
