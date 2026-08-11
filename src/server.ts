import handler, { createServerEntry } from "@tanstack/react-start/server-entry";
import { ONLYOFFICE_EMBED_SDK_VERSION } from "./components/onlyoffice-embed-sdk/compat/version";

const STATIC_RESOURCE_ORIGIN =
  "https://onlyoffice-embed-resource.pages.dev";
const SDK_PACKAGE_NAME = "@agentbridges-ai/onlyoffice-embed-sdk";
const VERSION_API_PATH = "/api/version";

function fetchVersion(request: Request): Response | null {
  const url = new URL(request.url);
  if (url.pathname !== VERSION_API_PATH) return null;

  const headers = new Headers({
    "Access-Control-Allow-Methods": "GET, HEAD, OPTIONS",
    "Access-Control-Allow-Origin": "*",
    "Cache-Control": "no-store",
    "Content-Type": "application/json; charset=utf-8",
    "X-Content-Type-Options": "nosniff",
  });

  if (request.method === "OPTIONS") {
    return new Response(null, { status: 204, headers });
  }

  if (request.method !== "GET" && request.method !== "HEAD") {
    headers.set("Allow", "GET, HEAD, OPTIONS");
    return Response.json(
      { error: "Method Not Allowed" },
      { status: 405, headers },
    );
  }

  const body = JSON.stringify({
    name: SDK_PACKAGE_NAME,
    version: ONLYOFFICE_EMBED_SDK_VERSION,
    release: `sdk-v${ONLYOFFICE_EMBED_SDK_VERSION}`,
  });

  return new Response(request.method === "HEAD" ? null : body, {
    status: 200,
    headers,
  });
}

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
      fetchVersion(request) ??
      (await fetchStaticResource(request)) ??
      (await handler.fetch(request, options));
    const headers = new Headers(response.headers);
    headers.set("Origin-Agent-Cluster", "?1");

    return new Response(response.body, {
      status: response.status,
      statusText: response.statusText,
      headers,
    });
  },
});
