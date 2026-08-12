import handler, { createServerEntry } from "@tanstack/react-start/server-entry";
import {
  ONLYOFFICE_EMBED_HOST_ASSET_DIGEST,
  ONLYOFFICE_EMBED_HOST_BUILD_ID,
  ONLYOFFICE_EMBED_HOST_MANIFEST,
  ONLYOFFICE_EMBED_SDK_VERSION,
} from "./components/onlyoffice-embed-sdk/compat/version";

const STATIC_RESOURCE_ORIGIN =
  "https://onlyoffice-embed-resource.pages.dev";
const SDK_PACKAGE_NAME = "@agentbridges-ai/onlyoffice-embed-sdk";
const VERSION_API_PATH = "/api/version";
const VERSIONED_RUNTIME_PREFIX =
  `/onlyoffice/runtime/${ONLYOFFICE_EMBED_HOST_BUILD_ID}`;
const SOURCE_RUNTIME_PREFIX =
  `/onlyoffice/${ONLYOFFICE_EMBED_HOST_MANIFEST.onlyofficeVersion}`;
const IMMUTABLE_CACHE_CONTROL = "public, max-age=31536000, immutable";
const REVALIDATED_CACHE_CONTROL =
  "public, max-age=3600, stale-while-revalidate=86400";
const ZODIAC_SLOT_PATTERN =
  /^(?:rat|ox|tiger|rabbit|dragon|snake|horse|goat|monkey|rooster|dog|pig)\./;

function canonicalStaticOrigin(url: URL) {
  if (
    ZODIAC_SLOT_PATTERN.test(url.hostname) &&
    url.hostname.endsWith(".onlyoffice.agent-bridges.com")
  ) {
    return "https://onlyoffice.agent-bridges.com";
  }
  if (
    ZODIAC_SLOT_PATTERN.test(url.hostname) &&
    url.hostname.endsWith(".onlyoffice.localhost")
  ) {
    return `${url.protocol}//onlyoffice.localhost${url.port ? `:${url.port}` : ""}`;
  }
  return null;
}

async function rewriteSubframeAssetUrls(
  request: Request,
  response: Response,
): Promise<Response> {
  const url = new URL(request.url);
  const canonicalOrigin = canonicalStaticOrigin(url);
  if (
    !canonicalOrigin ||
    url.pathname !== "/subframe" ||
    !response.headers.get("Content-Type")?.toLowerCase().startsWith("text/html")
  ) {
    return response;
  }
  const headers = new Headers(response.headers);
  headers.delete("Content-Length");
  const html = (await response.text()).replace(
    /(["'])\/assets\//g,
    `$1${canonicalOrigin}/assets/`,
  );
  return new Response(html, {
    status: response.status,
    statusText: response.statusText,
    headers,
  });
}

const EDITOR_SHELL_PATH = new RegExp(
  `^${VERSIONED_RUNTIME_PREFIX.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}` +
    "/web-apps/apps/(?:documenteditor|spreadsheeteditor|presentationeditor|pdfeditor|visioeditor|common)/" +
    "(?:main|embed|mobile|forms)/(?:index|index_loader|index_internal)\\.html$",
);

const DOCS_API_PATH =
  `${VERSIONED_RUNTIME_PREFIX}/web-apps/apps/api/documents/api.js`;

function redirectZodiacRuntimeAsset(request: Request): Response | null {
  const url = new URL(request.url);
  const canonicalOrigin = canonicalStaticOrigin(url);
  if (
    !canonicalOrigin ||
    !url.pathname.startsWith(`${VERSIONED_RUNTIME_PREFIX}/`) ||
    url.pathname === DOCS_API_PATH ||
    EDITOR_SHELL_PATH.test(url.pathname) ||
    (request.method !== "GET" && request.method !== "HEAD")
  ) {
    return null;
  }

  return new Response(null, {
    status: 308,
    headers: {
      "Cache-Control": IMMUTABLE_CACHE_CONTROL,
      Location: `${canonicalOrigin}${url.pathname}${url.search}`,
    },
  });
}

/**
 * Keep the real editor document on its zodiac origin while resolving its
 * relative, immutable UI/runtime assets against the canonical origin. The
 * document URL remains same-origin with the compatibility host; only
 * subresource URLs use the shared browser-cache key.
 */
async function rewriteEditorShellAssetBase(
  request: Request,
  response: Response,
): Promise<Response> {
  const url = new URL(request.url);
  const canonicalOrigin = canonicalStaticOrigin(url);
  if (
    !canonicalOrigin ||
    !EDITOR_SHELL_PATH.test(url.pathname) ||
    !response.headers.get("Content-Type")?.toLowerCase().startsWith("text/html")
  ) {
    return response;
  }

  const pathEnd = url.pathname.lastIndexOf("/") + 1;
  const baseHref = `${canonicalOrigin}${url.pathname.slice(0, pathEnd)}`;
  const html = (await response.text()).replace(
    /<head(\s[^>]*)?>/i,
    (head) => `${head}\n    <base href="${baseHref}">`,
  );
  const headers = new Headers(response.headers);
  headers.delete("Content-Length");
  return new Response(html, {
    status: response.status,
    statusText: response.statusText,
    headers,
  });
}

function cacheControlFor(request: Request, response: Response) {
  const url = new URL(request.url);
  const contentType = response.headers.get("Content-Type") || "";
  if (
    request.method !== "GET" && request.method !== "HEAD"
  ) {
    return "no-store";
  }
  if (response.status >= 300 && response.status < 400) {
    const canonicalOrigin = canonicalStaticOrigin(url);
    const location = response.headers.get("Location");
    if (
      canonicalOrigin &&
      url.pathname.startsWith(`${VERSIONED_RUNTIME_PREFIX}/`) &&
      location?.startsWith(`${canonicalOrigin}${VERSIONED_RUNTIME_PREFIX}/`)
    ) {
      return IMMUTABLE_CACHE_CONTROL;
    }
    return "no-store";
  }
  if (
    response.status >= 400 ||
    url.pathname === VERSION_API_PATH ||
    url.pathname === "/subframe" ||
    contentType.toLowerCase().startsWith("text/html")
  ) {
    return "no-store";
  }
  const immutableAsset =
    /^\/assets\/[^/?]+-[A-Za-z0-9_-]{8,}\.[A-Za-z0-9]+$/.test(
      url.pathname,
    ) ||
    url.pathname.startsWith(`${VERSIONED_RUNTIME_PREFIX}/`) ||
    /^\/(?:packages\/)?onlyoffice\/x2t\/v[^/]+\//.test(url.pathname);
  if (immutableAsset) return IMMUTABLE_CACHE_CONTROL;
  if (
    url.pathname.startsWith(`${SOURCE_RUNTIME_PREFIX}/`) ||
    url.pathname.startsWith(`/packages${SOURCE_RUNTIME_PREFIX}/`)
  ) {
    return REVALIDATED_CACHE_CONTROL;
  }
  return null;
}

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
    hostIdentity: {
      packageVersion: ONLYOFFICE_EMBED_SDK_VERSION,
      hostBuildId: ONLYOFFICE_EMBED_HOST_BUILD_ID,
      assetManifestDigest: ONLYOFFICE_EMBED_HOST_ASSET_DIGEST,
    },
    runtimeManifest: ONLYOFFICE_EMBED_HOST_MANIFEST,
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
  const sourcePath = resourcePath.startsWith(`${VERSIONED_RUNTIME_PREFIX}/`)
    ? `${SOURCE_RUNTIME_PREFIX}${resourcePath.slice(VERSIONED_RUNTIME_PREFIX.length)}`
    : resourcePath;
  const targetPath = sourcePath.endsWith("/index.html")
    ? sourcePath.slice(0, -"index.html".length)
    : sourcePath;
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
    let response =
      fetchVersion(request) ??
      redirectZodiacRuntimeAsset(request) ??
      (await fetchStaticResource(request)) ??
      (await handler.fetch(request, options));
    response = await rewriteSubframeAssetUrls(request, response);
    response = await rewriteEditorShellAssetBase(request, response);
    const headers = new Headers(response.headers);
    const cacheControl = cacheControlFor(request, response);
    if (cacheControl) headers.set("Cache-Control", cacheControl);
    const pathname = new URL(request.url).pathname;
    if (
      pathname.startsWith("/assets/") ||
      pathname.startsWith("/onlyoffice/") ||
      pathname.startsWith("/packages/onlyoffice/")
    ) {
      headers.set("Access-Control-Allow-Origin", "*");
      headers.set("Cross-Origin-Resource-Policy", "cross-origin");
      headers.set("Timing-Allow-Origin", "*");
    }
    headers.set("Origin-Agent-Cluster", "?1");

    return new Response(response.body, {
      status: response.status,
      statusText: response.statusText,
      headers,
    });
  },
});
