import { expect, test } from "playwright/test";
import { ONLYOFFICE_EMBED_HOST_BUILD_ID } from "../../../src/components/onlyoffice-embed-sdk/compat/version";

const immutable = "public, max-age=31536000, immutable";

test("hosted resources use version-aware browser cache policies", async ({ request }) => {
  const subframe = await request.get("/subframe?runtime=compat");
  expect(subframe.status()).toBe(200);
  expect(subframe.headers()["cache-control"]).toBe("no-store");
  const html = await subframe.text();
  const hashedAsset = html.match(/(?:src|href)="(\/assets\/[^\"]+-[^\"]+)"/)?.[1];
  expect(hashedAsset, "subframe did not reference a hashed asset").toBeTruthy();

  const asset = await request.get(hashedAsset!);
  expect(asset.status()).toBe(200);
  expect(asset.headers()["cache-control"]).toBe(immutable);
  expect(asset.headers()["access-control-allow-origin"]).toBe("*");

  const appPort = Number(process.env.PLAYWRIGHT_APP_PORT ?? 3001);
  const zodiacSubframe = await request.get(
    `http://rat.onlyoffice.localhost:${appPort}/subframe?runtime=compat`,
  );
  expect(zodiacSubframe.status()).toBe(200);
  expect(await zodiacSubframe.text()).toContain(
    `http://onlyoffice.localhost:${appPort}/assets/`,
  );

  const runtimeApi = await request.get(
    `/onlyoffice/runtime/${ONLYOFFICE_EMBED_HOST_BUILD_ID}/web-apps/apps/api/documents/api.js`,
  );
  expect(runtimeApi.status()).toBe(200);
  expect(runtimeApi.headers()["cache-control"]).toBe(immutable);
  expect(runtimeApi.headers()["access-control-allow-origin"]).toBe("*");
  expect(runtimeApi.headers()["cross-origin-resource-policy"]).toBe(
    "cross-origin",
  );

  const zodiacRuntimeRoot =
    `http://rat.onlyoffice.localhost:${appPort}/onlyoffice/runtime/${ONLYOFFICE_EMBED_HOST_BUILD_ID}`;
  const zodiacApi = await request.get(
    `${zodiacRuntimeRoot}/web-apps/apps/api/documents/api.js`,
  );
  expect(zodiacApi.status()).toBe(200);
  expect(zodiacApi.headers()["cache-control"]).toBe(immutable);

  const zodiacEditorShell = await request.get(
    `${zodiacRuntimeRoot}/web-apps/apps/documenteditor/main/index.html`,
  );
  expect(zodiacEditorShell.status()).toBe(200);
  expect(zodiacEditorShell.headers()["cache-control"]).toBe("no-store");
  expect(await zodiacEditorShell.text()).toContain(
    `<base href="http://onlyoffice.localhost:${appPort}/onlyoffice/runtime/${ONLYOFFICE_EMBED_HOST_BUILD_ID}/web-apps/apps/documenteditor/main/">`,
  );

  const zodiacRuntimeAsset = await request.get(
    `${zodiacRuntimeRoot}/web-apps/vendor/requirejs/require.js`,
    { maxRedirects: 0 },
  );
  expect(zodiacRuntimeAsset.status()).toBe(308);
  expect(zodiacRuntimeAsset.headers()["cache-control"]).toBe(immutable);
  expect(zodiacRuntimeAsset.headers().location).toBe(
    `http://onlyoffice.localhost:${appPort}/onlyoffice/runtime/${ONLYOFFICE_EMBED_HOST_BUILD_ID}/web-apps/vendor/requirejs/require.js`,
  );

  const x2t = await request.get("/onlyoffice/x2t/v9.3.0+4/x2t.js");
  expect(x2t.status()).toBe(200);
  expect(x2t.headers()["cache-control"]).toBe(immutable);
  expect(x2t.headers()["access-control-allow-origin"]).toBe("*");

  const mutableSourcePath = await request.get(
    "/onlyoffice/9.4.0-develop/web-apps/apps/api/documents/api.js",
  );
  expect(mutableSourcePath.status()).toBe(200);
  expect(mutableSourcePath.headers()["cache-control"]).toBe(
    "public, max-age=3600, stale-while-revalidate=86400",
  );

  const version = await request.get("/api/version");
  expect(version.headers()["cache-control"]).toBe("no-store");
});
