import { expect, test } from "playwright/test";
import { createHash } from "node:crypto";
import {
  ONLYOFFICE_EMBED_HOST_ASSET_DIGEST,
  ONLYOFFICE_EMBED_HOST_BUILD_ID,
  ONLYOFFICE_EMBED_HOST_MANIFEST,
  ONLYOFFICE_EMBED_SDK_VERSION,
} from "../../../src/components/onlyoffice-embed-sdk/compat/version";

const expectedPayload = {
  name: "@agentbridges-ai/onlyoffice-embed-sdk",
  version: ONLYOFFICE_EMBED_SDK_VERSION,
  release: `sdk-v${ONLYOFFICE_EMBED_SDK_VERSION}`,
  hostIdentity: {
    packageVersion: ONLYOFFICE_EMBED_SDK_VERSION,
    hostBuildId: ONLYOFFICE_EMBED_HOST_BUILD_ID,
    assetManifestDigest: ONLYOFFICE_EMBED_HOST_ASSET_DIGEST,
  },
  runtimeManifest: ONLYOFFICE_EMBED_HOST_MANIFEST,
};

test("version API exposes the published SDK version", async ({ request }) => {
  expect(
    createHash("sha256")
      .update(JSON.stringify(ONLYOFFICE_EMBED_HOST_MANIFEST))
      .digest("hex"),
  ).toBe(ONLYOFFICE_EMBED_HOST_ASSET_DIGEST);
  const response = await request.get("/api/version");

  expect(response.status()).toBe(200);
  expect(response.headers()["access-control-allow-origin"]).toBe("*");
  expect(response.headers()["cache-control"]).toBe("no-store");
  expect(await response.json()).toEqual(expectedPayload);

  const head = await request.head("/api/version");
  expect(head.status()).toBe(200);
  expect(await head.body()).toHaveLength(0);

  const options = await request.fetch("/api/version", { method: "OPTIONS" });
  expect(options.status()).toBe(204);
  expect(options.headers()["access-control-allow-methods"]).toContain("GET");

  const unsupported = await request.post("/api/version");
  expect(unsupported.status()).toBe(405);
  expect(unsupported.headers().allow).toBe("GET, HEAD, OPTIONS");
});

test("home page displays the API version", async ({ page }) => {
  await page.goto("/");

  await expect(page.getByTestId("sdk-version")).toHaveText(
    `SDK v${ONLYOFFICE_EMBED_SDK_VERSION}`,
  );
});
