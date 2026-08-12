import { readFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { brotliDecompressSync } from "node:zlib";
import path from "node:path";
import { expect, test } from "playwright/test";
import { ONLYOFFICE_X2T_RELEASE } from "../../../src/components/onlyoffice-embed-sdk/compat/version";
import {
  OFFICE_THEME,
  OFFICE_THEME_OPTIONS,
  STATIC_RESOURCE,
} from "../../../src/components/onlyoffice-embed-sdk/const";

function sha256(value: Uint8Array) {
  return createHash("sha256").update(value).digest("hex");
}

test("canonical cross-origin bridge matches the deployed asset", async () => {
  const root = process.cwd();
  const canonical = await readFile(
    path.join(
      root,
      "scripts/assets/onlyoffice/onlyoffice-cross-origin-bridge.js",
    ),
  );
  const deployed = await readFile(
    path.join(
      root,
      "public/packages/onlyoffice/9.4.0-develop/web-apps/vendor/onlyoffice-cross-origin-bridge.js",
    ),
  );

  expect(deployed.equals(canonical)).toBe(true);
  const source = canonical.toString("utf8");
  expect(source).toContain("event.source !== window.parent");
  expect(source).toContain("event.origin !== parentOrigin");
  expect(source).toContain("message.bridgeInstanceId !== bridgeInstanceId");
  expect(source).toContain("body.byteOffset + body.byteLength");
  expect(source).toContain("parentOrigin,");
  expect(source).toContain(
    'parsed.pathname.slice(-"/config.json".length) === "/config.json"',
  );
});

test("deployed x2t bytes match the immutable release lock", async () => {
  const root = process.cwd();
  const releaseDirectory = path.join(
    root,
    "public/packages/onlyoffice/x2t/v9.3.0+4",
  );
  const lockBytes = await readFile(
    path.join(root, "config/x2t-release.lock.json"),
  );
  const deployedLock = await readFile(
    path.join(releaseDirectory, "release-lock.json"),
  );
  const lock = JSON.parse(lockBytes.toString("utf8"));

  expect(deployedLock.equals(lockBytes)).toBe(true);
  expect(lock.tag).toBe(ONLYOFFICE_X2T_RELEASE.tag);
  expect(lock.sourceCommit).toBe(ONLYOFFICE_X2T_RELEASE.sourceCommit);
  expect(STATIC_RESOURCE.x2t.root).toBe(
    "/packages/onlyoffice/x2t/v9.3.0+4",
  );

  const script = await readFile(path.join(releaseDirectory, "x2t.js"));
  const compressedWasm = await readFile(
    path.join(releaseDirectory, "x2t.wasm"),
  );
  const rawWasm = brotliDecompressSync(compressedWasm);

  expect(sha256(script)).toBe(ONLYOFFICE_X2T_RELEASE.files.scriptSha256);
  expect(sha256(compressedWasm)).toBe(
    ONLYOFFICE_X2T_RELEASE.files.wasmBrotliSha256,
  );
  expect(sha256(rawWasm)).toBe(
    ONLYOFFICE_X2T_RELEASE.files.wasmRawSha256,
  );
});

test("hosted runtime exposes only modern interface themes", async () => {
  const root = process.cwd();
  const runtimeRoot = path.join(
    root,
    "public/packages/onlyoffice/9.4.0-develop",
  );
  const forbiddenThemeClass =
    /(^|[^a-z0-9_-])\.theme-(?:system|light|classic-light|dark|contrast-dark)(?![a-z0-9_-])/i;
  const forbiddenThemeIds = [
    "theme-system",
    "theme-light",
    "theme-classic-light",
    "theme-dark",
    "theme-contrast-dark",
  ];

  expect(OFFICE_THEME).toEqual({
    WHITE: "theme-white",
    NIGHT: "theme-night",
  });
  expect(OFFICE_THEME_OPTIONS.map(({ id }) => id)).toEqual([
    "theme-white",
    "theme-night",
  ]);

  for (const editor of [
    "documenteditor",
    "spreadsheeteditor",
    "presentationeditor",
  ]) {
    const appSource = await readFile(
      path.join(runtimeRoot, "web-apps/apps", editor, "main/app.js"),
      "utf8",
    );
    const cssSource = await readFile(
      path.join(
        runtimeRoot,
        "web-apps/apps",
        editor,
        "main/resources/css/app.css",
      ),
      "utf8",
    );
    const moduleStart = appSource.indexOf(
      'define("common/main/lib/controller/Themes"',
    );
    const moduleEnd = appSource.indexOf(
      'define("common/main/lib/util/utils"',
      moduleStart,
    );
    const themeController = appSource.slice(moduleStart, moduleEnd);

    expect(moduleStart).toBeGreaterThanOrEqual(0);
    expect(moduleEnd).toBeGreaterThan(moduleStart);
    expect(themeController).toContain('"theme-white":');
    expect(themeController).toContain('"theme-night":');
    for (const id of forbiddenThemeIds) {
      expect(themeController).not.toContain(`"${id}":`);
    }
    expect(cssSource).toContain(".theme-white");
    expect(cssSource).toContain(".theme-night");
    expect(forbiddenThemeClass.test(cssSource)).toBe(false);
  }

  for (const relativePath of [
    "themes.json",
    "web-apps/apps/common/main/resources/themes/themes.json",
  ]) {
    const config = JSON.parse(
      await readFile(path.join(runtimeRoot, relativePath), "utf8"),
    );
    expect(config).toEqual({ themes: [] });
  }
});
