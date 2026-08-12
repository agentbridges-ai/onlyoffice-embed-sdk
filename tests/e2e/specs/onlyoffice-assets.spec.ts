import { readFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { brotliDecompressSync } from "node:zlib";
import path from "node:path";
import { expect, test } from "playwright/test";
import { ONLYOFFICE_X2T_RELEASE } from "../../../src/components/onlyoffice-embed-sdk/compat/version";
import { STATIC_RESOURCE } from "../../../src/components/onlyoffice-embed-sdk/const";

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
