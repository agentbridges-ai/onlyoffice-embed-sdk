import { readFile } from "node:fs/promises";
import path from "node:path";
import { expect, test } from "playwright/test";

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
});
