import { defineConfig, devices } from "playwright/test";

const appPort = Number(process.env.PLAYWRIGHT_APP_PORT ?? 3001);
const cdnPort = Number(process.env.PLAYWRIGHT_CDN_PORT ?? 3010);
const host = process.env.PLAYWRIGHT_HOST ?? "127.0.0.1";
const isCI = !!process.env.CI;
const headless = true;
const appCommand = isCI
  ? `pnpm exec vite preview --host ${host} --port ${appPort}`
  : `pnpm exec vite dev --host ${host} --port ${appPort}`;

export default defineConfig({
  testDir: "./tests/e2e/specs",
  timeout: 60_000,
  expect: {
    timeout: 10_000,
  },
  fullyParallel: false,
  workers: 1,
  reporter: [["list"], ["html", { open: "never" }]],
  use: {
    baseURL: `http://${host}:${appPort}`,
    trace: "retain-on-failure",
    screenshot: "only-on-failure",
    video: "retain-on-failure",
    headless,
  },
  projects: [
    {
      name: "chromium",
      use: { ...devices["Desktop Chrome"] },
    },
  ],
  webServer: [
    {
      command: appCommand,
      url: `http://${host}:${appPort}`,
      reuseExistingServer: !isCI,
      timeout: 120_000,
    },
    {
      command: `node scripts/static-server.mjs public/packages --port ${cdnPort} --host ${host}`,
      url: `http://${host}:${cdnPort}/onlyoffice/9.4.0-develop/web-apps/apps/api/documents/api.js`,
      reuseExistingServer: !isCI,
      timeout: 120_000,
    },
  ],
});
