import { expect, test } from "playwright/test";

test("multi-instance editors use isolated subframe origins", async ({ page }) => {
  const response = await page.goto("/docs/demos/multi", {
    waitUntil: "domcontentloaded",
  });

  expect(response).not.toBeNull();
  expect(response?.headers()["origin-agent-cluster"]).toBe("?1");

  await page
    .locator('iframe[data-onlyoffice-subframe="true"]')
    .first()
    .waitFor();
  await page.getByRole("button", { name: "+ Excel" }).click();
  await page.getByRole("button", { name: "+ PPT" }).click();

  const subframes = page.locator('iframe[data-onlyoffice-subframe="true"]');
  await expect(subframes).toHaveCount(3);

  const origins = await subframes.evaluateAll((frames) =>
    frames.map((frame) => new URL((frame as HTMLIFrameElement).src).origin),
  );
  const port = new URL(page.url()).port;
  expect(origins).toEqual([
    `http://a.b.localhost:${port}`,
    `http://b.b.localhost:${port}`,
    `http://c.b.localhost:${port}`,
  ]);

  for (const origin of origins) {
    const childResponse = await page.request.get(`${origin}/subframe`);
    expect(childResponse.headers()["origin-agent-cluster"]).toBe("?1");
  }
});
