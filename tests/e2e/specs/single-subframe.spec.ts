import { expect, test } from "playwright/test";

test("single-instance editor uses the rat subframe origin", async ({ page }) => {
  const response = await page.goto("/docs/demos/single", {
    waitUntil: "domcontentloaded",
  });

  expect(response).not.toBeNull();
  expect(response?.headers()["origin-agent-cluster"]).toBe("?1");

  const subframe = page.locator('iframe[data-onlyoffice-subframe="true"]');
  await subframe.waitFor();
  await expect(subframe).toHaveAttribute("data-onlyoffice-zodiac", "rat");

  const origin = await subframe.evaluate(
    (frame) => new URL((frame as HTMLIFrameElement).src).origin,
  );
  const port = new URL(page.url()).port;
  expect(origin).toBe(`http://rat.onlyoffice.localhost:${port}`);

  const childResponse = await page.request.get(`${origin}/subframe`);
  expect(childResponse.headers()["origin-agent-cluster"]).toBe("?1");
});
