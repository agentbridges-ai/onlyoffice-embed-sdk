import { expect, test } from "playwright/test";

const expectedSteps = [
  "x2t worker lifecycle",
  "x2t queue and timeout",
  "editor server latest wins",
  "cross-origin bridge isolation",
  "bridge lifecycle races",
  "plugin config proxy allowlist",
  "Editor.bin source detection",
  "compatibility facade contracts",
  "compatibility native output callbacks",
  "cross-document compat mount",
];

test("runtime regression contracts", async ({ page }) => {
  await page.goto("/e2e/runtime-regressions", {
    waitUntil: "domcontentloaded",
  });
  await page.waitForFunction(
    (expectedStepCount) => {
      const status =
        document.querySelector('[data-testid="regression-status"]')
          ?.textContent ?? "";
      const rawResult =
        document.querySelector('[data-testid="regression-result"]')
          ?.textContent ?? "[]";
      let steps: unknown = [];
      try {
        steps = JSON.parse(rawResult);
      } catch {
        return false;
      }
      return (
        ["passed", "failed"].includes(status) &&
        Array.isArray(steps) &&
        steps.length === expectedStepCount
      );
    },
    expectedSteps.length,
    { timeout: 10_000 },
  );

  const status = await page.getByTestId("regression-status").innerText();
  const steps = JSON.parse(
    await page.getByTestId("regression-result").innerText(),
  ) as Array<{ name: string; status: string; detail?: string }>;

  expect(steps.map((step) => step.name)).toEqual(expectedSteps);
  expect(steps, JSON.stringify(steps, null, 2)).toEqual(
    expectedSteps.map((name) => ({ name, status: "passed" })),
  );
  expect(status).toBe("passed");
});
