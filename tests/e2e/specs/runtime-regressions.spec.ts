import { expect, test } from "playwright/test";

const expectedSteps = [
  "x2t worker lifecycle",
  "x2t queue and timeout",
  "x2t isolated workers",
  "editor server latest wins",
  "cross-origin bridge isolation",
  "bridge lifecycle races",
  "plugin config proxy allowlist",
  "Editor.bin source detection",
  "compatibility facade contracts",
  "native preview print logo configuration",
  "compatibility native output callbacks",
  "cross-document compat mount",
];

test("runtime regression contracts", async ({ page }) => {
  test.setTimeout(150_000);
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
    // The page executes heavyweight browser/runtime checks serially. Shared CI
    // runners can pause for substantially longer than local production
    // previews, so keep every result assertion strict while allowing the
    // aggregate harness enough scheduling slack to settle.
    { timeout: 120_000 },
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
