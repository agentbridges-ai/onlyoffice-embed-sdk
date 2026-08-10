import { createFileRoute } from "@tanstack/react-router";
import { RuntimeRegressionsE2EPage } from "../../../tests/e2e/specs/runtime-regressions.page";

export const Route = createFileRoute("/e2e/runtime-regressions")({
  component: RuntimeRegressionsE2EPage,
});
