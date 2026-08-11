import { createFileRoute } from "@tanstack/react-router";
import { OnlyOfficeCompatSubframePage } from "@/features/compat-subframe/onlyoffice-compat-subframe-page";
import { OnlyOfficeSubframePage } from "@/features/demo/onlyoffice-subframe-page";

export const Route = createFileRoute("/subframe")({
  validateSearch: (search): { runtime?: string } => ({
    runtime: typeof search.runtime === "string" ? search.runtime : undefined,
  }),
  component: SubframeRoute,
});

function SubframeRoute() {
  const { runtime } = Route.useSearch();
  return runtime === "compat" ? (
    <OnlyOfficeCompatSubframePage />
  ) : (
    <OnlyOfficeSubframePage />
  );
}
