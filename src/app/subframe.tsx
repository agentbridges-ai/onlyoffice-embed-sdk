import { createFileRoute } from "@tanstack/react-router";
import { OnlyOfficeSubframePage } from "@/features/demo/onlyoffice-subframe-page";

export const Route = createFileRoute("/subframe")({
  component: OnlyOfficeSubframePage,
});
