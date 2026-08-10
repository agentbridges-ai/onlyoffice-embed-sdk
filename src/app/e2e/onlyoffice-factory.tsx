import { createFileRoute } from "@tanstack/react-router";
import { OnlyOfficeFactoryE2EPage } from "../../../tests/e2e/specs/onlyoffice-factory.page";

export const Route = createFileRoute("/e2e/onlyoffice-factory")({
  component: OnlyOfficeFactoryE2EPage,
});
