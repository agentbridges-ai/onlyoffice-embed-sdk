import { Outlet, createFileRoute } from "@tanstack/react-router";

export const Route = createFileRoute("/docs/demos")({
  component: DemosLayout,
});

function DemosLayout() {
  return <Outlet />;
}
