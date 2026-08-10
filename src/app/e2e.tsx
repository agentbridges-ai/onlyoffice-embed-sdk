import { Outlet, createFileRoute } from "@tanstack/react-router";

export const Route = createFileRoute("/e2e")({
  component: E2ELayout,
});

function E2ELayout() {
  return <Outlet />;
}
