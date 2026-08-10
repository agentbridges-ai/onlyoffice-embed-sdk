import { Outlet, createFileRoute } from "@tanstack/react-router";
import { DocsShell } from "@/features/docs/components/docs-shell";

export const Route = createFileRoute("/docs")({
  component: DocsLayout,
});

function DocsLayout() {
  return (
    <DocsShell>
      <Outlet />
    </DocsShell>
  );
}
