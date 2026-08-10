import { createFileRoute, redirect } from "@tanstack/react-router";

export const Route = createFileRoute("/docs/demos/")({
  beforeLoad: () => {
    throw redirect({ to: "/docs/demos/single" });
  },
});
