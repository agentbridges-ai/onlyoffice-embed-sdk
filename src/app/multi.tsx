import { createFileRoute, redirect } from "@tanstack/react-router";

export const Route = createFileRoute("/multi")({
  beforeLoad: () => {
    throw redirect({ to: "/docs/demos/multi" });
  },
});
