import { useLocation } from "@tanstack/react-router";
import { SiteLayout } from "@/features/shell";

export default function ClientLayoutWrapper({
  children,
}: {
  children: React.ReactNode;
}) {
  const pathname = useLocation().pathname;

  if (pathname === "/subframe") {
    return <>{children}</>;
  }

  return <SiteLayout>{children}</SiteLayout>;
}
