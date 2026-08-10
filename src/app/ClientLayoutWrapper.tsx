'use client'

import { usePathname } from "next/navigation";
import { SiteLayout } from '@/features/shell'

export default function ClientLayoutWrapper({
  children,
}: {
  children: React.ReactNode
}) {
  const pathname = usePathname();

  if (pathname === "/subframe") {
    return <>{children}</>;
  }

  return <SiteLayout>{children}</SiteLayout>
}
