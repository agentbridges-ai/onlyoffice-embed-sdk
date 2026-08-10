"use client";

/**
 * 站点顶栏：桌面端横向导航；移动端左右布局（Logo | 菜单），
 * 展开时为全屏 fixed 合成层（100dvh）。导航项见 {@link ./site-nav.ts}。
 */
import { Link, useLocation } from "@tanstack/react-router";
import { useEffect, useState } from "react";
import { createPortal } from "react-dom";
import { BrandLogo } from "./brand-logo";
import { SITE_GITHUB, SiteLinkButton } from "./site-ui";
import { SITE_NAV } from "./site-nav";

function MenuToggleIcon({ open }: { open: boolean }) {
  return (
    <svg
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.5"
      className="h-5 w-5"
      aria-hidden
    >
      {open ? (
        <path strokeLinecap="square" d="M6 6l12 12M18 6L6 18" />
      ) : (
        <>
          <path strokeLinecap="square" d="M4 7h16M4 12h16M4 17h16" />
        </>
      )}
    </svg>
  );
}

function NavLinks({
  pathname,
  onNavigate,
  className,
  variant = "desktop",
}: {
  pathname: string;
  onNavigate?: () => void;
  className?: string;
  variant?: "desktop" | "mobile";
}) {
  return (
    <nav className={className} aria-label="主导航">
      {SITE_NAV.map((item) => {
        const active = item.isActive(pathname);
        const linkClass =
          variant === "mobile"
            ? `block rounded-sm px-3 py-3.5 text-[15px] ${
                active
                  ? "bg-neutral-50 font-medium text-neutral-950"
                  : "text-neutral-700 hover:bg-neutral-50"
              }`
            : `shrink-0 px-2.5 py-2 text-[13px] sm:px-3 ${
                active
                  ? "font-medium text-neutral-950"
                  : "text-neutral-500 hover:text-neutral-900"
              }`;

        return (
          <Link
            key={item.href}
            to={item.href}
            onClick={onNavigate}
            className={linkClass}
          >
            {item.label}
          </Link>
        );
      })}
    </nav>
  );
}

function MobileMenuLayer({
  open,
  pathname,
  onClose,
}: {
  open: boolean;
  pathname: string;
  onClose: () => void;
}) {
  const [mounted, setMounted] = useState(false);

  useEffect(() => {
    setMounted(true);
  }, []);

  if (!open || !mounted) return null;

  return createPortal(
    <div
      id="site-mobile-menu"
      className="fixed inset-0 z-[60] flex h-dvh max-h-dvh flex-col bg-white md:hidden"
      role="dialog"
      aria-modal="true"
      aria-label="站点菜单"
    >
      <div className="flex h-14 shrink-0 items-center justify-between border-b border-neutral-200 px-4 sm:px-6">
        <BrandLogo />
        <button
          type="button"
          className="inline-flex h-9 w-9 items-center justify-center border border-neutral-200 text-neutral-700"
          aria-label="关闭菜单"
          onClick={onClose}
        >
          <MenuToggleIcon open />
        </button>
      </div>

      <div className="flex min-h-0 flex-1 flex-col overflow-y-auto">
        <NavLinks
          pathname={pathname}
          onNavigate={onClose}
          variant="mobile"
          className="flex flex-col px-2 py-4"
        />

        <div className="mt-auto border-t border-neutral-100 px-4 py-4">
          <div className="flex flex-col gap-2">
            <SiteLinkButton
              href={SITE_GITHUB}
              external
              variant="secondary"
              className="h-10 w-full justify-center px-3"
              onClick={onClose}
            >
              GitHub
            </SiteLinkButton>
            <SiteLinkButton
              href="/docs/demos/single"
              className="h-10 w-full justify-center px-3"
              onClick={onClose}
            >
              开始使用
            </SiteLinkButton>
          </div>
        </div>
      </div>
    </div>,
    document.body,
  );
}

export function SiteHeader() {
  const pathname = useLocation().pathname;
  const [menuOpen, setMenuOpen] = useState(false);

  useEffect(() => {
    setMenuOpen(false);
  }, [pathname]);

  useEffect(() => {
    document.body.style.overflow = menuOpen ? "hidden" : "";
    return () => {
      document.body.style.overflow = "";
    };
  }, [menuOpen]);

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") setMenuOpen(false);
    };
    if (menuOpen) {
      window.addEventListener("keydown", onKeyDown);
    }
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [menuOpen]);

  return (
    <>
      <header className="sticky top-0 z-50 border-b border-neutral-200 bg-white/95 backdrop-blur-sm">
        <div className="mx-auto flex h-14 max-w-7xl items-center justify-between gap-4 px-4 sm:px-6 md:justify-start">
          <BrandLogo />

          <NavLinks
            pathname={pathname}
            className="hidden min-w-0 flex-1 items-center gap-0.5 md:flex"
          />

          <div className="hidden shrink-0 items-center gap-2 md:flex">
            <SiteLinkButton
              href={SITE_GITHUB}
              external
              variant="secondary"
              className="hidden h-8 px-3 sm:inline-flex"
            >
              GitHub
            </SiteLinkButton>
            <SiteLinkButton href="/docs/demos/single" className="h-8 px-3.5">
              开始使用
            </SiteLinkButton>
          </div>

          <button
            type="button"
            className="inline-flex h-9 w-9 shrink-0 items-center justify-center border border-neutral-200 text-neutral-700 md:hidden"
            aria-expanded={menuOpen}
            aria-controls="site-mobile-menu"
            aria-label={menuOpen ? "关闭菜单" : "打开菜单"}
            onClick={() => setMenuOpen((open) => !open)}
          >
            <MenuToggleIcon open={menuOpen} />
          </button>
        </div>
      </header>

      <MobileMenuLayer
        open={menuOpen}
        pathname={pathname}
        onClose={() => setMenuOpen(false)}
      />
    </>
  );
}
