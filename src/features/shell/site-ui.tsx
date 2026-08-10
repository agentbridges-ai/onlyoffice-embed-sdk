import { Link } from "@tanstack/react-router";
import type { ButtonHTMLAttributes, ReactNode } from "react";

export const SITE_GITHUB =
  "https://github.com/agentbridges-ai/onlyoffice-embed-sdk";

const buttonBase =
  "inline-flex h-9 items-center justify-center border px-4 text-[13px] leading-none transition-colors duration-200 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-neutral-900";

export function SiteButtonPrimary({
  className = "",
  ...props
}: ButtonHTMLAttributes<HTMLButtonElement>) {
  return (
    <button
      type="button"
      className={`${buttonBase} border-neutral-900 bg-neutral-900 text-white hover:bg-neutral-800 ${className}`}
      {...props}
    />
  );
}

export function SiteButtonSecondary({
  className = "",
  ...props
}: ButtonHTMLAttributes<HTMLButtonElement>) {
  return (
    <button
      type="button"
      className={`${buttonBase} border-neutral-300 bg-white text-neutral-800 hover:border-neutral-500 hover:bg-neutral-50 ${className}`}
      {...props}
    />
  );
}

type SiteLinkButtonProps = {
  href: string;
  external?: boolean;
  variant?: "primary" | "secondary";
  children: ReactNode;
  className?: string;
  onClick?: () => void;
};

export function SiteLinkButton({
  href,
  external,
  variant = "primary",
  children,
  className = "",
  onClick,
}: SiteLinkButtonProps) {
  const classes =
    variant === "primary"
      ? `${buttonBase} border-neutral-900 bg-neutral-900 text-white hover:bg-neutral-800 ${className}`
      : `${buttonBase} border-neutral-300 bg-white text-neutral-800 hover:border-neutral-500 hover:bg-neutral-50 ${className}`;

  if (external) {
    return (
      <a
        href={href}
        target="_blank"
        rel="noopener noreferrer"
        className={classes}
        onClick={onClick}
      >
        {children}
      </a>
    );
  }

  return (
    <Link to={href} className={classes} onClick={onClick}>
      {children}
    </Link>
  );
}

export function SiteSectionTitle({
  children,
  className = "",
}: {
  children: ReactNode;
  className?: string;
}) {
  return (
    <h2
      className={`text-xl font-semibold tracking-[-0.02em] text-neutral-950 sm:text-2xl ${className}`}
    >
      {children}
    </h2>
  );
}

export function SiteCard({
  children,
  className = "",
}: {
  children: ReactNode;
  className?: string;
}) {
  return (
    <div
      className={`border border-neutral-200 bg-white p-5 sm:p-6 ${className}`}
    >
      {children}
    </div>
  );
}
