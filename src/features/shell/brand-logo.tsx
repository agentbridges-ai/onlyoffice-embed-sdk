import { Link } from "@tanstack/react-router";

type BrandLogoProps = {
  compact?: boolean;
  href?: string;
};

export function BrandMark({ className = "h-8 w-8" }: { className?: string }) {
  return (
    <svg
      viewBox="0 0 32 32"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
      className={className}
      aria-hidden
    >
      <rect x="5" y="7" width="14" height="18" fill="#fff" stroke="#171717" strokeWidth="1.5" />
      <rect x="11" y="4" width="14" height="18" fill="#fff" stroke="#171717" strokeWidth="1.5" />
      <path d="M14 12h8M14 15h6M14 18h8" stroke="#525252" strokeWidth="1.2" strokeLinecap="square" />
      <path d="M8 12h4M8 15h3M8 18h4" stroke="#a3a3a3" strokeWidth="1" strokeLinecap="square" />
    </svg>
  );
}

export function BrandLogo({ compact = false, href = "/" }: BrandLogoProps) {
  return (
    <Link to={href} className="inline-flex items-center gap-2.5">
      <BrandMark />
      {!compact && (
        <span className="flex items-baseline gap-2">
          <span className="text-[15px] font-semibold tracking-[-0.02em] text-neutral-950">
            OnlyOffice
          </span>
          <span className="hidden h-3.5 w-px bg-neutral-300 sm:block" />
          <span className="hidden text-[13px] text-neutral-500 sm:inline">
            Embed SDK
          </span>
        </span>
      )}
    </Link>
  );
}
