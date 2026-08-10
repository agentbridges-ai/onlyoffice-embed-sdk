import { DemoButton, DemoField } from "./demo-toolbar";
import { OnlyOfficeManager } from "@/components/onlyoffice-web-comp";

export const DEFAULT_DEMO_CDN_ORIGIN =
  "https://d69d942c.onlyoffice-embed-resource.pages.dev";

export type DemoResourceMode = "local" | "cdn";

type DemoResourceState = {
  mode: DemoResourceMode;
  cdnOrigin: string;
  revision: number;
};

const DEMO_RESOURCE_CHANGE_EVENT = "onlyoffice-demo-resource-change";

let demoResourceState: DemoResourceState = {
  mode: "local",
  cdnOrigin: DEFAULT_DEMO_CDN_ORIGIN,
  revision: 0,
};

export function getDemoResourceState() {
  return demoResourceState;
}

export function applyDemoResourceMode(
  mode: DemoResourceMode,
  cdnOrigin = demoResourceState.cdnOrigin,
) {
  if (mode === "local" && demoResourceState.mode === "local") {
    return demoResourceState;
  }

  if (mode === "cdn") {
    OnlyOfficeManager.registerStaticResource({ cdnOrigin });
  } else {
    OnlyOfficeManager.resetStaticResource();
  }

  demoResourceState = {
    mode,
    cdnOrigin,
    revision: demoResourceState.revision + 1,
  };

  window.dispatchEvent(
    new CustomEvent<DemoResourceState>(DEMO_RESOURCE_CHANGE_EVENT, {
      detail: demoResourceState,
    }),
  );

  return demoResourceState;
}

export function subscribeDemoResourceChange(
  listener: (state: DemoResourceState) => void,
) {
  const handler = (event: Event) => {
    listener((event as CustomEvent<DemoResourceState>).detail);
  };
  window.addEventListener(DEMO_RESOURCE_CHANGE_EVENT, handler);
  return () => window.removeEventListener(DEMO_RESOURCE_CHANGE_EVENT, handler);
}

type ResourceSwitcherProps = {
  cdnOrigin: string;
  disabled?: boolean;
  onCdnOriginChange: (value: string) => void;
  onLoad: () => void;
};

export function ResourceSwitcher({
  cdnOrigin,
  disabled = false,
  onCdnOriginChange,
  onLoad,
}: ResourceSwitcherProps) {
  return (
    <div className="flex min-w-0 items-center gap-1.5">
      <DemoField label="资源" className="min-w-0 flex-1">
        <input
          value={cdnOrigin}
          disabled={disabled}
          onChange={(event) => onCdnOriginChange(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === "Enter") {
              onLoad();
            }
          }}
          className="h-6 w-[280px] max-w-[calc(100vw-160px)] min-w-0 border-0 bg-transparent px-0.5 text-[13px] text-neutral-800 outline-none disabled:opacity-60"
        />
      </DemoField>
      <DemoButton className="shrink-0" disabled={disabled} onClick={onLoad}>
        加载
      </DemoButton>
    </div>
  );
}
