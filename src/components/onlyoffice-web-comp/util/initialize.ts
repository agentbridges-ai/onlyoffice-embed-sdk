import { STATIC_RESOURCE } from "../const";

type DocsApiWindow = Window & {
  DocsAPI?: {
    DocEditor?: unknown;
  };
};

type InitializeState = {
  promise: Promise<void> | null;
  apiUrl: string;
};

const initializeStates = new WeakMap<Window, InitializeState>();

function getInitializeState(ownerWindow: Window): InitializeState {
  let state = initializeStates.get(ownerWindow);
  if (!state) {
    state = { promise: null, apiUrl: "" };
    initializeStates.set(ownerWindow, state);
  }
  return state;
}

function resetLoadedDocsApi(
  ownerWindow: DocsApiWindow,
  state: InitializeState,
  apiUrl: string,
) {
  if (!state.apiUrl || state.apiUrl === apiUrl) {
    return;
  }

  ownerWindow.document
    .querySelectorAll<HTMLScriptElement>(
      'script[src*="/web-apps/apps/api/documents/api.js"]',
    )
    .forEach((script) => script.remove());
  ownerWindow.document
    .querySelectorAll<HTMLIFrameElement>("iframe[data-onlyoffice-preload]")
    .forEach((iframe) => iframe.remove());

  try {
    delete ownerWindow.DocsAPI;
  } catch {
    ownerWindow.DocsAPI = undefined;
  }
}

function preloadEditorFrame(ownerWindow: Window) {
  const ownerDocument = ownerWindow.document;
  if (
    ownerDocument.querySelector(
      `iframe[data-onlyoffice-preload="${STATIC_RESOURCE.onlyoffice.preloadHtml}"]`,
    )
  ) {
    return;
  }

  const iframe = ownerDocument.createElement("iframe");
  iframe.src = STATIC_RESOURCE.onlyoffice.preloadUrl;
  iframe.dataset.onlyofficePreload = STATIC_RESOURCE.onlyoffice.preloadHtml;
  iframe.className = "w-0 h-0 hidden absolute -z-10";
  ownerDocument.body.appendChild(iframe);
}

export async function initializeOnlyOffice(ownerWindow?: Window) {
  const targetWindow =
    ownerWindow ?? (typeof window === "undefined" ? undefined : window);
  if (!targetWindow) return;

  const apiUrl = STATIC_RESOURCE.onlyoffice.apiUrl;
  const state = getInitializeState(targetWindow);

  if (state.promise && state.apiUrl === apiUrl) {
    return state.promise;
  }

  if (state.promise && state.apiUrl !== apiUrl) {
    state.promise = null;
  }
  resetLoadedDocsApi(targetWindow, state, apiUrl);
  state.apiUrl = apiUrl;

  state.promise = new Promise<void>((resolve, reject) => {
    preloadEditorFrame(targetWindow);

    if (targetWindow.DocsAPI?.DocEditor) {
      resolve();
      return;
    }

    let script = targetWindow.document.querySelector<HTMLScriptElement>(
      `script[src="${apiUrl}"]`,
    );

    if (!script) {
      script = targetWindow.document.createElement("script");
      script.src = apiUrl;
      targetWindow.document.head.appendChild(script);
    }

    script.addEventListener("load", () => resolve(), { once: true });
    script.addEventListener(
      "error",
      () => {
        state.promise = null;
        reject(new Error("Failed to load OnlyOffice DocsAPI script"));
      },
      { once: true },
    );
  });

  return state.promise;
}
