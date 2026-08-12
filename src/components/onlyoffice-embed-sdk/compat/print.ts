function assertTrustedPrintOrigin(ownerWindow: Window, ownerDocument: Document) {
  const candidates: string[] = [ownerDocument.baseURI, ownerDocument.referrer];
  try {
    if (ownerWindow.opener) candidates.push(ownerWindow.opener.location.href);
  } catch {
    // A cross-origin opener is never a trusted print peer.
  }
  try {
    candidates.push(ownerWindow.location.href);
  } catch {
    // Keep checking the inherited same-origin document coordinates.
  }
  for (const candidate of candidates) {
    if (!candidate) continue;
    try {
      const parsed = new URL(candidate);
      if (parsed.protocol === "http:" || parsed.protocol === "https:") return;
    } catch {
      // Try the next independently derived coordinate.
    }
  }
  throw new Error("Printing requires a trusted http(s) document origin");
}

const PRINT_PDF_LOAD_TIMEOUT_MS = 15_000;
const PRINT_PDF_LOAD_FALLBACK_MS = 3_000;
const PRINT_PDF_RENDER_SETTLE_MS = 1_000;

function waitForPdfNavigation(options: {
  ownerWindow: Window;
  frame: HTMLIFrameElement;
  navigate: () => void;
}) {
  const { ownerWindow, frame, navigate } = options;
  return new Promise<void>((resolve, reject) => {
    let settled = false;
    let renderTimer: number | undefined;
    let loadFallbackTimer: number | undefined;
    const finish = (error?: Error) => {
      if (settled) return;
      settled = true;
      ownerWindow.clearTimeout(timeout);
      if (renderTimer !== undefined) ownerWindow.clearTimeout(renderTimer);
      if (loadFallbackTimer !== undefined) {
        ownerWindow.clearTimeout(loadFallbackTimer);
      }
      frame.removeEventListener("load", handleLoad);
      frame.removeEventListener("error", handleError);
      error ? reject(error) : resolve();
    };
    const settleAfterRender = () => {
      if (renderTimer !== undefined) return;
      renderTimer = ownerWindow.setTimeout(
        () => finish(),
        PRINT_PDF_RENDER_SETTLE_MS,
      );
    };
    const handleLoad = () => {
      if (loadFallbackTimer !== undefined) {
        ownerWindow.clearTimeout(loadFallbackTimer);
      }
      settleAfterRender();
    };
    const handleError = () => finish(new Error("Failed to load printable PDF"));
    const timeout = ownerWindow.setTimeout(
      () => finish(new Error("Timed out loading printable PDF")),
      PRINT_PDF_LOAD_TIMEOUT_MS,
    );
    frame.addEventListener("load", handleLoad, { once: true });
    frame.addEventListener("error", handleError, { once: true });
    try {
      navigate();
      // Headless Chromium and some PDF plug-in configurations do not expose a
      // load event for a PDF document. Keep the direct-PDF target and use a
      // conservative fallback rather than reverting to a printable wrapper.
      loadFallbackTimer = ownerWindow.setTimeout(
        settleAfterRender,
        PRINT_PDF_LOAD_FALLBACK_MS,
      );
    } catch (error) {
      finish(error instanceof Error ? error : new Error(String(error)));
    }
  });
}

export async function printOfficePdfFile(options: {
  ownerWindow: Window;
  ownerDocument: Document;
  file: File;
}) {
  const { ownerWindow, ownerDocument, file } = options;
  assertTrustedPrintOrigin(ownerWindow, ownerDocument);
  const URLConstructor =
    (ownerWindow as Window & { URL?: typeof URL }).URL || URL;
  const objectUrl = URLConstructor.createObjectURL(file);
  const frame = ownerDocument.createElement("iframe");
  frame.dataset.onlyofficePrintTarget = "pdf";
  frame.setAttribute("aria-hidden", "true");
  frame.style.cssText =
    "position:fixed;left:-10000px;top:0;width:1px;height:1px;border:0;opacity:0;pointer-events:none";

  try {
    const header = new TextDecoder("ascii").decode(
      (await file.arrayBuffer()).slice(0, 5),
    );
    if (header !== "%PDF-") {
      throw new Error("Printable output is not a PDF");
    }

    frame.src = objectUrl;
    await waitForPdfNavigation({
      ownerWindow,
      frame,
      navigate: () => {
        ownerDocument.body.appendChild(frame);
      },
    });
    const targetWindow = frame.contentWindow;
    if (!targetWindow) throw new Error("Failed to create a printable window");
    targetWindow.focus();
    targetWindow.print();
  } catch (error) {
    frame.remove();
    URLConstructor.revokeObjectURL(objectUrl);
    throw error;
  }
  // window.print() is synchronous while Chromium's preview is open. Once it
  // returns, retaining the PDF browsing context can only make a later print
  // action observe the previous target, so dispose it immediately.
  frame.remove();
  URLConstructor.revokeObjectURL(objectUrl);
}
