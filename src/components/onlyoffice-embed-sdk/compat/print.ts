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

export function openOfficePrintWindow(
  ownerWindow: Window,
  ownerDocument: Document,
) {
  assertTrustedPrintOrigin(ownerWindow, ownerDocument);
  return ownerWindow.open("", "_blank");
}

export async function printOfficePdfFile(options: {
  ownerWindow: Window;
  ownerDocument: Document;
  file: File;
  printWindow: Window | null;
}) {
  const { ownerWindow, ownerDocument, file, printWindow } = options;
  const URLConstructor =
    (ownerWindow as Window & { URL?: typeof URL }).URL || URL;
  const objectUrl = URLConstructor.createObjectURL(file);
  const frame = printWindow ? null : ownerDocument.createElement("iframe");
  if (frame) {
    frame.style.cssText =
      "position:fixed;left:-10000px;top:0;width:1px;height:1px;border:0;opacity:0;pointer-events:none";
  }

  try {
    const header = new TextDecoder("ascii").decode(
      (await file.arrayBuffer()).slice(0, 5),
    );
    if (header !== "%PDF-") {
      throw new Error("Printable output is not a PDF");
    }

    if (frame) {
      ownerDocument.body.appendChild(frame);
      await new Promise<void>((resolve, reject) => {
        const timeout = ownerWindow.setTimeout(
          () => reject(new Error("Timed out creating the print frame")),
          5_000,
        );
        frame.onload = () => {
          ownerWindow.clearTimeout(timeout);
          resolve();
        };
        frame.onerror = () => {
          ownerWindow.clearTimeout(timeout);
          reject(new Error("Failed to create the print frame"));
        };
        frame.src = "about:blank";
      });
    }

    const targetWindow = printWindow || frame?.contentWindow;
    if (!targetWindow) throw new Error("Failed to create a printable window");
    const targetDocument = targetWindow.document;
    targetDocument.title = `Print ${file.name}`;
    targetDocument.documentElement.style.cssText =
      "width:100%;height:100%;margin:0";
    targetDocument.body.style.cssText = "width:100%;height:100%;margin:0";
    const embed = targetDocument.createElement("embed");
    embed.type = "application/pdf";
    embed.src = objectUrl;
    embed.style.cssText = "width:100%;height:100%;border:0";
    targetDocument.body.replaceChildren(embed);

    await new Promise<void>((resolve, reject) => {
      let settled = false;
      const finish = (error?: Error) => {
        if (settled) return;
        settled = true;
        ownerWindow.clearTimeout(timeout);
        error ? reject(error) : resolve();
      };
      embed.onload = () => ownerWindow.setTimeout(() => finish(), 250);
      embed.onerror = () => finish(new Error("Failed to load printable PDF"));
      const timeout = ownerWindow.setTimeout(() => finish(), 2_500);
    });
    targetWindow.focus();
    targetWindow.print();
  } catch (error) {
    printWindow?.close();
    frame?.remove();
    URLConstructor.revokeObjectURL(objectUrl);
    throw error;
  }

  ownerWindow.setTimeout(() => {
    frame?.remove();
    URLConstructor.revokeObjectURL(objectUrl);
  }, 60_000);
}
