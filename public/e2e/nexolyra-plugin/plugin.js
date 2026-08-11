(function (ownerWindow) {
  "use strict";

  var protocol = "onlyoffice-browser-plugin/v1";
  var pluginGuid = "asc.{E2E4D0B6-6F1E-4B80-9A4D-8F6B1C2D3E40}";
  var bridgeHost = ownerWindow.parent.parent;
  var pluginInstanceId =
    "e2e-plugin-" + Date.now() + "-" + Math.random().toString(36).slice(2);
  var readyPosted = false;

  ownerWindow.Asc.plugin.init = function () {
    announceReady();
  };
  ownerWindow.Asc.plugin.button = function () {};

  var readyPoll = ownerWindow.setInterval(function () {
    if (announceReady()) ownerWindow.clearInterval(readyPoll);
  }, 100);
  ownerWindow.setTimeout(function () {
    ownerWindow.clearInterval(readyPoll);
  }, 30000);

  function announceReady() {
    var editorType = String(
      (ownerWindow.Asc.plugin.info &&
        ownerWindow.Asc.plugin.info.editorType) ||
        "",
    );
    if (
      readyPosted ||
      !editorType ||
      typeof ownerWindow.Asc.plugin.executeMethod !== "function"
    ) {
      return readyPosted;
    }
    readyPosted = true;
    post({ type: "READY", editorType: editorType });
    return true;
  }

  ownerWindow.addEventListener("message", function (event) {
    var message = event.data;
    if (
      event.source !== bridgeHost ||
      event.origin !== ownerWindow.location.origin ||
      !message ||
      message.protocol !== protocol ||
      message.pluginGuid !== pluginGuid ||
      message.pluginInstanceId !== pluginInstanceId ||
      message.type !== "INVOKE" ||
      typeof message.requestId !== "string"
    ) {
      return;
    }

    post({
      type: "RESULT",
      requestId: message.requestId,
      ok: true,
      result: {
        pong: message.payload && message.payload.type === "ping",
        editorType: String(ownerWindow.Asc.plugin.info.editorType || ""),
        bridgeHostIsTop: bridgeHost === ownerWindow.top,
        entryPath: ownerWindow.location.pathname,
      },
    });
  });

  function post(message) {
    bridgeHost.postMessage(
      Object.assign(
        {
          protocol: protocol,
          pluginGuid: pluginGuid,
          pluginInstanceId: pluginInstanceId,
        },
        message,
      ),
      ownerWindow.location.origin,
    );
  }
})(window);
