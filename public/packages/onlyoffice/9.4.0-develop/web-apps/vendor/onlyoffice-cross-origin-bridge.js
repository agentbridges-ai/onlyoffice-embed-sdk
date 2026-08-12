(function () {
  "use strict";

  var BRIDGE_SOURCE = "onlyoffice-bridge";
  var BRIDGE_MESSAGE = {
    EDITOR_COMMAND: "editor:command",
    EDITOR_RESPONSE: "editor:response",
    EDITOR_EVENT: "editor:event",
    EDITOR_SET_READONLY: "editor:set-readonly",
  };
  var EDITOR_COMMAND = {
    EDITOR_SUBSCRIBE: "editor:subscribe",
    INTERFACE_SET_THEME: "interface:set-theme",
    DOCUMENT_PRINT_PDF: "document:print-pdf",
    DOCUMENT_RENAME: "document:rename",
    COMMENT_ADD: "comment:add",
    COMMENT_UPDATE: "comment:update",
    COMMENT_REMOVE: "comment:remove",
    COMMENT_GO_TO: "comment:go-to",
    COMMENT_LIST: "comment:list",
    COMMENT_SUBSCRIBE: "comment:subscribe",
    REVISION_ADD_DEMO: "revision:add-demo",
    REVISION_LIST: "revision:list",
    REVISION_SET_TRACK: "revision:set-track",
    REVISION_IS_TRACK: "revision:is-track",
    REVISION_HAVE_CHANGES: "revision:have-changes",
    REVISION_PREPARE_REVIEW: "revision:prepare-review",
    REVISION_NEXT: "revision:next",
    REVISION_PREV: "revision:prev",
    REVISION_GO_TO: "revision:go-to",
    REVISION_ACCEPT: "revision:accept",
    REVISION_REJECT: "revision:reject",
    REVISION_ACCEPT_ALL: "revision:accept-all",
    REVISION_REJECT_ALL: "revision:reject-all",
    REVISION_ACCEPT_SELECTION: "revision:accept-selection",
    REVISION_REJECT_SELECTION: "revision:reject-selection",
    REVISION_SUBSCRIBE: "revision:subscribe",
  };
  var EDITOR_EVENT = {
    ADD_COMMENT: "asc_onAddComment",
    CHANGE_COMMENT: "asc_onChangeCommentData",
    REMOVE_COMMENT: "asc_onRemoveComment",
    SHOW_REVISIONS_CHANGE: "asc_onShowRevisionsChange",
    TRACK_REVISIONS_CHANGE: "asc_onOnTrackRevisionsChange",
    DOCUMENT_MODIFIED_CHANGED: "asc_onDocumentModifiedChanged",
  };
  var ASC_RESTRICTION_NONE = 0;
  var ASC_RESTRICTION_VIEW = 128;
  var requestSeq = 0;
  var pendingHttp = Object.create(null);
  var commentCallbacksRegistered = false;
  var editorCallbacksRegistered = Object.create(null);

  function getFrameEditorId() {
    try {
      return (
        new URLSearchParams(window.location.search).get("frameEditorId") || ""
      );
    } catch (error) {
      return "";
    }
  }

  function isSameOriginParent() {
    try {
      if (!window.parent || window.parent === window) {
        return true;
      }
      return window.parent.location.origin === window.location.origin;
    } catch (error) {
      return false;
    }
  }

  var frameEditorId = getFrameEditorId();
  var parentOrigin = "";
  try {
    parentOrigin = new URL(window.parentOrigin).origin;
  } catch (error) {
    parentOrigin = "";
  }
  if (!frameEditorId || !parentOrigin) {
    return;
  }
  var sameOriginParent = isSameOriginParent();

  function createBridgeInstanceId() {
    var randomPart = "";
    try {
      var values = new Uint32Array(2);
      window.crypto.getRandomValues(values);
      randomPart = values[0].toString(36) + values[1].toString(36);
    } catch (error) {
      randomPart = Math.random().toString(36).slice(2);
    }
    return frameEditorId + "-" + Date.now().toString(36) + "-" + randomPart;
  }

  var bridgeInstanceId = createBridgeInstanceId();

  function post(message) {
    if (!window.parent || window.parent === window) {
      return;
    }
    window.parent.postMessage(
      Object.assign({}, message, {
        source: BRIDGE_SOURCE,
        frameEditorId: frameEditorId,
        bridgeInstanceId: bridgeInstanceId,
      }),
      parentOrigin,
    );
  }

  var handshakeReady = false;
  var queuedSocketEmits = [];
  var queuedSocketEvents = [];

  var resourceLoadingState = {
    loading: true,
    resourceStatus: "checking",
    resourceDownload: false,
    transferredBytes: 0,
    resourceCount: 0,
  };
  var resourceLoadingFinalTimer = 0;
  var observedResourceNames = Object.create(null);
  var cachedResourceCount = 0;
  var downloadedResourceCount = 0;

  function postResourceLoadingState() {
    if (!window.parent || window.parent === window) return;
    window.parent.postMessage({
      source: "onlyoffice-resource-observer",
      type: BRIDGE_MESSAGE.EDITOR_EVENT,
      frameEditorId: frameEditorId,
      state: Object.assign({}, resourceLoadingState),
    }, parentOrigin);
  }

  function isHostedStaticResource(entry) {
    try {
      var url = new URL(entry.name, document.baseURI);
      var base = new URL(document.baseURI);
      return (
        url.pathname.indexOf("/onlyoffice/runtime/") !== -1 ||
        (url.origin === base.origin && url.origin !== window.location.origin)
      );
    } catch (error) {
      return false;
    }
  }

  function finishResourceLoadingAfterQuietPeriod() {
    if (resourceLoadingFinalTimer) {
      window.clearTimeout(resourceLoadingFinalTimer);
    }
    resourceLoadingFinalTimer = window.setTimeout(function () {
      resourceLoadingFinalTimer = 0;
      resourceLoadingState.loading = false;
      resourceLoadingState.resourceDownload = false;
      resourceLoadingState.resourceStatus = downloadedResourceCount
        ? "downloaded"
        : cachedResourceCount
          ? "cache-hit"
          : "not-observed";
      postResourceLoadingState();
    }, 300);
  }

  function observeStaticResource(entry) {
    if (!isHostedStaticResource(entry) || observedResourceNames[entry.name]) {
      return;
    }
    observedResourceNames[entry.name] = true;
    var deliveryType = typeof entry.deliveryType === "string" ? entry.deliveryType : "";
    var encodedBodySize = Number(entry.encodedBodySize) || 0;
    var transferSize = Number(entry.transferSize) || 0;
    var isCacheHit =
      deliveryType === "cache" ||
      ((transferSize === 0 || encodedBodySize === 0) &&
        Number(entry.decodedBodySize) > 0);
    var downloaded = !isCacheHit && encodedBodySize > 0 && transferSize > 0;
    if (downloaded) {
      var wasDownloading = resourceLoadingState.resourceDownload;
      downloadedResourceCount += 1;
      resourceLoadingState.transferredBytes += encodedBodySize;
      resourceLoadingState.resourceCount =
        downloadedResourceCount + cachedResourceCount;
      resourceLoadingState.loading = true;
      resourceLoadingState.resourceStatus = "downloading";
      resourceLoadingState.resourceDownload = true;
      if (!wasDownloading) postResourceLoadingState();
    } else if (isCacheHit) {
      cachedResourceCount += 1;
      resourceLoadingState.resourceCount =
        downloadedResourceCount + cachedResourceCount;
    }
    finishResourceLoadingAfterQuietPeriod();
  }

  try {
    performance.getEntriesByType("resource").forEach(observeStaticResource);
    new PerformanceObserver(function (list) {
      list.getEntries().forEach(observeStaticResource);
    }).observe({ type: "resource", buffered: true });
    window.addEventListener("load", finishResourceLoadingAfterQuietPeriod, {
      once: true,
    });
  } catch (error) {
    resourceLoadingState.loading = false;
    resourceLoadingState.resourceStatus = "not-observed";
    finishResourceLoadingAfterQuietPeriod();
  }

  if (sameOriginParent) {
    return;
  }

  function flushSocketEmits() {
    if (!handshakeReady) {
      return;
    }
    while (queuedSocketEmits.length) {
      post(queuedSocketEmits.shift());
    }
  }

  function sendHello() {
    post({ type: "hello" });
  }

  var helloTimer = window.setInterval(sendHello, 50);
  sendHello();

  function createSocket(options) {
    var handlers = Object.create(null);
    var onceHandlers = Object.create(null);
    var managerHandlers = Object.create(null);
    var queuedMessageEvents = [];
    var manager = {
      opts: Object.assign({}, options || {}),
      setOpenToken: function () {
        return manager;
      },
      setSessionToken: function () {
        return manager;
      },
      on: function (event, handler) {
        if (typeof handler === "function") {
          (managerHandlers[event] || (managerHandlers[event] = [])).push(
            handler,
          );
        }
        return manager;
      },
      off: function (event, handler) {
        if (!event) {
          managerHandlers = Object.create(null);
          return manager;
        }
        if (!handler) {
          delete managerHandlers[event];
          return manager;
        }
        managerHandlers[event] = (managerHandlers[event] || []).filter(
          function (item) {
            return item !== handler;
          },
        );
        return manager;
      },
      reconnectionAttempts: function () {
        return manager;
      },
      reconnectionDelay: function () {
        return manager;
      },
      reconnectionDelayMax: function () {
        return manager;
      },
      timeout: function () {
        return manager;
      },
      transports: function () {
        return manager;
      },
      upgrade: function () {
        return manager;
      },
      upgradeTransport: function () {
        return manager;
      },
      upgradeTimeout: function () {
        return manager;
      },
    };
    var socket = {
      active: true,
      connected: true,
      disconnected: false,
      recovered: false,
      id: frameEditorId,
      io: manager,
      nsp: "/",
      on: function (event, handler) {
        if (typeof handler !== "function") {
          return socket;
        }
        (handlers[event] || (handlers[event] = [])).push(handler);
        if (event === "message" && queuedMessageEvents.length) {
          var queued = queuedMessageEvents.splice(0);
          queued.forEach(function (args) {
            dispatchSocketEvent("message", args);
          });
        }
        return socket;
      },
      once: function (event, handler) {
        if (typeof handler !== "function") {
          return socket;
        }
        (onceHandlers[event] || (onceHandlers[event] = [])).push(handler);
        return socket;
      },
      off: function (event, handler) {
        if (!event) {
          handlers = Object.create(null);
          onceHandlers = Object.create(null);
          return socket;
        }
        if (!handler) {
          delete handlers[event];
          delete onceHandlers[event];
          return socket;
        }
        handlers[event] = (handlers[event] || []).filter(function (item) {
          return item !== handler;
        });
        onceHandlers[event] = (onceHandlers[event] || []).filter(
          function (item) {
            return item !== handler;
          },
        );
        return socket;
      },
      removeListener: function (event, handler) {
        return socket.off(event, handler);
      },
      removeAllListeners: function (event) {
        return socket.off(event);
      },
      send: function () {
        var args = ["message"].concat(Array.prototype.slice.call(arguments));
        return socket.emit.apply(socket, args);
      },
      emit: function (event) {
        var args = Array.prototype.slice.call(arguments, 1);
        var message = { type: "socket:emit", event: event, args: args };
        if (handshakeReady) {
          post(message);
        } else {
          queuedSocketEmits.push(message);
        }
        return socket;
      },
      open: function () {
        return socket.connect();
      },
      connect: function () {
        socket.connected = true;
        socket.disconnected = false;
        sendHello();
        dispatchSocketEvent("connect", []);
        return socket;
      },
      disconnect: function () {
        socket.connected = false;
        socket.disconnected = true;
        post({ type: "socket:emit", event: "disconnect", args: [] });
        dispatchSocketEvent("disconnect", []);
        return socket;
      },
      close: function () {
        return socket.disconnect();
      },
      compress: function () {
        return socket;
      },
    };

    function dispatchSocketEvent(event, args) {
      if (event === "message" && !(handlers.message || onceHandlers.message)) {
        queuedMessageEvents.push(args);
        return;
      }

      (handlers[event] || []).slice().forEach(function (handler) {
        handler.apply(socket, args);
      });

      var once = onceHandlers[event];
      if (once && once.length) {
        delete onceHandlers[event];
        once.slice().forEach(function (handler) {
          handler.apply(socket, args);
        });
      }
    }

    socket.__onlyofficeDispatch = dispatchSocketEvent;
    window.setTimeout(function () {
      dispatchSocketEvent("connect", []);
    }, 0);

    return socket;
  }

  function registerScopedSocket() {
    window.__ONLYOFFICE_SCOPED_IO__ = window.__ONLYOFFICE_SCOPED_IO__ || {};
    window.__ONLYOFFICE_SCOPED_IO__[frameEditorId] = function (url, options) {
      return createSocket(options);
    };
  }

  registerScopedSocket();

  function arrayBufferToBase64(buffer) {
    var bytes = new Uint8Array(buffer);
    var binary = "";
    for (var i = 0; i < bytes.length; i += 1) {
      binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
  }

  function base64ToArrayBuffer(base64) {
    var binary = atob(base64 || "");
    var bytes = new Uint8Array(binary.length);
    for (var i = 0; i < binary.length; i += 1) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes.buffer;
  }

  function shouldProxyUrl(url) {
    if (
      !url ||
      typeof url !== "string" ||
      url.indexOf("blob:") === 0 ||
      url.indexOf("data:") === 0
    ) {
      return false;
    }
    try {
      var parsed = new URL(url, window.location.href);
      var isCrossOriginPluginConfig =
        parsed.origin !== window.location.origin &&
        parsed.pathname.slice(-"/config.json".length) === "/config.json";
      return (
        parsed.pathname.indexOf("/cache/files/") !== -1 ||
        parsed.pathname.indexOf("/downloadas/") !== -1 ||
        parsed.pathname.indexOf("/upload/") !== -1 ||
        isCrossOriginPluginConfig
      );
    } catch (error) {
      return false;
    }
  }

  function isForbiddenRequestHeader(name) {
    var lowerName = String(name || "").toLowerCase();
    return (
      lowerName === "accept-charset" ||
      lowerName === "accept-encoding" ||
      lowerName === "access-control-request-headers" ||
      lowerName === "access-control-request-method" ||
      lowerName === "connection" ||
      lowerName === "content-length" ||
      lowerName === "cookie" ||
      lowerName === "cookie2" ||
      lowerName === "date" ||
      lowerName === "dnt" ||
      lowerName === "expect" ||
      lowerName === "host" ||
      lowerName === "keep-alive" ||
      lowerName === "origin" ||
      lowerName === "referer" ||
      lowerName === "te" ||
      lowerName === "trailer" ||
      lowerName === "transfer-encoding" ||
      lowerName === "upgrade" ||
      lowerName === "via" ||
      lowerName.indexOf("proxy-") === 0 ||
      lowerName.indexOf("sec-") === 0
    );
  }

  function setHeader(target, name, value) {
    if (!name || isForbiddenRequestHeader(name)) {
      return;
    }
    target[name] = value;
  }

  function headersToObject(headers) {
    var result = {};
    if (!headers) {
      return result;
    }
    if (headers instanceof Headers) {
      headers.forEach(function (value, key) {
        setHeader(result, key, value);
      });
      return result;
    }
    if (Array.isArray(headers)) {
      headers.forEach(function (entry) {
        setHeader(result, entry[0], entry[1]);
      });
      return result;
    }
    Object.keys(headers).forEach(function (key) {
      setHeader(result, key, headers[key]);
    });
    return result;
  }

  function normalizeBody(body) {
    if (body == null || typeof body === "string") {
      return Promise.resolve({ body: body == null ? null : body });
    }
    if (body instanceof ArrayBuffer) {
      return Promise.resolve({
        body: arrayBufferToBase64(body),
        bodyEncoding: "base64",
      });
    }
    if (ArrayBuffer.isView(body)) {
      return Promise.resolve({
        body: arrayBufferToBase64(
          body.buffer.slice(body.byteOffset, body.byteOffset + body.byteLength),
        ),
        bodyEncoding: "base64",
      });
    }
    if (body instanceof Blob) {
      return body.arrayBuffer().then(function (buffer) {
        return { body: arrayBufferToBase64(buffer), bodyEncoding: "base64" };
      });
    }
    return Promise.resolve({ body: String(body) });
  }

  function requestViaParent(method, url, headers, body, responseType) {
    return normalizeBody(body).then(function (normalized) {
      return new Promise(function (resolve) {
        var requestId = "bridge-" + Date.now() + "-" + requestSeq++;
        pendingHttp[requestId] = resolve;
        post({
          type: "http",
          requestId: requestId,
          method: method,
          url: new URL(url, window.location.href).href,
          headers: headersToObject(headers),
          body: normalized.body,
          bodyEncoding: normalized.bodyEncoding,
          responseType: responseType,
        });
      });
    });
  }

  function installFetchProxy() {
    var nativeFetch = window.fetch && window.fetch.bind(window);
    if (!nativeFetch) {
      return;
    }

    window.fetch = function bridgeFetch(input, init) {
      var request = input instanceof Request ? input : null;
      var url = request ? request.url : String(input);
      if (!shouldProxyUrl(url)) {
        return nativeFetch(input, init);
      }

      var method =
        (init && init.method) || (request && request.method) || "GET";
      var headers = (init && init.headers) || (request && request.headers);
      var body = init && "body" in init ? init.body : null;
      return requestViaParent(method, url, headers, body, "arraybuffer").then(
        function (message) {
          var responseBody = base64ToArrayBuffer(message.responseBody);
          return new Response(responseBody, {
            status: message.status || 200,
            headers: message.responseHeaders || {},
          });
        },
      );
    };
  }

  function installXhrProxy() {
    var NativeXHR = window.XMLHttpRequest;
    if (!NativeXHR) {
      return;
    }

    window.XMLHttpRequest = function BridgeXMLHttpRequest() {
      var nativeXhr = new NativeXHR();
      var method = "GET";
      var url = "";
      var async = true;
      var requestHeaders = {};
      var responseHeaders = {};
      var listeners = {};
      var xhr = this;
      var responseType = "";

      this.readyState = 0;
      this.status = 0;
      this.statusText = "";
      this.response = null;
      this.responseText = "";
      this.onreadystatechange = null;
      this.onload = null;
      this.onerror = null;
      this.upload = nativeXhr.upload;

      Object.defineProperty(this, "responseType", {
        get: function () {
          return responseType;
        },
        set: function (value) {
          responseType = value || "";
          try {
            nativeXhr.responseType = responseType;
          } catch (error) {
            /* Native XHR can reject responseType changes after send. */
          }
        },
      });

      ["timeout", "withCredentials"].forEach(function (property) {
        Object.defineProperty(xhr, property, {
          get: function () {
            return nativeXhr[property];
          },
          set: function (value) {
            nativeXhr[property] = value;
          },
        });
      });

      function dispatch(eventName) {
        if (
          eventName === "readystatechange" &&
          typeof xhr.onreadystatechange === "function"
        ) {
          xhr.onreadystatechange.call(xhr);
        }
        if (eventName === "load" && typeof xhr.onload === "function") {
          xhr.onload.call(xhr);
        }
        if (eventName === "error" && typeof xhr.onerror === "function") {
          xhr.onerror.call(xhr);
        }
        (listeners[eventName] || []).slice().forEach(function (handler) {
          handler.call(xhr, {
            type: eventName,
            target: xhr,
            currentTarget: xhr,
          });
        });
      }

      function setReadyState(value) {
        xhr.readyState = value;
        dispatch("readystatechange");
      }

      this.open = function (nextMethod, nextUrl, nextAsync) {
        method = nextMethod || "GET";
        url = String(nextUrl || "");
        async = nextAsync !== false;
        if (!shouldProxyUrl(url)) {
          nativeXhr.open.apply(nativeXhr, arguments);
          return;
        }
        setReadyState(1);
      };

      this.send = function (body) {
        if (!shouldProxyUrl(url)) {
          copyNativeEvents(nativeXhr, xhr, dispatch);
          nativeXhr.send(body);
          return;
        }
        if (!async) {
          throw new Error("OnlyOffice bridge does not support synchronous XHR");
        }
        requestViaParent(
          method,
          url,
          requestHeaders,
          body,
          xhr.responseType || "",
        ).then(function (message) {
          responseHeaders = message.responseHeaders || {};
          xhr.status = message.status || 0;
          xhr.statusText =
            xhr.status >= 200 && xhr.status < 400 ? "OK" : "Error";
          setReadyState(2);
          setReadyState(3);
          if (
            xhr.responseType === "arraybuffer" ||
            xhr.responseType === "blob"
          ) {
            var buffer = base64ToArrayBuffer(message.responseBody);
            xhr.response =
              xhr.responseType === "blob" ? new Blob([buffer]) : buffer;
          } else {
            xhr.responseText = message.responseBody || "";
            xhr.response = xhr.responseText;
          }
          setReadyState(4);
          dispatch("load");
          dispatch("loadend");
        });
      };

      this.abort = function () {
        if (!shouldProxyUrl(url)) {
          nativeXhr.abort();
        }
      };

      this.setRequestHeader = function (name, value) {
        if (isForbiddenRequestHeader(name)) {
          return;
        }
        if (!shouldProxyUrl(url)) {
          nativeXhr.setRequestHeader(name, value);
          return;
        }
        setHeader(requestHeaders, name, value);
      };

      this.overrideMimeType = function (mimeType) {
        if (nativeXhr.overrideMimeType) {
          nativeXhr.overrideMimeType(mimeType);
        }
      };

      this.getResponseHeader = function (name) {
        if (!shouldProxyUrl(url)) {
          return nativeXhr.getResponseHeader(name);
        }
        return (
          responseHeaders[String(name).toLowerCase()] ||
          responseHeaders[name] ||
          null
        );
      };

      this.getAllResponseHeaders = function () {
        if (!shouldProxyUrl(url)) {
          return nativeXhr.getAllResponseHeaders();
        }
        return Object.keys(responseHeaders)
          .map(function (key) {
            return key + ": " + responseHeaders[key];
          })
          .join("\r\n");
      };

      this.addEventListener = function (eventName, handler) {
        (listeners[eventName] || (listeners[eventName] = [])).push(handler);
      };

      this.removeEventListener = function (eventName, handler) {
        listeners[eventName] = (listeners[eventName] || []).filter(
          function (item) {
            return item !== handler;
          },
        );
      };
    };
  }

  function copyNativeEvents(nativeXhr, bridgeXhr, dispatch) {
    nativeXhr.onreadystatechange = function () {
      bridgeXhr.readyState = nativeXhr.readyState;
      bridgeXhr.status = nativeXhr.status;
      bridgeXhr.statusText = nativeXhr.statusText;
      bridgeXhr.response = nativeXhr.response;
      try {
        bridgeXhr.responseText = nativeXhr.responseText;
      } catch (error) {
        bridgeXhr.responseText = "";
      }
      dispatch("readystatechange");
    };
    nativeXhr.onload = function () {
      dispatch("load");
    };
    nativeXhr.onerror = function () {
      dispatch("error");
    };
  }

  function extractDownloadFileName(url) {
    try {
      var parsed = new URL(url, window.location.href);
      return (
        parsed.searchParams.get("filename") ||
        decodeURIComponent(parsed.pathname.split("/").pop() || "download")
      );
    } catch (error) {
      return "download";
    }
  }

  function parseContentDispositionFileName(header) {
    if (!header) {
      return "";
    }
    var utf8 = /filename\*=UTF-8''([^;]+)/i.exec(header);
    if (utf8 && utf8[1]) {
      try {
        return decodeURIComponent(utf8[1]);
      } catch (error) {
        return utf8[1];
      }
    }
    var ascii = /filename="([^"]+)"/i.exec(header);
    return ascii && ascii[1] ? ascii[1] : "";
  }

  function triggerBlobDownload(blob, fileName) {
    var objectUrl = URL.createObjectURL(blob);
    var link = document.createElement("a");
    link.href = objectUrl;
    link.download = fileName || "download";
    link.style.display = "none";
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(objectUrl);
  }

  function installNamedDownloadPatch(retries) {
    if (window.__ONLYOFFICE_GETFILE_PATCHED__) {
      return;
    }
    if (!window.AscCommon || !window.AscCommon.getFile) {
      if (retries > 0) {
        window.setTimeout(function () {
          installNamedDownloadPatch(retries - 1);
        }, 50);
      }
      return;
    }

    var nativeGetFile = window.AscCommon.getFile.bind(window.AscCommon);
    window.AscCommon.getFile = function (url) {
      if (typeof url !== "string" || url.indexOf("/cache/files/") === -1) {
        nativeGetFile(url);
        return;
      }
      var fallbackName = extractDownloadFileName(url);
      window
        .fetch(url)
        .then(function (response) {
          return response.blob().then(function (blob) {
            return {
              blob: blob,
              fileName:
                parseContentDispositionFileName(
                  response.headers.get("Content-Disposition"),
                ) || fallbackName,
            };
          });
        })
        .then(function (result) {
          triggerBlobDownload(result.blob, result.fileName);
        })
        .catch(function () {
          nativeGetFile(url);
        });
    };
    window.__ONLYOFFICE_GETFILE_PATCHED__ = true;
  }

  function installPrintFramePatch() {
    if (window.__ONLYOFFICE_PRINT_FRAME_PATCHED__) {
      return;
    }

    var iframePrototype = window.HTMLIFrameElement.prototype;
    var srcDescriptor = Object.getOwnPropertyDescriptor(
      iframePrototype,
      "src",
    );
    if (!srcDescriptor || !srcDescriptor.get || !srcDescriptor.set) {
      return;
    }

    var nativeGetter = srcDescriptor.get;
    var nativeSetter = srcDescriptor.set;
    Object.defineProperty(iframePrototype, "src", {
      configurable: srcDescriptor.configurable,
      enumerable: srcDescriptor.enumerable,
      get: nativeGetter,
      set: function (value) {
        var target = this;
        if (
          target.id !== "id-print-frame" ||
          typeof value !== "string" ||
          value.indexOf("/cache/files/") === -1
        ) {
          nativeSetter.call(target, value);
          return;
        }

        var nativeOnLoad = target.onload;
        var printNavigationPending = true;
        if (typeof nativeOnLoad === "function") {
          target.onload = function (event) {
            if (printNavigationPending) {
              return;
            }
            target.onload = nativeOnLoad;
            return nativeOnLoad.call(target, event);
          };
        }

        window
          .fetch(value)
          .then(function (response) {
            if (!response.ok) {
              throw new Error("Print PDF fetch failed: " + response.status);
            }
            return response.blob();
          })
          .then(function (blob) {
            var objectUrl = URL.createObjectURL(blob);
            printNavigationPending = false;
            nativeSetter.call(target, objectUrl);
            window.setTimeout(function () {
              URL.revokeObjectURL(objectUrl);
            }, 60000);
          })
          .catch(function (error) {
            console.warn("[OnlyOffice] print PDF fetch failed:", error);
            printNavigationPending = false;
            nativeSetter.call(target, value);
          });
      },
    });

    window.__ONLYOFFICE_PRINT_FRAME_PATCHED__ = true;
  }

  function syncControllerReadOnly(readOnly) {
    try {
      var common = window.Common;
      var controller = common && common.Controllers && common.Controllers.Main;
      if (controller && controller.mode) {
        controller.mode.isEdit = !readOnly;
        controller.mode.canEdit = !readOnly;
      }
      if (
        controller &&
        controller.editorConfig &&
        controller.editorConfig.mode
      ) {
        controller.editorConfig.mode.isEdit = !readOnly;
        controller.editorConfig.mode.canEdit = !readOnly;
      }
      if (
        common &&
        common.NotificationCenter &&
        common.NotificationCenter.trigger
      ) {
        common.NotificationCenter.trigger("editing:disable", readOnly);
      }
    } catch (error) {
      /* best effort */
    }
  }

  function setReadOnly(readOnly) {
    try {
      var editor = (window.Asc && window.Asc.editor) || window.editor;
      if (!editor) {
        return false;
      }
      if (typeof editor.asc_setRestriction === "function") {
        if (readOnly) {
          editor.asc_setRestriction(ASC_RESTRICTION_VIEW);
        } else {
          if (typeof editor.asc_removeRestriction === "function") {
            editor.asc_removeRestriction(ASC_RESTRICTION_VIEW);
          }
          editor.asc_setRestriction(ASC_RESTRICTION_NONE);
        }
      }
      if (typeof editor.asc_setCanSendChanges === "function") {
        editor.asc_setCanSendChanges(!readOnly);
      }
      syncControllerReadOnly(readOnly);
      return true;
    } catch (error) {
      return false;
    }
  }

  function scheduleReadOnly(readOnly, retries) {
    if (setReadOnly(readOnly) || retries <= 0) {
      return;
    }
    window.setTimeout(function () {
      scheduleReadOnly(readOnly, retries - 1);
    }, 50);
  }

  function getEditorApi() {
    return (window.Asc && window.Asc.editor) || window.editor;
  }

  var syncingRevisionStack = false;
  var revisionCallbacksRegistered = false;

  function callGetter(target, name) {
    return target && typeof target[name] === "function"
      ? target[name]()
      : undefined;
  }

  function commentDataToPlain(data) {
    if (!data || typeof data !== "object") {
      return {};
    }

    var source = {};
    Object.keys(data).forEach(function (key) {
      var value = data[key];
      if (typeof value !== "function") {
        source[key] = value;
      }
    });

    var solved =
      source.Solved !== undefined
        ? source.Solved
        : source.solved !== undefined
          ? source.solved
          : callGetter(data, "asc_getSolved");

    return Object.assign({}, source, {
      Text:
        source.Text !== undefined
          ? source.Text
          : callGetter(data, "asc_getText"),
      UserName:
        source.UserName !== undefined
          ? source.UserName
          : callGetter(data, "asc_getUserName"),
      Time:
        source.Time !== undefined
          ? source.Time
          : callGetter(data, "asc_getTime"),
      QuoteText:
        source.QuoteText !== undefined
          ? source.QuoteText
          : callGetter(data, "asc_getQuoteText"),
      Solved: !!solved,
      UserData:
        source.UserData !== undefined
          ? source.UserData
          : callGetter(data, "asc_getUserData"),
      Replies: Array.isArray(source.Replies) ? source.Replies : [],
    });
  }

  function normalizeCommentItems(raw) {
    if (!Array.isArray(raw)) {
      return [];
    }
    return raw.map(function (item, index) {
      var source = item && typeof item === "object" ? item : {};
      return {
        Id: String(source.Id || source.id || "comment-" + index),
        Data: commentDataToPlain(source.Data || source.data || source),
      };
    });
  }

  function getAllComments(api) {
    return normalizeCommentItems(
      api.pluginMethod_GetAllComments ? api.pluginMethod_GetAllComments() : [],
    );
  }

  function registerCommentCallbacks(api) {
    if (commentCallbacksRegistered || !api.asc_registerCallback) {
      return;
    }

    api.asc_registerCallback(EDITOR_EVENT.ADD_COMMENT, function (id, data) {
      post({
        type: BRIDGE_MESSAGE.EDITOR_EVENT,
        event: EDITOR_EVENT.ADD_COMMENT,
        args: [String(id), commentDataToPlain(data)],
      });
    });
    api.asc_registerCallback(EDITOR_EVENT.CHANGE_COMMENT, function (id, data) {
      post({
        type: BRIDGE_MESSAGE.EDITOR_EVENT,
        event: EDITOR_EVENT.CHANGE_COMMENT,
        args: [String(id), commentDataToPlain(data)],
      });
    });
    api.asc_registerCallback(EDITOR_EVENT.REMOVE_COMMENT, function (id) {
      post({
        type: BRIDGE_MESSAGE.EDITOR_EVENT,
        event: EDITOR_EVENT.REMOVE_COMMENT,
        args: [String(id)],
      });
    });
    commentCallbacksRegistered = true;
  }

  function getLogicDocument(api) {
    if (api && typeof api.te === "function") {
      return api.te() || null;
    }
    if (api && typeof api.Ge === "function") {
      return api.Ge() || null;
    }
    if (api && typeof api.Gg === "function") {
      return api.Gg() || null;
    }
    return null;
  }

  function getAscRevisionEnums() {
    var Asc = window.Asc || {};
    var changeType = Asc.c_oAscRevisionsChangeType || {};
    var moveType = Asc.c_oAscRevisionsMove || {};
    return {
      TextAdd: changeType.TextAdd,
      TextRem: changeType.TextRem,
      ParaAdd: changeType.ParaAdd,
      ParaRem: changeType.ParaRem,
      TextPr: changeType.TextPr,
      ParaPr: changeType.ParaPr,
      TablePr: changeType.TablePr,
      RowsAdd: changeType.RowsAdd,
      RowsRem: changeType.RowsRem,
      TableRowPr: changeType.TableRowPr,
      MoveMark: changeType.MoveMark,
      Unknown: changeType.Unknown,
      MoveTo: moveType.MoveTo,
      MoveFrom: moveType.MoveFrom,
      NoMove: moveType.NoMove,
    };
  }

  function revisionTypeName(type, enums) {
    var entries = [
      ["TextAdd", enums.TextAdd],
      ["TextRem", enums.TextRem],
      ["ParaAdd", enums.ParaAdd],
      ["ParaRem", enums.ParaRem],
      ["TextPr", enums.TextPr],
      ["ParaPr", enums.ParaPr],
      ["TablePr", enums.TablePr],
      ["RowsAdd", enums.RowsAdd],
      ["RowsRem", enums.RowsRem],
      ["TableRowPr", enums.TableRowPr],
      ["MoveMark", enums.MoveMark],
      ["Unknown", enums.Unknown],
    ];
    for (var i = 0; i < entries.length; i += 1) {
      if (entries[i][1] !== undefined && entries[i][1] === type) {
        return entries[i][0];
      }
    }
    return String(type);
  }

  function revisionType(raw) {
    if (raw && typeof raw.get_Type === "function") {
      return raw.get_Type();
    }
    if (raw && typeof raw.Gc === "function") {
      return raw.Gc();
    }
    return undefined;
  }

  function revisionValue(raw) {
    var value;
    if (raw && typeof raw.get_Value === "function") {
      value = raw.get_Value();
      return value == null ? undefined : String(value);
    }
    if (raw && typeof raw.Ym === "function") {
      value = raw.Ym();
      return value == null ? undefined : String(value);
    }
    return undefined;
  }

  function revisionStart(raw) {
    var pos;
    if (raw && typeof raw.get_StartPos === "function") {
      pos = raw.get_StartPos();
      if (typeof pos === "number") return pos;
    }
    if (raw && typeof raw.SU === "function") {
      pos = raw.SU();
      if (typeof pos === "number") return pos;
    }
    if (raw && typeof raw.wa === "number") return raw.wa;
    if (raw && typeof raw.j9 === "function") return raw.j9();
    if (raw && typeof raw.sa === "number") return raw.sa;
    return undefined;
  }

  function revisionEnd(raw) {
    var pos;
    if (raw && typeof raw.get_EndPos === "function") {
      pos = raw.get_EndPos();
      if (typeof pos === "number") return pos;
    }
    if (raw && typeof raw.h5 === "function") {
      pos = raw.h5();
      if (typeof pos === "number") return pos;
    }
    if (raw && typeof raw.xa === "number") return raw.xa;
    if (raw && typeof raw.yza === "function") return raw.yza();
    if (raw && typeof raw.ra === "number") return raw.ra;
    return undefined;
  }

  function revisionIdentity(raw) {
    if (raw && typeof raw.ed === "function") {
      return String(raw.ed());
    }
    return null;
  }

  function isSameRevision(a, b) {
    if (a === b) return true;
    var idA = revisionIdentity(a);
    var idB = revisionIdentity(b);
    return idA != null && idA === idB;
  }

  function revisionSortKey(raw) {
    var y = raw && typeof raw.$Zd === "function" ? raw.$Zd() : 0;
    var x = raw && typeof raw.Vtb === "function" ? raw.Vtb() : 0;
    return y * 1e9 + x * 1e3 + (revisionStart(raw) || 0);
  }

  function revisionChangeToPlain(raw, enums) {
    var type = revisionType(raw);
    var lockUserId = raw && raw.get_LockUserId ? raw.get_LockUserId() : null;
    return {
      Type: type,
      TypeName: type !== undefined ? revisionTypeName(type, enums) : undefined,
      UserId:
        (raw && typeof raw.get_UserId === "function"
          ? raw.get_UserId()
          : undefined) ||
        (raw && typeof raw.Qy === "function" ? raw.Qy() : undefined),
      UserName:
        (raw && typeof raw.get_UserName === "function"
          ? raw.get_UserName()
          : undefined) ||
        (raw && typeof raw.wV === "function" ? raw.wV() : undefined),
      DateTime:
        (raw && typeof raw.get_DateTime === "function"
          ? raw.get_DateTime()
          : undefined) ||
        (raw && typeof raw.nX === "function" ? raw.nX() : undefined),
      Value: revisionValue(raw),
      X:
        (raw && typeof raw.get_X === "function" ? raw.get_X() : undefined) ||
        (raw && typeof raw.$Zd === "function" ? raw.$Zd() : undefined),
      Y:
        (raw && typeof raw.get_Y === "function" ? raw.get_Y() : undefined) ||
        (raw && typeof raw.Vtb === "function" ? raw.Vtb() : undefined),
      MoveType:
        (raw && typeof raw.get_MoveType === "function"
          ? raw.get_MoveType()
          : undefined) ||
        (raw && typeof raw.hZd === "function" ? raw.hZd() : undefined),
      MoveId: raw && raw.get_MoveId ? raw.get_MoveId() : undefined,
      Locked: lockUserId != null && lockUserId !== "",
      LockUserId: lockUserId || undefined,
    };
  }

  function ensureRevisionsIndexed(api) {
    var doc = getLogicDocument(api);
    if (doc && doc.Wq && doc.Wq.z9a) doc.Wq.z9a();
    if (doc && doc.Um && doc.Um.MQc) doc.Um.MQc();
  }

  function collectRevisionRawsFromWq(doc) {
    var wq = doc && doc.Wq;
    var qt = wq && wq.wih && wq.wih();
    if (!qt || typeof qt !== "object" || Object.keys(qt).length === 0) {
      return [];
    }

    var raws = [];
    var orderedIds = [];
    (wq.Tpa || []).forEach(function (el) {
      var id =
        typeof el.Yb === "function"
          ? String(el.Yb())
          : typeof el.ed === "function"
            ? String(el.ed())
            : "";
      if (id && qt[id] && orderedIds.indexOf(id) === -1) {
        orderedIds.push(id);
      }
    });

    Object.keys(qt)
      .filter(function (id) {
        return orderedIds.indexOf(id) === -1;
      })
      .sort(function (a, b) {
        var ga = qt[a] && qt[a][0];
        var gb = qt[b] && qt[b][0];
        return ga && gb ? revisionSortKey(ga) - revisionSortKey(gb) : 0;
      })
      .forEach(function (id) {
        orderedIds.push(id);
      });

    (orderedIds.length ? orderedIds : Object.keys(qt)).forEach(function (id) {
      var group = qt[id];
      if (!Array.isArray(group)) return;
      raws.push.apply(
        raws,
        group.length > 1
          ? group.slice().sort(function (a, b) {
              return revisionSortKey(a) - revisionSortKey(b);
            })
          : group,
      );
    });

    return raws;
  }

  function readRevisionStack(api, allowSyncEnd, forceRefresh) {
    var existing;
    if (!forceRefresh) {
      existing =
        (api.asc_GetRevisionsChangesStack &&
          api.asc_GetRevisionsChangesStack()) ||
        [];
      if (existing.length > 0 || !allowSyncEnd || syncingRevisionStack) {
        return existing;
      }
    }

    if (!allowSyncEnd || syncingRevisionStack) {
      return (
        (api.asc_GetRevisionsChangesStack &&
          api.asc_GetRevisionsChangesStack()) ||
        []
      );
    }

    syncingRevisionStack = true;
    try {
      api.sync_BeginCatchRevisionsChanges &&
        api.sync_BeginCatchRevisionsChanges();
      var doc = getLogicDocument(api);
      if (doc && doc.Um && doc.Um.Dmf) doc.Um.Dmf();
      api.sync_EndCatchRevisionsChanges &&
        api.sync_EndCatchRevisionsChanges(false);
      return (
        (api.asc_GetRevisionsChangesStack &&
          api.asc_GetRevisionsChangesStack()) ||
        []
      );
    } finally {
      syncingRevisionStack = false;
    }
  }

  function mapRevisionItems(raws, idOf) {
    var enums = getAscRevisionEnums();
    return raws.map(function (raw, index) {
      return {
        Id: idOf(raw, index),
        Index: index,
        Data: revisionChangeToPlain(raw, enums),
        Raw: {},
      };
    });
  }

  function collectRevisionItems(api, options) {
    options = options || {};
    var allowSyncEnd = options.allowSyncEnd !== false;
    var forceRefreshStack = !!options.forceRefreshStack;
    ensureRevisionsIndexed(api);

    var revId = function (raw, index) {
      return "rev-" + (typeof raw.ed === "function" ? String(raw.ed()) : index);
    };
    var doc = getLogicDocument(api);
    var wqRaws = doc ? collectRevisionRawsFromWq(doc) : [];
    if (wqRaws.length > 0) return mapRevisionItems(wqRaws, revId);

    var um = doc && doc.Um;
    var yq = um && um.yif && um.yif();
    if (yq && Object.keys(yq).length > 0) {
      var ids = [];
      (um.T8 || []).forEach(function (el) {
        var id = typeof el.ed === "function" ? String(el.ed()) : "";
        if (id && yq[id]) ids.push(id);
      });
      Object.keys(yq)
        .filter(function (id) {
          return ids.indexOf(id) === -1;
        })
        .sort(function (a, b) {
          var ga = yq[a] && yq[a][0];
          var gb = yq[b] && yq[b][0];
          return ga && gb ? revisionSortKey(ga) - revisionSortKey(gb) : 0;
        })
        .forEach(function (id) {
          ids.push(id);
        });

      var yqRaws = [];
      (ids.length ? ids : Object.keys(yq)).forEach(function (id) {
        var group = yq[id];
        if (!Array.isArray(group)) return;
        yqRaws.push.apply(
          yqRaws,
          group.length > 1
            ? group.slice().sort(function (a, b) {
                return revisionSortKey(a) - revisionSortKey(b);
              })
            : group,
        );
      });
      if (yqRaws.length > 0) return mapRevisionItems(yqRaws, revId);
    }

    var report =
      api.asc_GetTrackRevisionsReportByAuthors &&
      api.asc_GetTrackRevisionsReportByAuthors();
    if (report && typeof report === "object") {
      var reportRaws = [];
      Object.keys(report).forEach(function (key) {
        if (Array.isArray(report[key]))
          reportRaws.push.apply(reportRaws, report[key]);
      });
      if (reportRaws.length > 0) return mapRevisionItems(reportRaws, revId);
    }

    return mapRevisionItems(
      readRevisionStack(api, allowSyncEnd, forceRefreshStack),
      function (_, index) {
        return "rev-stack-" + index;
      },
    );
  }

  function resolveRevisionShowChanges(stack, api) {
    if (Array.isArray(stack) && stack.length > 0) {
      return mapRevisionItems(stack, function (_, index) {
        return "rev-stack-" + index;
      });
    }
    return collectRevisionItems(api, { allowSyncEnd: false });
  }

  function resolveFreshRevisionRaw(doc, raw) {
    var id = revisionIdentity(raw);
    var qt = doc && doc.Wq && doc.Wq.wih && doc.Wq.wih();
    if (!id || !qt) return raw;
    for (var key in qt) {
      if (!Array.isArray(qt[key])) continue;
      for (var i = 0; i < qt[key].length; i += 1) {
        if (revisionIdentity(qt[key][i]) === id) return qt[key][i];
      }
    }
    return raw;
  }

  function findRevisionRaw(api, target) {
    var id = typeof target === "string" ? target : target && target.id;
    var index =
      target && typeof target.index === "number" ? target.index : undefined;
    ensureRevisionsIndexed(api);

    var raws = collectRevisionRawsFromWq(getLogicDocument(api));
    if (!raws.length) {
      raws = readRevisionStack(api, false, false);
    }

    if (id) {
      for (var i = 0; i < raws.length; i += 1) {
        var rawId =
          "rev-" +
          (typeof raws[i].ed === "function" ? String(raws[i].ed()) : i);
        if (rawId === id || "rev-stack-" + i === id) {
          return {
            raw: raws[i],
            index: i,
            item: collectRevisionItems(api, { allowSyncEnd: false })[i],
          };
        }
      }
    }

    if (typeof index === "number" && raws[index]) {
      return {
        raw: raws[index],
        index: index,
        item: collectRevisionItems(api, { allowSyncEnd: false })[index],
      };
    }

    return null;
  }

  function isRevisionGroupedMove(raw) {
    return raw && Array.isArray(raw.OL) && raw.OL.length > 0;
  }

  function applyRevisionSelection(doc, raw) {
    if (isRevisionGroupedMove(raw) && raw.vva) {
      var enums = getAscRevisionEnums();
      doc.lc && doc.lc();
      doc.Ihc && doc.Ihc(raw.vva, raw.nW === enums.MoveFrom, false, false);
      return true;
    }

    var element =
      raw.Element || (typeof raw.ee === "function" ? raw.ee() : undefined);
    var start = typeof raw.wa === "number" ? raw.wa : revisionStart(raw);
    if (element == null || start == null) return false;
    var end = typeof raw.xa === "number" ? raw.xa : revisionEnd(raw) || start;

    try {
      doc.lc && doc.lc();
      if (
        typeof element.Dp === "function" &&
        typeof element.Gs === "function"
      ) {
        if (doc.Wq && doc.Wq.u0d && doc.Wq.u0d([element])) return false;
        element.Dp(start, false, -1, -1);
        if (element.Selection) element.Selection.Na = true;
        element.Gs(start, end, false);
        element.Ft && element.Ft(false);
        return true;
      }

      if (
        typeof element.qp === "function" &&
        typeof element.mo === "function"
      ) {
        doc.ec && doc.ec();
        element.qp(start, false, -1, -1);
        if (element.Selection) element.Selection.La = true;
        element.mo(start, end, false);
        element.hq && element.hq(false);
        doc.Pf && doc.Pf(false);
        doc.yf && doc.yf(true);
        return true;
      }
    } catch (error) {
      return false;
    }
    return false;
  }

  function scrollToRevisionPosition(api, data) {
    if (data && data.Y != null && api.ViewScrollToY) api.ViewScrollToY(data.Y);
    if (data && data.X != null && api.ViewScrollToX) api.ViewScrollToX(data.X);
  }

  function activateRevision(doc, raw) {
    var wq = doc && doc.Wq;
    if (!wq) return;
    wq.t0d && wq.t0d();
    wq.TLh && wq.TLh(raw);
    wq.Ydf && wq.Ydf(true);
    if (doc.Ek && doc.Ek.f8) doc.Ek.f8(wq);
    wq.Znf && wq.Znf();
    doc.ze && doc.ze(true);
  }

  function navigateToRevisionByStep(api, doc, targetRaw, targetIndex) {
    if (typeof api.asc_GetNextRevisionsChange !== "function") return false;
    if (doc.Wq && doc.Wq.$_d) doc.Wq.$_d();
    var steps = Math.max(targetIndex + 1, 1);
    for (var i = 0; i < steps; i += 1) {
      api.asc_GetNextRevisionsChange();
      if (doc.Wq && doc.Wq.kz && isSameRevision(doc.Wq.kz, targetRaw))
        return true;
    }
    return !!(doc.Wq && doc.Wq.kz && isSameRevision(doc.Wq.kz, targetRaw));
  }

  function goToRevision(api, target) {
    var resolved = findRevisionRaw(api, target);
    if (!resolved || !resolved.raw) return false;
    var doc = getLogicDocument(api);
    if (!doc) return false;

    var raw = resolveFreshRevisionRaw(doc, resolved.raw);
    var item = resolved.item || {};
    var data = item.Data || revisionChangeToPlain(raw, getAscRevisionEnums());
    var moveType =
      (typeof raw.get_MoveType === "function"
        ? raw.get_MoveType()
        : undefined) || (typeof raw.hZd === "function" ? raw.hZd() : undefined);
    var enums = getAscRevisionEnums();

    if (
      moveType !== undefined &&
      (moveType === enums.MoveTo || moveType === enums.MoveFrom)
    ) {
      api.asc_FollowRevisionMove && api.asc_FollowRevisionMove(raw);
      scrollToRevisionPosition(api, data);
    }

    ensureRevisionsIndexed(api);
    if (doc.Wq && doc.Wq.TLh) doc.Wq.TLh(raw);
    if (applyRevisionSelection(doc, raw)) {
      scrollToRevisionPosition(api, data);
      activateRevision(doc, raw);
      return true;
    }
    if (navigateToRevisionByStep(api, doc, raw, resolved.index)) {
      scrollToRevisionPosition(api, data);
      return true;
    }
    scrollToRevisionPosition(api, data);
    activateRevision(doc, raw);
    return true;
  }

  function applyRevisionChange(api, mode, target) {
    var resolved = findRevisionRaw(api, target);
    if (!resolved || !resolved.raw) return false;
    goToRevision(api, target);
    var doc = getLogicDocument(api);
    var raw = doc ? resolveFreshRevisionRaw(doc, resolved.raw) : resolved.raw;
    if (mode === "accept") {
      api.asc_AcceptChanges && api.asc_AcceptChanges(raw);
    } else {
      api.asc_RejectChanges && api.asc_RejectChanges(raw);
    }
    return true;
  }

  function dismissReviewChangesDialog() {
    document
      .querySelectorAll(".asc-window.review-changes.modal-dlg")
      .forEach(function (el) {
        var close = el.querySelector(".close, .btn-close");
        if (close && close.click) {
          close.click();
        } else {
          el.style.display = "none";
        }
      });
  }

  function prepareRevisionReviewDisplay(api) {
    api.asc_SetLocalTrackRevisions && api.asc_SetLocalTrackRevisions(true);
    if (typeof api.pluginMethod_SetDisplayModeInReview === "function") {
      api.pluginMethod_SetDisplayModeInReview("markup");
      dismissReviewChangesDialog();
      return;
    }
    if (typeof api.asc_SetDisplayModeInReview === "function") {
      var Asc = window.Asc || {};
      var markup =
        (Asc.c_oAscReviewDisplay && Asc.c_oAscReviewDisplay.Edit) ||
        (Asc.c_oAscReviewDisplay && Asc.c_oAscReviewDisplay.Markup) ||
        (Asc.Xja && Asc.Xja.AKa);
      if (markup !== undefined) api.asc_SetDisplayModeInReview(markup);
    }
    dismissReviewChangesDialog();
  }

  function collectRevisionsAfterMutation(api) {
    return collectRevisionItems(api, { forceRefreshStack: true });
  }

  function registerEditorCallback(api, eventName) {
    if (!api.asc_registerCallback) {
      return false;
    }
    if (eventName !== EDITOR_EVENT.DOCUMENT_MODIFIED_CHANGED) {
      throw new Error("Unsupported OnlyOffice editor event: " + eventName);
    }
    if (editorCallbacksRegistered[eventName]) {
      return true;
    }

    api.asc_registerCallback(eventName, function () {
      post({
        type: BRIDGE_MESSAGE.EDITOR_EVENT,
        event: eventName,
        args: Array.prototype.slice.call(arguments),
      });
    });
    editorCallbacksRegistered[eventName] = true;
    return true;
  }

  function registerRevisionCallbacks(api) {
    if (revisionCallbacksRegistered || !api.asc_registerCallback) {
      return;
    }
    api.asc_registerCallback(
      EDITOR_EVENT.SHOW_REVISIONS_CHANGE,
      function (stack) {
        post({
          type: BRIDGE_MESSAGE.EDITOR_EVENT,
          event: EDITOR_EVENT.SHOW_REVISIONS_CHANGE,
          args: [resolveRevisionShowChanges(stack, api)],
        });
      },
    );
    api.asc_registerCallback(
      EDITOR_EVENT.TRACK_REVISIONS_CHANGE,
      function (enabled) {
        post({
          type: BRIDGE_MESSAGE.EDITOR_EVENT,
          event: EDITOR_EVENT.TRACK_REVISIONS_CHANGE,
          args: [!!enabled],
        });
      },
    );
    revisionCallbacksRegistered = true;
  }

  function createCommentData(payload) {
    var AscCtor = window.Asc && window.Asc.asc_CCommentDataWord;
    if (!AscCtor) {
      return payload;
    }

    var comment = new AscCtor(null);
    if (payload && payload.Text != null) {
      comment.asc_putText && comment.asc_putText(String(payload.Text));
    }
    if (payload && payload.UserName != null) {
      comment.asc_putUserName &&
        comment.asc_putUserName(String(payload.UserName));
    }
    if (payload && payload.Time != null) {
      comment.asc_putTime && comment.asc_putTime(String(payload.Time));
    }
    if (payload && payload.QuoteText != null) {
      comment.asc_putQuoteText &&
        comment.asc_putQuoteText(String(payload.QuoteText));
    }
    if (payload && typeof payload.Solved === "boolean") {
      comment.asc_putSolved && comment.asc_putSolved(payload.Solved);
    }
    if (payload && payload.UserData != null) {
      comment.asc_putUserData &&
        comment.asc_putUserData(String(payload.UserData));
    }
    return comment;
  }

  function runEditorCommand(command, payload) {
    var api = getEditorApi();
    if (!api) {
      throw new Error("OnlyOffice SDK API is not ready");
    }

    switch (command) {
      case EDITOR_COMMAND.EDITOR_SUBSCRIBE:
        return registerEditorCallback(api, payload && payload.event);
      case EDITOR_COMMAND.INTERFACE_SET_THEME: {
        var interfaceTheme = payload && payload.theme;
        if (interfaceTheme !== "theme-white" && interfaceTheme !== "theme-night") {
          throw new Error("Unsupported OnlyOffice interface theme");
        }
        var themes = window.Common && window.Common.UI && window.Common.UI.Themes;
        if (!themes || typeof themes.setTheme !== "function") {
          throw new Error("OnlyOffice native interface theme controller is not ready");
        }
        themes.setTheme(interfaceTheme, "sdk");
        return interfaceTheme;
      }
      case EDITOR_COMMAND.DOCUMENT_PRINT_PDF:
        if (typeof api.asc_nativeGetPDF !== "function") {
          throw new Error("OnlyOffice native PDF API is not available");
        }
        var previousNative = window.native;
        window.native = Object.assign({}, previousNative || {}, {
          Save_End:
            previousNative && typeof previousNative.Save_End === "function"
              ? previousNative.Save_End
              : function () {},
        });
        try {
          return api.asc_nativeGetPDF({ isPrint: true });
        } finally {
          if (previousNative) {
            window.native = previousNative;
          } else {
            delete window.native;
          }
        }
      case EDITOR_COMMAND.DOCUMENT_RENAME:
        if (!payload || !payload.fileName) {
          throw new Error("Document file name is required");
        }
        if (typeof api.asc_wopi_renameFile !== "function") {
          throw new Error("OnlyOffice WOPI rename API is not available");
        }
        api.asc_wopi_renameFile(String(payload.fileName));
        return true;
      case EDITOR_COMMAND.COMMENT_ADD: {
        var data = payload && payload.data ? payload.data : {};
        return (
          (api.pluginMethod_AddComment && api.pluginMethod_AddComment(data)) ||
          (api.asc_addComment && api.asc_addComment(createCommentData(data))) ||
          ""
        );
      }
      case EDITOR_COMMAND.COMMENT_UPDATE: {
        if (!payload || !payload.id) {
          throw new Error("Comment id is required");
        }
        var nextData = payload.data || {};
        if (typeof api.pluginMethod_ChangeComment === "function") {
          api.pluginMethod_ChangeComment(String(payload.id), nextData);
        } else if (api.asc_changeComment) {
          api.asc_changeComment(
            String(payload.id),
            createCommentData(nextData),
          );
        }
        return true;
      }
      case EDITOR_COMMAND.COMMENT_REMOVE: {
        if (!payload || !payload.id) {
          throw new Error("Comment id is required");
        }
        if (api.asc_removeComment) {
          api.asc_removeComment(String(payload.id));
        }
        return true;
      }
      case EDITOR_COMMAND.COMMENT_GO_TO: {
        if (!payload || !payload.id) {
          throw new Error("Comment id is required");
        }
        if (api.asc_selectComment) {
          api.asc_selectComment(String(payload.id));
        }
        if (payload.showBalloon && api.asc_showComment) {
          api.asc_showComment(String(payload.id));
        }
        return true;
      }
      case EDITOR_COMMAND.COMMENT_LIST:
        return getAllComments(api);
      case EDITOR_COMMAND.COMMENT_SUBSCRIBE:
        registerCommentCallbacks(api);
        return true;
      case EDITOR_COMMAND.REVISION_ADD_DEMO: {
        var text =
          payload && payload.text
            ? String(payload.text)
            : "审批修订 " + new Date().toLocaleTimeString();
        api.asc_SetGlobalTrackRevisions &&
          api.asc_SetGlobalTrackRevisions(true);
        api.asc_SetLocalTrackRevisions && api.asc_SetLocalTrackRevisions(true);
        if (api.pluginMethod_InputText) {
          api.pluginMethod_InputText(text);
        } else if (api.pluginMethod_PasteText) {
          api.pluginMethod_PasteText(text);
        } else if (api.asc_AddText) {
          api.asc_AddText(text);
        } else {
          throw new Error("OnlyOffice text insertion API is not available");
        }
        return collectRevisionsAfterMutation(api);
      }
      case EDITOR_COMMAND.REVISION_LIST:
        return collectRevisionItems(api, {
          forceRefreshStack: !!(payload && payload.forceRefreshStack),
        });
      case EDITOR_COMMAND.REVISION_SET_TRACK:
        api.asc_SetGlobalTrackRevisions &&
          api.asc_SetGlobalTrackRevisions(!!(payload && payload.enabled));
        api.asc_SetLocalTrackRevisions &&
          api.asc_SetLocalTrackRevisions(!!(payload && payload.enabled));
        return !!(payload && payload.enabled);
      case EDITOR_COMMAND.REVISION_IS_TRACK:
        return !!(
          api.asc_GetGlobalTrackRevisions && api.asc_GetGlobalTrackRevisions()
        );
      case EDITOR_COMMAND.REVISION_HAVE_CHANGES:
        return !!(
          (api.asc_HaveRevisionsChanges &&
            api.asc_HaveRevisionsChanges(true)) ||
          (api.asc_HaveRevisionsChanges && api.asc_HaveRevisionsChanges()) ||
          collectRevisionItems(api, { allowSyncEnd: false }).length > 0
        );
      case EDITOR_COMMAND.REVISION_PREPARE_REVIEW:
        api.asc_SetGlobalTrackRevisions &&
          api.asc_SetGlobalTrackRevisions(true);
        prepareRevisionReviewDisplay(api);
        return collectRevisionItems(api, { forceRefreshStack: true });
      case EDITOR_COMMAND.REVISION_NEXT:
        api.asc_GetNextRevisionsChange && api.asc_GetNextRevisionsChange();
        return true;
      case EDITOR_COMMAND.REVISION_PREV:
        api.asc_GetPrevRevisionsChange && api.asc_GetPrevRevisionsChange();
        return true;
      case EDITOR_COMMAND.REVISION_GO_TO:
        return goToRevision(api, payload || {});
      case EDITOR_COMMAND.REVISION_ACCEPT:
        applyRevisionChange(api, "accept", payload || {});
        return collectRevisionsAfterMutation(api);
      case EDITOR_COMMAND.REVISION_REJECT:
        applyRevisionChange(api, "reject", payload || {});
        return collectRevisionsAfterMutation(api);
      case EDITOR_COMMAND.REVISION_ACCEPT_ALL:
        api.asc_AcceptChanges && api.asc_AcceptChanges();
        return collectRevisionsAfterMutation(api);
      case EDITOR_COMMAND.REVISION_REJECT_ALL:
        api.asc_RejectChanges && api.asc_RejectChanges();
        return collectRevisionsAfterMutation(api);
      case EDITOR_COMMAND.REVISION_ACCEPT_SELECTION:
        api.asc_AcceptChangesBySelection &&
          api.asc_AcceptChangesBySelection(payload && payload.all);
        return collectRevisionsAfterMutation(api);
      case EDITOR_COMMAND.REVISION_REJECT_SELECTION:
        api.asc_RejectChangesBySelection &&
          api.asc_RejectChangesBySelection(payload && payload.all);
        return collectRevisionsAfterMutation(api);
      case EDITOR_COMMAND.REVISION_SUBSCRIBE:
        registerRevisionCallbacks(api);
        return true;
      default:
        throw new Error("Unsupported OnlyOffice editor command: " + command);
    }
  }

  function handleEditorCommand(message) {
    if (!message.requestId || !message.command) {
      return;
    }

    try {
      post({
        type: BRIDGE_MESSAGE.EDITOR_RESPONSE,
        requestId: message.requestId,
        result: runEditorCommand(message.command, message.payload || {}),
      });
    } catch (error) {
      post({
        type: BRIDGE_MESSAGE.EDITOR_RESPONSE,
        requestId: message.requestId,
        error: error && error.message ? error.message : String(error),
      });
    }
  }

  window.addEventListener("message", function (event) {
    var message = event.data;
    if (
      event.source !== window.parent ||
      event.origin !== parentOrigin ||
      !message ||
      message.source !== BRIDGE_SOURCE ||
      message.frameEditorId !== frameEditorId ||
      message.bridgeInstanceId !== bridgeInstanceId
    ) {
      return;
    }

    if (message.type === "hello:ack") {
      handshakeReady = true;
      window.clearInterval(helloTimer);
      flushSocketEmits();
      return;
    }

    if (message.type === "socket:event") {
      var registry = window.__ONLYOFFICE_SCOPED_IO__;
      var sockets = registry && registry.__activeSockets;
      if (sockets && sockets.length) {
        sockets.slice().forEach(function (socket) {
          if (socket.__onlyofficeDispatch) {
            socket.__onlyofficeDispatch(message.event, message.args || []);
          }
        });
      } else {
        queuedSocketEvents.push({
          event: message.event,
          args: message.args || [],
        });
      }
      return;
    }

    if (
      message.type === "http:response" &&
      message.requestId &&
      pendingHttp[message.requestId]
    ) {
      pendingHttp[message.requestId](message);
      delete pendingHttp[message.requestId];
      return;
    }

    if (message.type === BRIDGE_MESSAGE.EDITOR_SET_READONLY) {
      scheduleReadOnly(!!message.readOnly, 20);
      return;
    }

    if (message.type === BRIDGE_MESSAGE.EDITOR_COMMAND) {
      handleEditorCommand(message);
    }
  });

  var originalFactory = window.__ONLYOFFICE_SCOPED_IO__[frameEditorId];
  window.__ONLYOFFICE_SCOPED_IO__.__activeSockets =
    window.__ONLYOFFICE_SCOPED_IO__.__activeSockets || [];
  window.__ONLYOFFICE_SCOPED_IO__[frameEditorId] = function () {
    var socket = originalFactory.apply(this, arguments);
    window.__ONLYOFFICE_SCOPED_IO__.__activeSockets.push(socket);
    if (queuedSocketEvents.length) {
      queuedSocketEvents.splice(0).forEach(function (queued) {
        socket.__onlyofficeDispatch(queued.event, queued.args || []);
      });
    }
    return socket;
  };

  installFetchProxy();
  installXhrProxy();
  installNamedDownloadPatch(100);
  installPrintFramePatch();
})();
