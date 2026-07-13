/* Puente JavaScript ↔ Kotlin de MEE SEG. */
(function () {
  const pending = new Map();
  let seq = 0;

  function nextId() {
    seq += 1;
    return "cb_" + Date.now().toString(36) + "_" + seq;
  }

  window.__meeResolve = function (id, payloadJson) {
    const entry = pending.get(id);
    if (!entry) return;
    pending.delete(id);
    clearTimeout(entry.timer);
    try { entry.resolve(payloadJson ? JSON.parse(payloadJson) : null); }
    catch (_) { entry.resolve(payloadJson); }
  };

  window.__meeReject = function (id, message) {
    const entry = pending.get(id);
    if (!entry) return;
    pending.delete(id);
    clearTimeout(entry.timer);
    entry.reject(new Error(message || "Operación cancelada."));
  };

  const MeeNative = {
    isNative() {
      return !!(window.MeeBridge && typeof window.MeeBridge.pickMaster === "function");
    },

    masterInfo() {
      if (!this.isNative() || typeof window.MeeBridge.getMasterInfo !== "function") return null;
      try { return JSON.parse(window.MeeBridge.getMasterInfo() || "null"); }
      catch (_) { return null; }
    },

    saveDiagnostics() {
      if (!this.isNative() || typeof window.MeeBridge.getSaveDiagnostics !== "function") return null;
      try { return JSON.parse(window.MeeBridge.getSaveDiagnostics() || "null"); }
      catch (_) { return null; }
    },

    pendingSubmissionInfo() {
      if (!this.isNative() || typeof window.MeeBridge.getPendingSubmissionInfo !== "function") return null;
      try { return JSON.parse(window.MeeBridge.getPendingSubmissionInfo() || "null"); }
      catch (_) { return null; }
    },

    getSyncDestinationEmail() {
      if (!this.isNative() || typeof window.MeeBridge.getSyncDestinationEmail !== "function") return "";
      try { return String(window.MeeBridge.getSyncDestinationEmail() || ""); }
      catch (_) { return ""; }
    },

    setSyncDestinationEmail(email) {
      if (!this.isNative() || typeof window.MeeBridge.setSyncDestinationEmail !== "function") return false;
      try { return !!window.MeeBridge.setSyncDestinationEmail(String(email || "")); }
      catch (_) { return false; }
    },

    _call(method, ...args) {
      return new Promise((resolve, reject) => {
        if (!this.isNative() || typeof window.MeeBridge[method] !== "function") {
          reject(new Error("Puente nativo no disponible: " + method));
          return;
        }
        const id = nextId();
        const timer = setTimeout(() => {
          if (!pending.has(id)) return;
          pending.delete(id);
          reject(new Error("La operación tardó demasiado. Volvé a intentarlo."));
        }, 120000);
        pending.set(id, { resolve, reject, timer });
        try { window.MeeBridge[method](id, ...args); }
        catch (error) {
          clearTimeout(timer);
          pending.delete(id);
          reject(error);
        }
      });
    },

    pickMaster() { return this._call("pickMaster"); },

    rememberMaster(uri, fileName) {
      if (!this.isNative() || typeof window.MeeBridge.rememberMaster !== "function") return false;
      window.MeeBridge.rememberMaster(String(uri || ""), String(fileName || "MEE_DATOS_COMITE_MASTER.json"), false, "readOnlyByDesign");
      return true;
    },

    readMaster() { return this._call("readMaster"); },

    createPendingSubmission(payloadJson, baseVersion, baseExportedAt, mode = "NORMAL") {
      return this._call(
        "createPendingSubmission",
        String(payloadJson || ""),
        Number(baseVersion || 0),
        String(baseExportedAt || ""),
        String(mode || "NORMAL")
      );
    },

    sharePendingSubmission(action = "outlook") {
      return this._call("sharePendingSubmission", String(action || "outlook"));
    },

    confirmPendingSubmission() { return this._call("confirmPendingSubmission"); },

    confirmPendingSubmissionAt(uri, fileName) {
      return this._call("confirmPendingSubmissionAt", String(uri || ""), String(fileName || ""));
    },

    pickSubmissionFolder() { return this._call("pickSubmissionFolder"); },

    sendPendingToFolder() { return this._call("sendPendingToFolder"); },

    completePendingSubmission(submissionId, serverAppliedAt) {
      return this._call("completePendingSubmission", String(submissionId || ""), String(serverAppliedAt || ""));
    },

    saveBinary(base64, fileName, mime, action) {
      return this._call("saveBinary", base64, fileName, mime, action || "save");
    }
  };

  function reportNativeError(message) {
    try {
      if (window.MeeBridge && typeof window.MeeBridge.reportError === "function") {
        window.MeeBridge.reportError(String(message || "Error JavaScript"));
      }
    } catch (_) {}
  }

  window.addEventListener("error", event => {
    reportNativeError(`${event.message || "Error"} @${event.filename || ""}:${event.lineno || 0}`);
  });
  window.addEventListener("unhandledrejection", event => {
    reportNativeError(event.reason?.stack || event.reason?.message || String(event.reason || "Promise rechazada"));
  });

  window.MeeNative = MeeNative;
})();
