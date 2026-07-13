class SimplePdfDocument {
  constructor(title = "Comité de Seguridad MEE") {
    this.pageWidth = 595;
    this.pageHeight = 842;
    this.marginX = 45;
    this.marginTop = 48;
    this.marginBottom = 48;
    this.pages = [[]];
    this.pageIndex = 0;
    this.y = this.pageHeight - this.marginTop;
    this.title = title;
  }

  clean(value) {
    return String(value ?? "")
      .replace(/[\u2013\u2014]/g, "-")
      .replace(/[\u2018\u2019]/g, "'")
      .replace(/[\u201c\u201d]/g, '"')
      .replace(/\u2022/g, "-")
      .replace(/\t/g, " ")
      .replace(/\r/g, "");
  }

  newPage() {
    this.pages.push([]);
    this.pageIndex += 1;
    this.y = this.pageHeight - this.marginTop;
  }

  ensureSpace(required) {
    if (this.y - required < this.marginBottom) this.newPage();
  }

  wrap(text, size, indent = 0) {
    const cleaned = this.clean(text);
    const usable = this.pageWidth - (this.marginX * 2) - indent;
    const maxChars = Math.max(18, Math.floor(usable / (size * 0.52)));
    const result = [];
    for (const paragraph of cleaned.split("\n")) {
      const words = paragraph.trim().split(/\s+/).filter(Boolean);
      if (!words.length) {
        result.push("");
        continue;
      }
      let line = "";
      for (const word of words) {
        if (!line) {
          line = word;
        } else if (`${line} ${word}`.length <= maxChars) {
          line += ` ${word}`;
        } else {
          result.push(line);
          line = word;
        }
      }
      if (line) result.push(line);
    }
    return result;
  }

  addText(text, options = {}) {
    const size = Number(options.size || 10);
    const bold = !!options.bold;
    const indent = Number(options.indent || 0);
    const lineHeight = Number(options.lineHeight || Math.ceil(size * 1.38));
    const before = Number(options.before || 0);
    const after = Number(options.after ?? 4);
    if (before) this.y -= before;
    const lines = this.wrap(text, size, indent);
    for (const line of lines) {
      this.ensureSpace(lineHeight + after);
      this.pages[this.pageIndex].push({
        x: this.marginX + indent,
        y: this.y,
        size,
        bold,
        text: line
      });
      this.y -= lineHeight;
    }
    this.y -= after;
    return this;
  }

  heading(text, level = 1) {
    const sizes = {1: 18, 2: 13, 3: 11};
    return this.addText(text, { size: sizes[level] || 11, bold: true, before: level === 1 ? 0 : 7, after: 6 });
  }

  field(label, value) {
    return this.addText(`${label}: ${value || "Sin informar"}`, { size: 10, after: 2 });
  }

  bullet(text, indent = 12) {
    return this.addText(`- ${text}`, { size: 9.5, indent, after: 2 });
  }

  encodeWinAnsi(text) {
    const map = new Map([
      [0x20ac,0x80],[0x201a,0x82],[0x0192,0x83],[0x201e,0x84],[0x2026,0x85],
      [0x2020,0x86],[0x2021,0x87],[0x02c6,0x88],[0x2030,0x89],[0x0160,0x8a],
      [0x2039,0x8b],[0x0152,0x8c],[0x017d,0x8e],[0x2018,0x91],[0x2019,0x92],
      [0x201c,0x93],[0x201d,0x94],[0x2022,0x95],[0x2013,0x96],[0x2014,0x97],
      [0x02dc,0x98],[0x2122,0x99],[0x0161,0x9a],[0x203a,0x9b],[0x0153,0x9c],
      [0x017e,0x9e],[0x0178,0x9f]
    ]);
    const bytes = [];
    for (const char of text) {
      const code = char.codePointAt(0);
      if (code <= 0xff) bytes.push(code);
      else if (map.has(code)) bytes.push(map.get(code));
      else bytes.push(0x3f);
    }
    return new Uint8Array(bytes);
  }

  pdfEscape(text) {
    return this.clean(text).replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)");
  }

  concat(chunks) {
    const length = chunks.reduce((sum, item) => sum + item.length, 0);
    const output = new Uint8Array(length);
    let offset = 0;
    for (const item of chunks) {
      output.set(item, offset);
      offset += item.length;
    }
    return output;
  }

  buildBytes() {
    const objects = [];
    objects[1] = this.encodeWinAnsi("<< /Type /Catalog /Pages 2 0 R >>");
    objects[3] = this.encodeWinAnsi("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");
    objects[4] = this.encodeWinAnsi("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>");

    const pageRefs = [];
    this.pages.forEach((items, index) => {
      const pageObj = 5 + index * 2;
      const contentObj = pageObj + 1;
      pageRefs.push(`${pageObj} 0 R`);
      const commands = [];
      commands.push("0.12 0.12 0.14 rg");
      for (const item of items) {
        const font = item.bold ? "F2" : "F1";
        commands.push(`BT /${font} ${item.size.toFixed(1)} Tf 1 0 0 1 ${item.x.toFixed(1)} ${item.y.toFixed(1)} Tm (${this.pdfEscape(item.text)}) Tj ET`);
      }
      commands.push(`BT /F1 8 Tf 0.4 0.4 0.4 rg 1 0 0 1 45 24 Tm (Pagina ${index + 1} de ${this.pages.length}) Tj ET`);
      const contentBytes = this.encodeWinAnsi(commands.join("\n"));
      objects[contentObj] = this.concat([
        this.encodeWinAnsi(`<< /Length ${contentBytes.length} >>\nstream\n`),
        contentBytes,
        this.encodeWinAnsi("\nendstream")
      ]);
      objects[pageObj] = this.encodeWinAnsi(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${this.pageWidth} ${this.pageHeight}] /Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents ${contentObj} 0 R >>`);
    });
    objects[2] = this.encodeWinAnsi(`<< /Type /Pages /Kids [${pageRefs.join(" ")}] /Count ${this.pages.length} >>`);

    const chunks = [this.encodeWinAnsi("%PDF-1.4\n%\xE2\xE3\xCF\xD3\n")];
    const offsets = [0];
    let current = chunks[0].length;
    for (let i = 1; i < objects.length; i += 1) {
      offsets[i] = current;
      const prefix = this.encodeWinAnsi(`${i} 0 obj\n`);
      const suffix = this.encodeWinAnsi("\nendobj\n");
      chunks.push(prefix, objects[i], suffix);
      current += prefix.length + objects[i].length + suffix.length;
    }
    const xrefOffset = current;
    let xref = `xref\n0 ${objects.length}\n0000000000 65535 f \n`;
    for (let i = 1; i < objects.length; i += 1) {
      xref += `${String(offsets[i]).padStart(10,"0")} 00000 n \n`;
    }
    xref += `trailer\n<< /Size ${objects.length} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;
    chunks.push(this.encodeWinAnsi(xref));
    return this.concat(chunks);
  }

  toBlob() {
    return new Blob([this.buildBytes()], { type: "application/pdf" });
  }
}

/* ------------------------------------------------------------------ *
 *  FileSyncManager (v1.9.3) — reescrito para Android nativo (SAF).
 *  - Import real vía puente nativo (SAF ACTION_OPEN_DOCUMENT).
 *  - Guardado real del maestro sobre la URI recordada, o "Guardar como".
 *  - Control de conflictos: relee y compara dataVersion antes de escribir.
 *  - PDF / Excel se entregan por SAF (Guardar) o FileProvider (Compartir).
 *  - En PC (sin puente) se conserva la ruta web original.
 *  El generador SimplePdfDocument NO se modificó.
 * ------------------------------------------------------------------ */
class FileSyncManager {
  constructor(adapter, config) {
    this.adapter = adapter;
    this.config = config;
    this.masterFileHandle = null;   // solo PC
    this.dbName = "MEE_FileSync_v1";
    this.storeName = "handles";
    this.webPendingKey = "mee_web_pending_submission_v1";
    this.webEmailKey = "mee_web_sync_destination_email_v1";
  }

  native() { return window.MeeNative && window.MeeNative.isNative() ? window.MeeNative : null; }
  isAndroidApp() { return !!this.native() || /Android/i.test(navigator.userAgent || ""); }
  webPendingInfo() {
    try { return JSON.parse(localStorage.getItem(this.webPendingKey) || "null"); }
    catch (_) { return null; }
  }
  hasPendingSubmission() {
    const nativePending = this.native()?.pendingSubmissionInfo?.();
    return !!nativePending?.hasPending || !!this.webPendingInfo()?.hasPending;
  }
  setWebPending(value) { localStorage.setItem(this.webPendingKey, JSON.stringify(value)); }
  clearWebPending() { localStorage.removeItem(this.webPendingKey); }
  getSyncDestinationEmail() {
    const nat = this.native();
    return nat ? (nat.getSyncDestinationEmail?.() || "") : (localStorage.getItem(this.webEmailKey) || "");
  }
  setSyncDestinationEmail(value) {
    const email = String(value || "").trim();
    if (!/^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$/.test(email)) return false;
    const nat = this.native();
    if (nat) return !!nat.setSyncDestinationEmail?.(email);
    localStorage.setItem(this.webEmailKey, email);
    return true;
  }

  async initialize() {
    if (this.native()) { this.masterFileHandle = null; return; }
    try { this.masterFileHandle = await this.loadHandle("master"); }
    catch (error) { console.warn("No se pudo restaurar el archivo vinculado.", error); }
  }

  /* ---------- IndexedDB (solo ruta PC / File System Access API) ---------- */
  async openDb() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(this.dbName, 1);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(this.storeName)) db.createObjectStore(this.storeName);
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }
  async saveHandle(key, handle) {
    const db = await this.openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(this.storeName, "readwrite");
      tx.objectStore(this.storeName).put(handle, key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }
  async loadHandle(key) {
    const db = await this.openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(this.storeName, "readonly");
      const request = tx.objectStore(this.storeName).get(key);
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(request.error);
    });
  }

  /* ---------- utilidades ---------- */
  dateStamp() {
    const now = new Date();
    const pad = v => String(v).padStart(2, "0");
    return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
  }
  bytesToBase64(bytes) {
    let binary = "";
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode(...bytes.subarray(i, Math.min(i + chunk, bytes.length)));
    }
    return btoa(binary);
  }
  utf8ToBase64(text) { return this.bytesToBase64(new TextEncoder().encode(text)); }
  parseJson(text, fileName = "archivo") {
    try { return JSON.parse(text); }
    catch (error) { throw new Error(`${fileName} no contiene un JSON válido.`); }
  }
  escapeHtml(value) {
    return String(value ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
  }
  date(value) {
    if (!value) return "Sin fecha";
    try { return new Intl.DateTimeFormat("es-AR", { dateStyle: "short", timeStyle: "short" }).format(new Date(value)); }
    catch { return String(value); }
  }
  peopleMap(store) { return new Map((store.people || []).map(p => [p.email, p])); }

  /* ---------- descarga/compartir (ruta PC) ---------- */
  download(content, fileName, mimeType) {
    const blob = content instanceof Blob ? content : new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url; link.download = fileName;
    document.body.appendChild(link); link.click(); link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }
  async shareBlobWeb(blob, fileName, title, text = "Archivo del Comité de Seguridad MEE") {
    const file = new File([blob], fileName, { type: blob.type || "application/octet-stream" });
    if (navigator.canShare && navigator.canShare({ files: [file] }) && navigator.share) {
      try {
        await navigator.share({ title, text, files: [file] });
        return { shared: true, fileName };
      } catch (error) {
        console.warn("No se pudo compartir; se descargara el archivo.", error);
      }
    }
    this.download(blob, fileName, blob.type || "application/octet-stream");
    return { shared: false, downloaded: true, fileName };
  }

  // Entrega un binario según plataforma. En Android usa SAF/FileProvider real.
  async deliverBlob(blob, fileName, title, action = "save") {
    const nat = this.native();
    if (nat) {
      const bytes = new Uint8Array(await blob.arrayBuffer());
      const res = await nat.saveBinary(this.bytesToBase64(bytes), fileName, blob.type || "application/octet-stream", action);
      return { shared: action === "share", saved: res?.mode === "created" || res?.mode === "overwrite", fileName: res?.fileName || fileName, uri: res?.uri || "" };
    }
    return this.shareBlobWeb(blob, fileName, title);
  }

  /* ---------- IMPORTAR MAESTRO ---------- */
  validateMasterSelection(pkg, fileName = "archivo") {
    const lower = String(fileName || "").trim().toLowerCase();
    if (lower !== "mee_datos_comite_master.json") {
      throw new Error("Seleccioná exactamente MEE_DATOS_COMITE_MASTER.json. Renombrá las copias con (1), (2) o MASTER1 antes de continuar.");
    }
    if (pkg?.fileType === "MEE_SUBMISSION" || pkg?.submissionType === "MEE_MASTER_SUBMISSION") {
      throw new Error("El archivo seleccionado es un envío MEE_SUBMISSION, no el archivo maestro.");
    }
    this.adapter.validateMasterPackage(pkg);
  }

  async pickAndImportMaster() {
    const nat = this.native();
    if (!nat) return { fallback: true };
    if (nat.pendingSubmissionInfo?.()?.hasPending) {
      throw new Error("Existe una sincronización pendiente. Confirmala antes de vincular otro archivo maestro.");
    }
    const picked = await nat.pickMaster();
    const pkg = this.parseJson(picked.text, picked.fileName);
    this.validateMasterSelection(pkg, picked.fileName);
    const result = await this.adapter.importMasterPackage(pkg, picked.fileName);
    nat.rememberMaster(picked.uri, picked.fileName);
    return { ...result, linkedFile: picked.fileName, canWrite: false, writeStatus: "readOnlyByDesign" };
  }

  async confirmPendingSync() {
    const nat = this.native();
    if (!nat) return { confirmed: false, status: "select_file", detail: "Seleccioná el maestro actualizado descargado de SharePoint." };
    const pending = nat.pendingSubmissionInfo?.() || null;
    if (!pending?.hasPending) return { confirmed: false, status: "none", detail: "No hay un envío pendiente." };

    const result = await nat.confirmPendingSubmission();
    return this.finishConfirmation(nat, result);
  }

  /**
   * Confirmación contra un maestro elegido manualmente (solo lectura).
   * Sirve cuando el URI recordado quedó apuntando a otro archivo o a una copia vieja.
   * No cambia el archivo recordado ni escribe nada en OneDrive.
   */
  async confirmPendingSyncWithPickedFile() {
    const nat = this.native();
    if (!nat) return { confirmed: false, status: "select_file", detail: "Seleccioná el maestro actualizado descargado de SharePoint." };
    const pending = nat.pendingSubmissionInfo?.() || null;
    if (!pending?.hasPending) return { confirmed: false, status: "none", detail: "No hay un envío pendiente." };

    const picked = await nat.pickMaster();
    const result = await nat.confirmPendingSubmissionAt(picked.uri, picked.fileName || "");
    return this.finishConfirmation(nat, result);
  }

  async confirmPendingSyncWithFile(file) {
    const pending = this.webPendingInfo();
    if (!pending?.hasPending) return { confirmed: false, status: "none", detail: "No hay un envío pendiente." };
    const text = await file.text();
    const remote = this.parseJson(text, file.name);
    this.validateMasterSelection(remote, file.name);
    if (String(remote.masterId || "") !== String(pending.masterId || "")) {
      return { confirmed: false, status: "conflict", detail: "El archivo pertenece a otra base maestra." };
    }
    const remoteVersion = Number(remote.dataVersion || 0);
    const target = Number(pending.targetDataVersion || 0);
    const ack = String(remote.syncSubmission?.lastAppliedSubmissionId || "");
    const payloadVersion = Number(remote.syncSubmission?.payloadDataVersion || 0);
    if (remoteVersion < target || !ack) {
      return { confirmed: false, status: "pending", remoteDataVersion: remoteVersion, targetDataVersion: target, ackSubmissionId: ack, detail: `El maestro está en versión ${remoteVersion}; se espera la versión ${target}.` };
    }
    if (remoteVersion > target || (remoteVersion === target && ack !== pending.submissionId)) {
      return { confirmed: false, status: "conflict", remoteDataVersion: remoteVersion, targetDataVersion: target, ackSubmissionId: ack, detail: "El maestro fue actualizado por otro envío o por una versión posterior." };
    }
    if (payloadVersion !== target) {
      return { confirmed: false, status: "pending", detail: "El ACK todavía no confirma la versión objetivo." };
    }
    await this.adapter.importMasterPackage(remote, file.name);
    this.clearWebPending();
    return { confirmed: true, status: "confirmed", submissionId: pending.submissionId, dataVersion: target, serverAppliedAt: remote.syncSubmission?.serverAppliedAt || "" };
  }

  /** Cierre atómico compartido: commit local + limpieza del pending SOLO con ACK confirmado. */
  async finishConfirmation(nat, result) {
    if (!result?.confirmed) return result || { confirmed: false, status: "pending" };

    const confirmedMaster = result.remoteMaster || result.payload;
    if (!confirmedMaster) throw new Error("El ACK fue válido, pero no se recibió el maestro remoto confirmado.");
    await this.adapter.importMasterPackage(
      confirmedMaster,
      result.masterFileName || nat.masterInfo?.()?.fileName || "MEE_DATOS_COMITE_MASTER.json"
    );
    await nat.completePendingSubmission(result.submissionId, result.serverAppliedAt);
    return {
      confirmed: true,
      status: "confirmed",
      submissionId: result.submissionId,
      dataVersion: result.targetDataVersion,
      serverAppliedAt: result.serverAppliedAt
    };
  }

  /** Autorización única de la carpeta MEE_SEG_ENTRADAS. */
  async configureSubmissionFolder() {
    const nat = this.native();
    if (!nat) throw new Error("El envío por carpeta está disponible únicamente en Android.");
    return nat.pickSubmissionFolder();
  }

  /** Deposita el pending como archivo nuevo en la carpeta autorizada. El pending se conserva hasta el ACK. */
  async sendPendingToFolder() {
    const nat = this.native();
    if (!nat) throw new Error("El envío por carpeta está disponible únicamente en Android.");
    const pending = nat.pendingSubmissionInfo?.() || null;
    if (!pending?.hasPending) throw new Error("No hay un envío pendiente para depositar.");
    return nat.sendPendingToFolder();
  }

  async reloadRememberedMaster() {
    const nat = this.native();
    if (!nat) return { fallback: true };

    const pending = nat.pendingSubmissionInfo?.() || null;
    if (pending?.hasPending) {
      const confirmation = await this.confirmPendingSync();
      if (confirmation.confirmed) {
        return { linkedFile: nat.masterInfo()?.fileName || "MEE_DATOS_COMITE_MASTER.json", confirmedSubmission: true, ...confirmation };
      }
      throw new Error(confirmation.detail || "El servidor todavía no confirmó el envío pendiente.");
    }

    const current = await nat.readMaster();
    if (current?.invalid) {
      const error = new Error(
        `${current.fileName || "El archivo maestro"} está vacío o incompleto. ` +
        "Los datos locales siguen protegidos. Usá Guardar cambios para preparar una recuperación segura."
      );
      error.code = "REMOTE_INVALID";
      throw error;
    }
    if (!current || current.missing || !current.text) return { missing: true };
    const pkg = this.parseJson(current.text, current.fileName || "MEE_DATOS_COMITE_MASTER.json");
    this.validateMasterSelection(pkg, current.fileName || "MEE_DATOS_COMITE_MASTER.json");
    const result = await this.adapter.importMasterPackage(pkg, current.fileName || "MEE_DATOS_COMITE_MASTER.json");
    return { ...result, linkedFile: current.fileName || "MEE_DATOS_COMITE_MASTER.json", canWrite: false, writeStatus: "readOnlyByDesign" };
  }

  async importMasterFile(file) {
    const text = await file.text();
    const pkg = this.parseJson(text, file.name);
    this.validateMasterSelection(pkg, file.name);
    return this.adapter.importMasterPackage(pkg, file.name);
  }

  async linkAndImportMaster() {
    if (this.native()) return this.pickAndImportMaster();
    if (!window.showOpenFilePicker) return { fallback: true };
    const [handle] = await window.showOpenFilePicker({
      multiple: false,
      types: [{ description: "Base maestra Comité MEE", accept: { "application/json": [".json"] } }]
    });
    const file = await handle.getFile();
    const result = await this.importMasterFile(file);
    this.masterFileHandle = handle;
    await this.saveHandle("master", handle);
    return { ...result, linkedFile: file.name };
  }

  /* ---------- CONTROL DE CONFLICTOS: SOLO LECTURA ---------- */
  async assertNoRemoteConflict() {
    const nat = this.native();
    if (!nat) return null;
    const local = await this.adapter.getSyncInfo();
    const remote = await nat.readMaster();
    if (remote?.invalid) {
      const error = new Error(
        `${remote.fileName || "El archivo maestro"} está vacío o no contiene un JSON válido. ` +
        "Se preparará un envío de recuperación; Android no sobrescribirá el archivo."
      );
      error.code = "REMOTE_INVALID";
      throw error;
    }
    if (!remote || remote.missing || !remote.text) return null;

    let remotePkg;
    try { remotePkg = this.parseJson(remote.text, remote.fileName || "archivo maestro"); }
    catch (parseError) { parseError.code = "REMOTE_INVALID"; throw parseError; }
    this.validateMasterSelection(remotePkg, remote.fileName || "archivo maestro");
    if (remotePkg.masterId && local.masterId && remotePkg.masterId !== local.masterId) {
      throw new Error("El archivo recordado pertenece a otra base maestra. Elegí nuevamente el archivo correcto.");
    }

    const remoteVersion = Number(remotePkg.dataVersion || 0);
    const localVersion = Number(local.dataVersion || 0);
    if (remoteVersion !== localVersion) {
      const error = new Error(
        `Conflicto de versión. La base externa está en la versión ${remoteVersion} y tu aplicación parte de la versión ${localVersion}. ` +
        "Actualizá desde el archivo maestro antes de preparar un envío."
      );
      error.code = "CONFLICT";
      throw error;
    }
    return remotePkg;
  }

  /* ---------- SINCRONIZACIÓN SEGURA ---------- */
  async saveMasterToLinkedFile() {
    const nat = this.native();
    if (!nat) {
      const existing = this.webPendingInfo();
      if (existing?.hasPending) return { mode: "pendingExists", pending: existing, pendingPreserved: true };
      const sync = await this.adapter.getSyncInfo();
      if (!sync.dirtyCount) return { noChanges: true };
      const payload = await this.adapter.prepareMasterPackage();
      const submissionId = (globalThis.crypto && typeof globalThis.crypto.randomUUID === "function")
        ? globalThis.crypto.randomUUID()
        : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
      const submittedAt = new Date().toISOString();
      const submission = {
        fileType: "MEE_SUBMISSION",
        submissionType: "MEE_MASTER_SUBMISSION",
        submissionMode: "NORMAL",
        schemaVersion: "1.0",
        submissionId,
        masterId: sync.masterId,
        baseDataVersion: Number(sync.dataVersion || 0),
        targetDataVersion: Number(payload.dataVersion || 0),
        baseExportedAt: sync.baseExportedAt || "",
        submittedAt,
        submittedBy: payload.exportedBy,
        payload
      };
      const fileName = `MEE_SUBMISSION_${submission.targetDataVersion}_${submissionId}.json`;
      const pending = { hasPending: true, submissionId, masterId: sync.masterId, baseDataVersion: submission.baseDataVersion, targetDataVersion: submission.targetDataVersion, submittedAt, fileName, submissionMode: "NORMAL", submission };
      this.setWebPending(pending);
      return { mode: "submissionPrepared", pending, pendingPreserved: true, recovery: false };
    }

    const existing = nat.pendingSubmissionInfo?.() || null;
    if (existing?.hasPending) return { mode: "pendingExists", pending: existing, pendingPreserved: true };

    const sync = await this.adapter.getSyncInfo();
    let recovery = false;
    try {
      await this.assertNoRemoteConflict();
    } catch (error) {
      if (error?.code === "REMOTE_INVALID") recovery = true;
      else throw error;
    }

    if (!sync.dirtyCount && !recovery) {
      return { noChanges: true, linkedFile: nat.masterInfo()?.fileName || "" };
    }

    const payload = recovery
      ? await this.adapter.prepareRecoveryMasterPackage()
      : await this.adapter.prepareMasterPackage();
    const mode = recovery ? "RECOVERY" : "NORMAL";
    const created = await nat.createPendingSubmission(
      JSON.stringify(payload),
      Number(sync.dataVersion || 0),
      String(sync.baseExportedAt || ""),
      mode
    );
    return { mode: "submissionPrepared", pending: created, recovery, pendingPreserved: true };
  }

  async sharePendingSubmission(action = "outlook") {
    const nat = this.native();
    if (nat) {
      const pending = nat.pendingSubmissionInfo?.() || null;
      if (!pending?.hasPending) throw new Error("No hay un envío pendiente para compartir.");
      return nat.sharePendingSubmission(action);
    }
    const pending = this.webPendingInfo();
    if (!pending?.hasPending || !pending.submission) throw new Error("No hay un envío pendiente para compartir.");
    const blob = new Blob([JSON.stringify(pending.submission, null, 2)], { type: "application/json" });
    const result = await this.shareBlobWeb(blob, pending.fileName, "Envío MEE SEG", "Adjuntar este JSON al correo de Power Automate");
    if (action === "outlook" && !result.shared) {
      const email = this.getSyncDestinationEmail();
      const subject = `MEE_SEG_SYNC | ${pending.masterId} | ${pending.submissionId} | v${pending.targetDataVersion}`;
      const body = `Adjuntar el archivo descargado: ${pending.fileName}`;
      const target = `mailto:${encodeURIComponent(email)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
      setTimeout(() => { window.location.href = target; }, 300);
    }
    return { ...result, pendingPreserved: true, fileName: pending.fileName };
  }

  async exportMasterDownload() {
    if (this.native()) return this.saveMasterToLinkedFile();
    const sync = await this.adapter.getSyncInfo();
    if (!sync.dirtyCount) return { noChanges: true };
    const pkg = await this.adapter.prepareMasterPackage();
    const fileName = this.config.sharePointMasterFileName || "MEE_DATOS_COMITE_MASTER.json";
    const result = await this.shareBlobWeb(new Blob([JSON.stringify(pkg, null, 2)], { type: "application/json" }), fileName, "Enviar archivo maestro Comité MEE");
    await this.adapter.commitMasterSave(pkg);
    return { ...pkg, ...result };
  }

  async exportChangesDownload() {
    const pkg = await this.adapter.buildChangesPackage();
    if (!pkg.tasks.length && !(pkg.deletedTaskCodes || []).length) return { empty: true };
    const email = String(pkg.exportedBy.email || "usuario").split("@")[0].replace(/[^a-z0-9_-]/gi, "_");
    const name = `MEE_CAMBIOS_${email}_${this.dateStamp()}.json`;
    const blob = new Blob([JSON.stringify(pkg, null, 2)], { type: "application/json" });
    const res = await this.deliverBlob(blob, name, "Compartir cambios Comité MEE", "share");
    await this.adapter.markChangesExported();
    return { empty: false, count: pkg.tasks.length + (pkg.deletedTaskCodes || []).length, fileName: name, ...res };
  }

  async consolidateFiles(files) {
    const packages = [];
    for (const file of files) packages.push(this.parseJson(await file.text(), file.name));
    return this.adapter.mergeChangePackages(packages);
  }

  /* ---------- PDF minuta por tarea ---------- */
  async generateTaskPdf(code) {
    const store = this.adapter.readStore();
    const task = (store.tasks || []).find(item => item.code === code);
    if (!task) throw new Error("Tarea no encontrada para generar la minuta.");
    const people = this.peopleMap(store);
    const personName = email => people.get(email)?.name || email || "Sin asignar";
    const pdf = new SimplePdfDocument(`Minuta ${task.code}`);
    pdf.heading("COMITÉ DE SEGURIDAD MEE", 1)
      .addText(`MINUTA DE TAREA — ${task.code}`, { size: 13, bold: true, after: 8 })
      .field("Generada", new Date().toLocaleString("es-AR"))
      .field("Generada por", store.currentUser?.name || "Usuario local")
      .heading(task.title, 2)
      .field("Estado", task.statusLabel)
      .field("Avance", `${task.progress || 0} %`)
      .field("Responsable", task.responsible)
      .field("Ubicación", task.location)
      .field("Prioridad", task.priority)
      .field("Grupo", task.group);
    pdf.heading("Problema detectado", 2).addText(task.problem || "Sin informar.");
    pdf.heading("Propuesta / solución planteada", 2).addText(task.proposal || "Sin propuesta registrada.");
    pdf.heading("Participantes generales", 2);
    if ((task.participants || []).length) (task.participants || []).forEach(email => pdf.bullet(personName(email)));
    else pdf.addText("Sin participantes generales.");
    pdf.heading("Aportes, opiniones y decisiones", 2);
    if ((task.contributions || []).length) {
      task.contributions.forEach(item => {
        pdf.addText(`${item.type || "APORTE"} — ${item.author || "Sin autor"} — ${this.date(item.at)}`, { bold: true, size: 9.5, after: 2 });
        pdf.addText(item.text || "", { size: 9.5, indent: 8, after: 5 });
      });
    } else pdf.addText("Sin aportes registrados.");
    pdf.heading("Plan de acción por sector", 2);
    if ((task.workItems || []).length) {
      task.workItems.forEach((item, index) => {
        pdf.addText(`${index + 1}. ${item.title || "Acción"}`, { bold: true, size: 10.5, before: 4, after: 2 });
        pdf.field("Sector", item.sector).field("Responsable", personName(item.responsibleEmail));
        pdf.field("Participantes", (item.participantEmails || []).map(personName).join(", ") || "Sin participantes");
        pdf.field("Avance", `${item.progress || 0} %`).field("Estado", item.status).field("Fecha objetivo", item.dueDate || "Sin fecha");
        if (item.description) pdf.addText(item.description, { indent: 8, size: 9.5 });
        if ((item.updates || []).length) {
          pdf.addText("Avances informados:", { bold: true, size: 9.5, after: 2 });
          item.updates.forEach(update => pdf.bullet(`${update.progress || 0}% — ${personName(update.authorEmail)} — ${this.date(update.at)} — ${update.comment || "Sin comentario"}`, 10));
        }
      });
    } else pdf.addText("Sin acciones por sector.");
    pdf.heading("Pendientes", 2);
    if ((task.pending || []).length) task.pending.forEach(item => pdf.bullet(`${item.status || "ABIERTO"} — ${item.description || ""} — Responsable: ${item.responsible || "Sin asignar"}`));
    else pdf.addText("Sin pendientes registrados.");
    if (task.closeResult) pdf.heading("Resultado de cierre", 2).addText(task.closeResult).field("Cerrada por", task.closedBy).field("Fecha de cierre", this.date(task.closedAt));
    pdf.heading("Historial", 2);
    if ((task.history || []).length) task.history.forEach(item => pdf.bullet(`${this.date(item.at)} — ${item.by || ""}: ${item.text || ""}`));
    else pdf.addText("Sin historial.");
    const fileName = `MEE_MINUTA_${String(task.code).replace(/[^A-Z0-9_-]/gi,"_")}_${this.dateStamp()}.pdf`;
    return { ...(await this.deliverBlob(pdf.toBlob(), fileName, `Minuta ${task.code}`, "save")), fileName };
  }

  /* ---------- PDF reporte general ---------- */
  async generatePdfReport() {
    const store = this.adapter.readStore();
    const tasks = store.tasks || [];
    const pdf = new SimplePdfDocument("Reporte Comité MEE");
    pdf.heading("COMITÉ DE SEGURIDAD MEE", 1)
      .addText("REPORTE GENERAL DE TAREAS", { size: 13, bold: true, after: 8 })
      .field("Generado", new Date().toLocaleString("es-AR"))
      .field("Generado por", store.currentUser?.name || "Usuario local")
      .field("Tareas totales", tasks.length)
      .field("Tareas activas", tasks.filter(item => item.status !== "FINALIZADA").length)
      .field("Pendientes abiertos", tasks.reduce((sum, item) => sum + (item.pending || []).filter(p => p.status !== "COMPLETADO").length, 0));
    tasks.forEach((task, index) => {
      pdf.heading(`${index + 1}. ${task.code} — ${task.title}`, 2)
        .field("Estado", task.statusLabel)
        .field("Avance", `${task.progress || 0} %`)
        .field("Responsable", task.responsible)
        .field("Ubicación", task.location)
        .addText(`Problema: ${task.problem || "Sin informar"}`, { size: 9.5, after: 3 })
        .addText(`Propuesta: ${task.proposal || "Sin informar"}`, { size: 9.5, after: 3 })
        .field("Acciones por sector", (task.workItems || []).length)
        .field("Aportes", (task.contributions || []).length)
        .field("Pendientes", (task.pending || []).length)
        .field("Participantes", (task.participants || []).length);
    });
    const fileName = `MEE_REPORTE_COMITE_${this.dateStamp()}.pdf`;
    return { ...(await this.deliverBlob(pdf.toBlob(), fileName, "Reporte Comité MEE", "save")), fileName };
  }

  /* ---------- Excel (HTML .xls, abre en Excel) ---------- */
  async generateExcelReport() {
    const store = this.adapter.readStore();
    const tasks = store.tasks || [];
    const peopleByEmail = this.peopleMap(store);
    const personName = email => peopleByEmail.get(email)?.name || email || "Sin asignar";
    const rowsTasks = tasks.map(task => `<tr><td>${this.escapeHtml(task.code)}</td><td>${this.escapeHtml(task.title)}</td><td>${this.escapeHtml(task.statusLabel)}</td><td>${task.progress || 0}%</td><td>${this.escapeHtml(task.responsible)}</td><td>${this.escapeHtml(task.location)}</td><td>${this.escapeHtml(task.priority)}</td><td>${this.escapeHtml(task.problem)}</td><td>${this.escapeHtml(task.proposal)}</td></tr>`).join("");
    const rowsActions = tasks.flatMap(task => (task.workItems || []).map(item => `<tr><td>${this.escapeHtml(task.code)}</td><td>${this.escapeHtml(item.title)}</td><td>${this.escapeHtml(item.sector)}</td><td>${this.escapeHtml(personName(item.responsibleEmail))}</td><td>${this.escapeHtml((item.participantEmails || []).map(personName).join(", "))}</td><td>${item.progress || 0}%</td><td>${this.escapeHtml(item.status)}</td><td>${this.escapeHtml(item.dueDate)}</td><td>${this.escapeHtml((item.updates || []).slice(-1)[0]?.comment || "")}</td></tr>`)).join("");
    const rowsContributions = tasks.flatMap(task => (task.contributions || []).map(item => `<tr><td>${this.escapeHtml(task.code)}</td><td>${this.escapeHtml(task.title)}</td><td>${this.escapeHtml(item.type)}</td><td>${this.escapeHtml(item.author)}</td><td>${this.escapeHtml(item.text)}</td><td>${this.escapeHtml(this.date(item.at))}</td></tr>`)).join("");
    const rowsPending = tasks.flatMap(task => (task.pending || []).map(item => `<tr><td>${this.escapeHtml(task.code)}</td><td>${this.escapeHtml(task.title)}</td><td>${this.escapeHtml(item.description)}</td><td>${this.escapeHtml(item.status)}</td><td>${this.escapeHtml(item.responsible)}</td></tr>`)).join("");
    const rowsParticipants = tasks.flatMap(task => (task.participants || []).map(email => `<tr><td>${this.escapeHtml(task.code)}</td><td>${this.escapeHtml(task.title)}</td><td>${this.escapeHtml(personName(email))}</td><td>${this.escapeHtml(email)}</td></tr>`)).join("");
    const active = tasks.filter(task => task.status !== "FINALIZADA").length;
    const openPending = tasks.reduce((total, task) => total + (task.pending || []).filter(item => item.status !== "COMPLETADO").length, 0);
    const html = `<!doctype html><html><head><meta charset="utf-8"><style>body{font-family:Arial,sans-serif;color:#222}h1{background:#222;color:#fff;padding:14px}h2{color:#d96b20;margin-top:28px}table{border-collapse:collapse;width:100%;margin-bottom:22px}th{background:#444;color:#fff}th,td{border:1px solid #aaa;padding:7px;text-align:left}.kpi{font-weight:bold;font-size:16px}</style></head><body><h1>Comité de Seguridad MEE — Reporte</h1><p>Generado: ${this.escapeHtml(new Date().toLocaleString("es-AR"))}</p><table><tr><th>Tareas totales</th><th>Tareas activas</th><th>Pendientes abiertos</th></tr><tr class="kpi"><td>${tasks.length}</td><td>${active}</td><td>${openPending}</td></tr></table><h2>Tareas generales</h2><table><tr><th>Código</th><th>Título</th><th>Estado</th><th>Avance</th><th>Responsable</th><th>Ubicación</th><th>Prioridad</th><th>Problema</th><th>Propuesta</th></tr>${rowsTasks}</table><h2>Plan de acción por sector</h2><table><tr><th>Tarea</th><th>Acción</th><th>Sector</th><th>Responsable</th><th>Participantes</th><th>Avance</th><th>Estado</th><th>Fecha objetivo</th><th>Último avance</th></tr>${rowsActions}</table><h2>Aportes y opiniones</h2><table><tr><th>Tarea</th><th>Título</th><th>Tipo</th><th>Autor</th><th>Detalle</th><th>Fecha</th></tr>${rowsContributions}</table><h2>Pendientes</h2><table><tr><th>Tarea</th><th>Título</th><th>Pendiente</th><th>Estado</th><th>Responsable</th></tr>${rowsPending}</table><h2>Participantes generales</h2><table><tr><th>Tarea</th><th>Título</th><th>Participante</th><th>Correo</th></tr>${rowsParticipants}</table></body></html>`;
    const fileName = `MEE_REPORTE_COMITE_${this.dateStamp()}.xls`;
    const blob = new Blob([`\ufeff${html}`], { type: "application/vnd.ms-excel;charset=utf-8" });
    return { ...(await this.deliverBlob(blob, fileName, "Reporte Excel Comité MEE", "save")), fileName };
  }

  /* ---------- estado mostrado en la tarjeta de archivo ---------- */
  async getStatus() {
    const info = await this.adapter.getSyncInfo();
    const nat = this.native();
    const master = nat ? nat.masterInfo() : null;
    const diag = nat ? nat.saveDiagnostics() : null;
    const pending = nat ? nat.pendingSubmissionInfo() : null;
    return {
      ...info,
      linkedFileName: master?.fileName || info.lastImportedFile || this.masterFileHandle?.name || "",
      canWrite: false,
      writeStatus: master ? (master.writeStatus || "readOnlyByDesign") : undefined,
      lastSavedAt: master?.lastSavedAt || "",
      lastReadAt: master?.lastReadAt || info.lastImportedAt || "",
      syncDestinationEmail: this.getSyncDestinationEmail(),
      submissionFolderName: master?.submissionFolderName || "",
      hasSubmissionFolder: !!master?.hasSubmissionFolder,
      pendingSubmission: nat ? (pending?.hasPending ? pending : null) : (this.webPendingInfo()?.hasPending ? this.webPendingInfo() : null),
      diagUriAuthority: diag?.uriAuthority || "",
      diagMimeType: diag?.mimeType || "",
      diagColumnFlags: diag?.columnFlags || "",
      diagLastReadSize: diag?.lastReadSize || "",
      diagLastReadAt: diag?.lastReadAt || "",
      diagLastWriteAttemptAt: diag?.lastWriteAttemptAt || "",
      diagLastWriteResult: diag?.lastWriteResult || ""
    };
  }

}

window.SimplePdfDocument = SimplePdfDocument;
window.FileSyncManager = FileSyncManager;
