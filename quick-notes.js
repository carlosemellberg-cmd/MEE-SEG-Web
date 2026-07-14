class QuickNotesStore {
  constructor() {
    this.storageKey = "mee_seg_quick_notes_v1101";
    this.draftPrefix = "mee_seg_quick_note_draft_v1101_";
  }

  canManage(user) {
    return ["ADMINISTRADOR", "JEFE"].includes(user?.role);
  }

  assertCanManage(user) {
    if (!this.canManage(user)) {
      throw new Error("El rol Operador es de solo lectura.");
    }
  }

  userKey(user) {
    return String(user?.email || "sin_usuario").trim().toLowerCase();
  }

  readAll() {
    try {
      const value = JSON.parse(localStorage.getItem(this.storageKey) || "[]");
      return Array.isArray(value) ? value : [];
    } catch (_) {
      return [];
    }
  }

  writeAll(notes) {
    localStorage.setItem(this.storageKey, JSON.stringify(Array.isArray(notes) ? notes : []));
  }

  createId() {
    if (globalThis.crypto && typeof globalThis.crypto.randomUUID === "function") {
      return globalThis.crypto.randomUUID();
    }
    return `NOTE-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  }

  listForUser(user) {
    const key = this.userKey(user);
    return this.readAll()
      .filter(note => note.authorEmail === key)
      .sort((a, b) => new Date(b.updatedAt || b.createdAt) - new Date(a.updatedAt || a.createdAt));
  }

  listPending(user) {
    return this.listForUser(user).filter(note => note.status !== "CONVERTIDO");
  }

  get(user, id) {
    return this.listForUser(user).find(note => note.id === id) || null;
  }

  getDraft(user) {
    try {
      const draft = JSON.parse(localStorage.getItem(this.draftPrefix + this.userKey(user)) || "null");
      return draft && typeof draft.text === "string" ? draft : { text: "", updatedAt: "" };
    } catch (_) {
      return { text: "", updatedAt: "" };
    }
  }

  saveDraft(user, text) {
    this.assertCanManage(user);
    const draft = {
      text: String(text || ""),
      updatedAt: new Date().toISOString()
    };
    localStorage.setItem(this.draftPrefix + this.userKey(user), JSON.stringify(draft));
    return draft;
  }

  clearDraft(user) {
    this.assertCanManage(user);
    localStorage.removeItem(this.draftPrefix + this.userKey(user));
  }

  createFromDraft(user) {
    this.assertCanManage(user);
    const draft = this.getDraft(user);
    const text = String(draft.text || "").trim();
    if (!text) throw new Error("Escribí el apunte antes de guardarlo.");

    const now = new Date().toISOString();
    const note = {
      id: this.createId(),
      text,
      status: "PENDIENTE",
      authorEmail: this.userKey(user),
      authorName: user?.name || "Usuario local",
      createdAt: now,
      updatedAt: now,
      convertedAt: "",
      convertedTaskCode: ""
    };

    const notes = this.readAll();
    notes.push(note);
    this.writeAll(notes);
    this.clearDraft(user);
    return note;
  }

  update(user, id, text) {
    this.assertCanManage(user);
    const value = String(text || "").trim();
    if (!value) throw new Error("El apunte no puede quedar vacío.");

    const key = this.userKey(user);
    const notes = this.readAll();
    const note = notes.find(item => item.id === id && item.authorEmail === key);
    if (!note) throw new Error("Apunte no encontrado.");

    note.text = value;
    note.updatedAt = new Date().toISOString();
    this.writeAll(notes);
    return note;
  }

  remove(user, id) {
    this.assertCanManage(user);
    const key = this.userKey(user);
    const notes = this.readAll();
    const next = notes.filter(item => !(item.id === id && item.authorEmail === key));
    if (next.length === notes.length) throw new Error("Apunte no encontrado.");
    this.writeAll(next);
    return true;
  }

  markConverted(user, id, taskCode) {
    this.assertCanManage(user);
    const key = this.userKey(user);
    const notes = this.readAll();
    const note = notes.find(item => item.id === id && item.authorEmail === key);
    if (!note) return null;

    note.status = "CONVERTIDO";
    note.convertedTaskCode = String(taskCode || "");
    note.convertedAt = new Date().toISOString();
    note.updatedAt = note.convertedAt;
    this.writeAll(notes);
    return note;
  }

  reopen(user, id) {
    this.assertCanManage(user);
    const key = this.userKey(user);
    const notes = this.readAll();
    const note = notes.find(item => item.id === id && item.authorEmail === key);
    if (!note) throw new Error("Apunte no encontrado.");

    note.status = "PENDIENTE";
    note.convertedTaskCode = "";
    note.convertedAt = "";
    note.updatedAt = new Date().toISOString();
    this.writeAll(notes);
    return note;
  }
}

window.QuickNotesStore = QuickNotesStore;
