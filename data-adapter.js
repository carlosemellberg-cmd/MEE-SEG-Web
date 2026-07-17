class DataAdapter {
  constructor(config) {
    this.config = config;
    this.mode = config.connectionMode || "demo";
    this.storageKey = "mee_comite_web_public_v1";
    this.deviceKey = "mee_comite_device_id_web_v1";
  }

  async initialize() {
    if (this.mode === "demo") return this.initializeDemo();
    if (this.mode === "file") return this.initializeFile();
    if (this.mode === "graph-test") return this.initializeDemo();
    if (this.mode === "powerAutomate") return this.initializePowerAutomate();
    if (this.mode === "graph") return this.initializeGraph();
    throw new Error(`Modo de conexión no soportado: ${this.mode}`);
  }

  initializeDemo() {
    const current = localStorage.getItem(this.storageKey);
    if (!current) {
      localStorage.setItem(this.storageKey, JSON.stringify(window.DEMO_DATA));
    }
    return true;
  }



  initializeFile() {
    this.initializeDemo();
    const store = this.readStore();
    this.ensureStoreShape(store);
    this.writeStore(store, { skipTracking: true });
    return true;
  }

  createId(prefix = "ID") {
    const raw = (globalThis.crypto && typeof globalThis.crypto.randomUUID === "function")
      ? globalThis.crypto.randomUUID()
      : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
    return `${prefix}-${raw}`;
  }

  getDeviceId() {
    let value = localStorage.getItem(this.deviceKey);
    if (!value) {
      value = this.createId("DEVICE");
      localStorage.setItem(this.deviceKey, value);
    }
    return value;
  }

  ensureStoreShape(store) {
    store.people = Array.isArray(store.people) ? store.people : [];
    store.tasks = Array.isArray(store.tasks) ? store.tasks : [];
    store.currentUser = store.currentUser || window.DEMO_DATA.currentUser;
    const currentPerson = (store.people || []).find(person => person.email === store.currentUser?.email);
    if (currentPerson) {
      const normalizedRole = this.roleForPerson(currentPerson);
      store.currentUser = {
        ...store.currentUser,
        name: currentPerson.name,
        email: currentPerson.email,
        cargo: currentPerson.cargo,
        group: currentPerson.group,
        role: normalizedRole.role,
        roleLabel: normalizedRole.roleLabel
      };
    }
    store.syncMeta = store.syncMeta || {};
    store.syncMeta.schemaVersion = "1.1";
    store.syncMeta.masterId = store.syncMeta.masterId || this.createId("MASTER");
    store.syncMeta.dataVersion = Number(store.syncMeta.dataVersion || 1);
    store.syncMeta.baseExportedAt = store.syncMeta.baseExportedAt || store.syncMeta.lastImportedExportedAt || "";
    store.syncMeta.deviceId = this.getDeviceId();
    store.syncMeta.dirtyTaskCodes = Array.isArray(store.syncMeta.dirtyTaskCodes) ? store.syncMeta.dirtyTaskCodes : [];
    store.syncMeta.deletedTaskCodes = Array.isArray(store.syncMeta.deletedTaskCodes) ? store.syncMeta.deletedTaskCodes : [];
    store.syncMeta.deletedTasks = Array.isArray(store.syncMeta.deletedTasks) ? store.syncMeta.deletedTasks : [];
    store.tasks.forEach(task => {
      task.sync = task.sync || {
        revision: 0,
        lastModifiedAt: task.createdAt || new Date().toISOString(),
        lastModifiedBy: task.createdBy || "Inicial",
        deviceId: store.syncMeta.deviceId
      };
    });
    return store;
  }

  taskComparable(task) {
    const copy = JSON.parse(JSON.stringify(task || {}));
    delete copy.sync;
    return copy;
  }

  initializePowerAutomate() {
    throw new Error("Conector Power Automate pendiente de configurar.");
  }

  initializeGraph() {
    throw new Error("Conector Microsoft Graph pendiente de configurar.");
  }

  readStore() {
    const parsed = JSON.parse(localStorage.getItem(this.storageKey) || "{}");
    return this.ensureStoreShape(parsed);
  }

  hasPendingSubmission() {
    try {
      const nativeInfo = window.MeeNative?.pendingSubmissionInfo?.();
      if (nativeInfo?.hasPending) return true;
      const webPending = JSON.parse(localStorage.getItem("mee_web_pending_submission_v1") || "null");
      return !!webPending?.hasPending;
    } catch (_) {
      return false;
    }
  }

  writeStore(store, options = {}) {
    if (!options.skipTracking && this.hasPendingSubmission()) {
      throw new Error("Existe una sincronización pendiente. Confirmala antes de realizar nuevos cambios.");
    }
    const previousRaw = localStorage.getItem(this.storageKey);
    const previous = previousRaw ? this.ensureStoreShape(JSON.parse(previousRaw)) : { tasks: [], syncMeta: { dirtyTaskCodes: [] } };
    this.ensureStoreShape(store);

    if (!options.skipTracking) {
      const previousByCode = new Map((previous.tasks || []).map(task => [task.code, task]));
      const dirty = new Set(store.syncMeta.dirtyTaskCodes || []);
      const now = new Date().toISOString();
      let changed = false;

      const currentCodes = new Set((store.tasks || []).map(task => task.code));
      const deleted = new Set(store.syncMeta.deletedTaskCodes || []);

      for (const task of store.tasks) {
        const oldTask = previousByCode.get(task.code);
        const isChanged = !oldTask || JSON.stringify(this.taskComparable(oldTask)) !== JSON.stringify(this.taskComparable(task));
        if (isChanged) {
          task.sync = {
            revision: Number(oldTask?.sync?.revision || 0) + 1,
            lastModifiedAt: now,
            lastModifiedBy: store.currentUser?.email || store.currentUser?.name || "Usuario local",
            deviceId: store.syncMeta.deviceId
          };
          deleted.delete(task.code);
          dirty.add(task.code);
          changed = true;
        }
      }

      const deletedAudit = new Map((store.syncMeta.deletedTasks || []).map(item => [item.code, item]));
      for (const oldTask of (previous.tasks || [])) {
        if (!currentCodes.has(oldTask.code)) {
          deleted.add(oldTask.code);
          dirty.add(oldTask.code);
          if (!deletedAudit.has(oldTask.code)) {
            deletedAudit.set(oldTask.code, {
              code: oldTask.code,
              title: oldTask.title || "",
              deletedAt: now,
              deletedBy: store.currentUser?.name || "Usuario local",
              deletedByEmail: store.currentUser?.email || "",
              deviceId: store.syncMeta.deviceId,
              lastRevision: Number(oldTask.sync?.revision || 0)
            });
          }
          changed = true;
        }
      }

      store.syncMeta.deletedTaskCodes = [...deleted];
      store.syncMeta.deletedTasks = [...deletedAudit.values()];
      store.syncMeta.dirtyTaskCodes = [...dirty];
      if (changed) store.syncMeta.lastLocalChangeAt = now;
    }

    localStorage.setItem(this.storageKey, JSON.stringify(store));
  }

  async getCurrentUser() {
    return this.readStore().currentUser;
  }

  async getPeople() {
    return this.readStore().people || [];
  }

  roleForPerson(person) {
    const normalize = value => String(value || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/\s+/g, " ")
      .trim();
    const email = normalize(person?.email);
    const cargo = normalize(person?.cargo);
    const name = normalize(person?.name);
    const authorizedPeople = [
      ["adrian", "espeche"],
      ["esteban", "cajal"],
      ["javier", "dos", "santos"],
      ["cesar", "campos"],
      ["diego", "conde"],
      ["alejandro", "dos", "santos"],
      ["jorge", "alvarez"],
      ["daniel", "arevalo"]
    ];
    const authorizedByName = authorizedPeople.some(tokens => tokens.every(token => name.includes(token)));
    if (email === "cemellberg@lomanegra.com" || cargo.includes("administrador")) {
      return { role: "ADMINISTRADOR", roleLabel: "Administrador" };
    }
    if (["supervisor", "coordinador", "lider", "jefe", "analista"].some(value => cargo.includes(value)) || authorizedByName) {
      return { role: "JEFE", roleLabel: "Gestor" };
    }
    return { role: "OPERADOR", roleLabel: "Operador" };
  }

  canManageStore(store) {
    return ["ADMINISTRADOR", "JEFE"].includes(store?.currentUser?.role);
  }

  assertCanManage(store, action = "modificar datos") {
    if (!this.canManageStore(store)) {
      throw new Error(`El rol Operador es de solo lectura y no puede ${action}.`);
    }
  }

  async setCurrentUser(email) {
    const store = this.readStore();
    const person = (store.people || []).find(item => item.email === email);
    if (!person) throw new Error("Usuario local no encontrado.");
    const parts = String(person.name || email).trim().split(/\s+/).filter(Boolean);
    const initials = `${parts[0]?.[0] || "U"}${parts[parts.length - 1]?.[0] || ""}`.toUpperCase();
    const role = this.roleForPerson(person);
    store.currentUser = {
      name: person.name,
      email: person.email,
      role: role.role,
      roleLabel: role.roleLabel,
      cargo: person.cargo,
      group: person.group,
      initials
    };
    this.writeStore(store, { skipTracking: true });
    return store.currentUser;
  }

  async getTasks() {
    return this.readStore().tasks || [];
  }

  async getTask(code) {
    return (await this.getTasks()).find(t => t.code === code);
  }

  async createTask(payload) {
    const store = this.readStore();
    this.assertCanManage(store, "crear tareas");
    const maxNumber = (store.tasks || []).reduce((max, item) => {
      const match = String(item.code || "").match(/(\d+)$/);
      return Math.max(max, match ? Number(match[1]) : 0);
    }, 0);
    const code = `MEE-2026-${String(maxNumber + 1).padStart(4, "0")}`;
    const task = {
      code,
      title: payload.title,
      problem: payload.problem,
      location: payload.location || "Sin especificar",
      type: payload.type,
      priority: payload.priority,
      status: "PRESENTADA",
      statusLabel: "Presentada",
      responsible: payload.responsible || "Sin asignar",
      group: payload.group || "GENERAL",
      progress: 0,
      createdBy: store.currentUser.name,
      createdAt: new Date().toISOString(),
      proposal: payload.proposal || "",
      workItems: [],
      contributions: [],
      pending: [],
      participants: (store.people || []).filter(person => person.group === (payload.group || "GENERAL")).map(person => person.email),
      documents: [],
      history: [{ text: "Tarea creada", by: store.currentUser.name, at: new Date().toISOString() }]
    };
    store.tasks = [task, ...(store.tasks || [])];
    this.writeStore(store);
    return task;
  }

  async editTaskText(code, field, value) {
    if (!['problem','proposal'].includes(field)) {
      throw new Error('Campo de tarea no permitido.');
    }
    const store = this.readStore();
    this.assertCanManage(store, "modificar tareas");
    const task = store.tasks.find(t => t.code === code);
    if (!task) throw new Error('Tarea no encontrada.');
    const before = task[field] || '';
    task[field] = value;
    const now = new Date().toISOString();
    task.history.unshift({
      text: field === 'problem' ? 'Editó el problema detectado' : 'Editó la propuesta inicial',
      by: store.currentUser.name,
      at: now,
      before
    });
    this.writeStore(store);
    return task;
  }

  recalculateTask(task) {
    const items=task.workItems || [];
    if(items.length){task.progress=Math.round(items.reduce((sum,w)=>sum+Number(w.progress||0),0)/items.length);}
    if(task.status!=="FINALIZADA"){
      if(task.progress>0) {task.status="EN_EJECUCION";task.statusLabel="En ejecución";}
      if(items.length && items.every(w=>w.progress===100)){task.status="PENDIENTE_VALIDACION";task.statusLabel="Pendiente de validación";}
    }
  }

  async addParticipant(code,email) {
    const store=this.readStore();this.assertCanManage(store,"modificar participantes");const task=store.tasks.find(t=>t.code===code); task.participants=task.participants||[];
    if(!task.participants.includes(email)) task.participants.push(email);
    task.history.unshift({text:"Agregó un participante general",by:store.currentUser.name,at:new Date().toISOString()});this.writeStore(store);return task;
  }

  async removeParticipant(code,email) {
    const store=this.readStore();this.assertCanManage(store,"modificar participantes");const task=store.tasks.find(t=>t.code===code); task.participants=(task.participants||[]).filter(x=>x!==email);
    task.history.unshift({text:"Eliminó un participante general",by:store.currentUser.name,at:new Date().toISOString()});this.writeStore(store);return task;
  }

  async addWorkItem(code,payload) {
    const store=this.readStore();this.assertCanManage(store,"crear tareas por sector");const task=store.tasks.find(t=>t.code===code);task.workItems=task.workItems||[];
    const item={id:crypto.randomUUID?crypto.randomUUID():String(Date.now()),title:payload.title,sector:payload.sector,description:payload.description||"",responsibleEmail:payload.responsibleEmail,participantEmails:payload.participantEmails||[payload.responsibleEmail],progress:0,status:"PLANIFICADA",dueDate:payload.dueDate||"",updates:[]};
    task.workItems.push(item);this.recalculateTask(task);task.history.unshift({text:`Creó la tarea por sector: ${item.title}`,by:store.currentUser.name,at:new Date().toISOString()});this.writeStore(store);return task;
  }

  async editWorkItem(code,id,payload) {
    const store=this.readStore();this.assertCanManage(store,"modificar tareas por sector");const task=store.tasks.find(t=>t.code===code),item=(task.workItems||[]).find(w=>w.id===id);if(!item)return task;
    Object.assign(item,{title:payload.title,sector:payload.sector,description:payload.description||"",responsibleEmail:payload.responsibleEmail,participantEmails:payload.participantEmails||[],dueDate:payload.dueDate||""});
    task.history.unshift({text:`Editó la tarea por sector: ${item.title}`,by:store.currentUser.name,at:new Date().toISOString()});this.writeStore(store);return task;
  }

  async deleteWorkItem(code,id) {
    const store=this.readStore();this.assertCanManage(store,"eliminar tareas por sector");const task=store.tasks.find(t=>t.code===code),item=(task.workItems||[]).find(w=>w.id===id);task.workItems=(task.workItems||[]).filter(w=>w.id!==id);this.recalculateTask(task);
    task.history.unshift({text:`Eliminó la tarea por sector: ${item?.title||id}`,by:store.currentUser.name,at:new Date().toISOString()});this.writeStore(store);return task;
  }

  async setWorkItemParticipants(code,id,responsibleEmail,participantEmails) {
    const store=this.readStore();this.assertCanManage(store,"modificar participantes");const task=store.tasks.find(t=>t.code===code),item=(task.workItems||[]).find(w=>w.id===id);if(!item)return task;
    item.responsibleEmail=responsibleEmail;item.participantEmails=[...new Set(participantEmails||[])];
    task.history.unshift({text:`Actualizó participantes de: ${item.title}`,by:store.currentUser.name,at:new Date().toISOString()});this.writeStore(store);return task;
  }

  async addWorkProgress(code,id,progress,comment) {
    const store=this.readStore();this.assertCanManage(store,"registrar avances");const task=store.tasks.find(t=>t.code===code),item=(task.workItems||[]).find(w=>w.id===id);if(!item)return task;
    item.progress=progress;item.status=progress>=100?"COMPLETADA":progress>0?"EN_EJECUCION":"PLANIFICADA";item.updates=item.updates||[];
    item.updates.push({id:crypto.randomUUID?crypto.randomUUID():String(Date.now()),progress,comment,authorEmail:store.currentUser.email,at:new Date().toISOString()});
    this.recalculateTask(task);task.history.unshift({text:`Informó avance ${progress}% en: ${item.title}`,by:store.currentUser.name,at:new Date().toISOString()});this.writeStore(store);return task;
  }

  async addContribution(code, text, type = "PROPUESTA") {
    const store = this.readStore();
    this.assertCanManage(store, "agregar aportes");
    const task = store.tasks.find(t => t.code === code);
    task.contributions.push({
      id: crypto.randomUUID ? crypto.randomUUID() : String(Date.now()),
      type,
      text,
      author: store.currentUser.name,
      group: store.currentUser.group,
      at: new Date().toISOString()
    });
    task.history.unshift({ text: `Agregó ${type.toLowerCase()}`, by: store.currentUser.name, at: new Date().toISOString() });
    this.writeStore(store);
    return task;
  }

  async addPending(code, description) {
    const store = this.readStore();
    this.assertCanManage(store, "agregar pendientes");
    const task = store.tasks.find(t => t.code === code);
    task.pending.push({
      id: crypto.randomUUID ? crypto.randomUUID() : String(Date.now()),
      description,
      status: "ABIERTO",
      responsible: store.currentUser.name,
      at: new Date().toISOString()
    });
    task.history.unshift({ text: "Agregó un pendiente", by: store.currentUser.name, at: new Date().toISOString() });
    this.writeStore(store);
    return task;
  }

  async editPending(code, pendingId, description) {
    const store = this.readStore();
    this.assertCanManage(store, "modificar pendientes");
    const task = store.tasks.find(t => t.code === code);
    const item = task.pending.find(p => p.id === pendingId);
    if (item) {
      item.description = description;
      item.updatedAt = new Date().toISOString();
      task.history.unshift({ text: "Editó un pendiente", by: store.currentUser.name, at: item.updatedAt });
      this.writeStore(store);
    }
    return task;
  }

  async deletePending(code, pendingId) {
    const store = this.readStore();
    this.assertCanManage(store, "eliminar pendientes");
    const task = store.tasks.find(t => t.code === code);
    task.pending = task.pending.filter(p => p.id !== pendingId);
    task.history.unshift({ text: "Eliminó un pendiente", by: store.currentUser.name, at: new Date().toISOString() });
    this.writeStore(store);
    return task;
  }

  async togglePending(code, pendingId) {
    const store = this.readStore();
    this.assertCanManage(store, "modificar pendientes");
    const task = store.tasks.find(t => t.code === code);
    const item = task.pending.find(p => p.id === pendingId);
    if (item) {
      item.status = item.status === "COMPLETADO" ? "ABIERTO" : "COMPLETADO";
      item.updatedAt = new Date().toISOString();
      task.history.unshift({ text: item.status === "COMPLETADO" ? "Completó un pendiente" : "Reabrió un pendiente", by: store.currentUser.name, at: item.updatedAt });
      this.writeStore(store);
    }
    return task;
  }

  async deleteTask(code) {
    const store = this.readStore();
    this.assertCanManage(store, "eliminar tareas");
    const task = (store.tasks || []).find(item => item.code === code);
    if (!task) throw new Error("Tarea no encontrada.");
    store.tasks = (store.tasks || []).filter(item => item.code !== code);
    this.writeStore(store);
    return { code, title: task.title };
  }

  async closeTask(code, result) {
    const store = this.readStore();
    this.assertCanManage(store, "finalizar tareas");
    const task = store.tasks.find(t => t.code === code);
    task.status = "FINALIZADA";
    task.statusLabel = "Finalizada";
    task.progress = 100;
    task.closedBy = store.currentUser.name;
    task.closedAt = new Date().toISOString();
    task.closeResult = result;
    task.history.unshift({ text: "Finalizó la tarea", by: store.currentUser.name, at: task.closedAt });
    this.writeStore(store);
    return task;
  }



  async getSyncInfo() {
    const store = this.readStore();
    return {
      masterId: store.syncMeta.masterId,
      dataVersion: store.syncMeta.dataVersion,
      baseExportedAt: store.syncMeta.baseExportedAt || "",
      dirtyCount: (store.syncMeta.dirtyTaskCodes || []).length,
      dirtyTaskCodes: [...(store.syncMeta.dirtyTaskCodes || [])],
      lastImportedAt: store.syncMeta.lastImportedAt || "",
      lastImportedFile: store.syncMeta.lastImportedFile || "",
      lastMasterExportAt: store.syncMeta.lastMasterExportAt || "",
      lastChangesExportAt: store.syncMeta.lastChangesExportAt || "",
      lastLocalChangeAt: store.syncMeta.lastLocalChangeAt || "",
      deviceId: store.syncMeta.deviceId,
      deletedTaskCodes: [...(store.syncMeta.deletedTaskCodes || [])],
      deletedTasks: JSON.parse(JSON.stringify(store.syncMeta.deletedTasks || []))
    };
  }

  async prepareMasterPackage() {
    const store = this.readStore();
    const exportedAt = new Date().toISOString();
    return {
      fileType: "MEE_MASTER",
      schemaVersion: "1.1",
      appVersion: this.config.version,
      masterId: store.syncMeta.masterId,
      dataVersion: Number(store.syncMeta.dataVersion || 0) + 1,
      exportedAt,
      lastModifiedAt: store.syncMeta.lastLocalChangeAt || exportedAt,
      deviceId: store.syncMeta.deviceId,
      exportedBy: {
        name: store.currentUser?.name || "Usuario local",
        email: store.currentUser?.email || ""
      },
      data: {
        people: store.people || [],
        tasks: store.tasks || [],
        deletedTasks: store.syncMeta.deletedTasks || []
      }
    };
  }

  async prepareRecoveryMasterPackage() {
    const store = this.readStore();
    const exportedAt = new Date().toISOString();
    return {
      fileType: "MEE_MASTER",
      schemaVersion: "1.1",
      appVersion: this.config.version,
      masterId: store.syncMeta.masterId,
      dataVersion: Number(store.syncMeta.dataVersion || 1),
      exportedAt,
      lastModifiedAt: store.syncMeta.lastLocalChangeAt || exportedAt,
      deviceId: store.syncMeta.deviceId,
      exportedBy: {
        name: store.currentUser?.name || "Usuario local",
        email: store.currentUser?.email || ""
      },
      data: {
        people: store.people || [],
        tasks: store.tasks || [],
        deletedTasks: store.syncMeta.deletedTasks || []
      }
    };
  }

  // Alias sin efecto secundario para compatibilidad con código anterior.
  async buildMasterPackage() {
    return this.prepareMasterPackage();
  }

  async commitMasterSave(pkg) {
    if (!pkg || pkg.fileType !== "MEE_MASTER") throw new Error("Paquete maestro inválido.");
    const store = this.readStore();
    store.syncMeta.dataVersion = Number(pkg.dataVersion || store.syncMeta.dataVersion || 1);
    store.syncMeta.lastMasterExportAt = pkg.exportedAt || new Date().toISOString();
    store.syncMeta.baseExportedAt = pkg.exportedAt || store.syncMeta.baseExportedAt || "";
    store.syncMeta.lastChangesExportAt = store.syncMeta.lastMasterExportAt;
    store.syncMeta.dirtyTaskCodes = [];
    store.syncMeta.deletedTaskCodes = [];
    this.writeStore(store, { skipTracking: true });
    return this.getSyncInfo();
  }

  async buildChangesPackage() {
    const store = this.readStore();
    const dirty = new Set(store.syncMeta.dirtyTaskCodes || []);
    const tasks = (store.tasks || []).filter(task => dirty.has(task.code));
    return {
      fileType: "MEE_CHANGES",
      schemaVersion: "1.0",
      appVersion: this.config.version,
      packageId: this.createId("CHANGES"),
      baseMasterId: store.syncMeta.masterId,
      baseDataVersion: store.syncMeta.dataVersion,
      exportedAt: new Date().toISOString(),
      exportedBy: {
        name: store.currentUser?.name || "Usuario local",
        email: store.currentUser?.email || ""
      },
      deviceId: store.syncMeta.deviceId,
      tasks,
      deletedTaskCodes: [...(store.syncMeta.deletedTaskCodes || [])],
      deletedTasks: JSON.parse(JSON.stringify((store.syncMeta.deletedTasks || []).filter(item => (store.syncMeta.deletedTaskCodes || []).includes(item.code))))
    };
  }

  async markChangesExported() {
    const store = this.readStore();
    store.syncMeta.dirtyTaskCodes = [];
    store.syncMeta.deletedTaskCodes = [];
    store.syncMeta.lastChangesExportAt = new Date().toISOString();
    this.writeStore(store, { skipTracking: true });
  }

  validateMasterPackage(pkg) {
    if (pkg?.fileType === "MEE_SUBMISSION" || pkg?.submissionType === "MEE_MASTER_SUBMISSION") {
      throw new Error("El archivo seleccionado es un envío MEE_SUBMISSION, no el archivo maestro.");
    }
    if (!pkg || pkg.fileType !== "MEE_MASTER" || !pkg.data || !Array.isArray(pkg.data.tasks) || !Array.isArray(pkg.data.people)) {
      throw new Error("El archivo no es una base maestra válida del Comité MEE.");
    }
  }

  async importMasterPackage(pkg, fileName = "") {
    this.validateMasterPackage(pkg);
    const current = this.readStore();
    const people = JSON.parse(JSON.stringify(pkg.data.people));
    const tasks = JSON.parse(JSON.stringify(pkg.data.tasks));
    let currentUser = current.currentUser;
    if (!people.some(person => person.email === currentUser?.email)) {
      const preferred = people.find(person => String(person.cargo || "").toLowerCase().includes("administrador")) || people[0];
      if (preferred) {
        const role = this.roleForPerson(preferred);
        const parts = String(preferred.name || preferred.email).trim().split(/\s+/).filter(Boolean);
        currentUser = {
          name: preferred.name,
          email: preferred.email,
          role: role.role,
          roleLabel: role.roleLabel,
          cargo: preferred.cargo,
          group: preferred.group,
          initials: `${parts[0]?.[0] || "U"}${parts[parts.length - 1]?.[0] || ""}`.toUpperCase()
        };
      }
    }
    const store = {
      currentUser,
      people,
      tasks,
      syncMeta: {
        schemaVersion: "1.1",
        masterId: pkg.masterId || this.createId("MASTER"),
        dataVersion: Number(pkg.dataVersion || 1),
        baseExportedAt: pkg.exportedAt || "",
        lastImportedExportedAt: pkg.exportedAt || "",
        deviceId: this.getDeviceId(),
        dirtyTaskCodes: [],
        deletedTaskCodes: [],
        deletedTasks: JSON.parse(JSON.stringify(pkg.data.deletedTasks || pkg.deletedTasks || [])),
        lastImportedAt: new Date().toISOString(),
        lastImportedFile: fileName || "Base maestra",
        lastLocalChangeAt: ""
      }
    };
    this.ensureStoreShape(store);
    this.writeStore(store, { skipTracking: true });
    return { tasks: tasks.length, people: people.length };
  }

  async mergeChangePackages(packages) {
    const store = this.readStore();
    const byCode = new Map((store.tasks || []).map(task => [task.code, task]));
    const localDirty = new Set(store.syncMeta.dirtyTaskCodes || []);
    const mergedDirty = new Set(store.syncMeta.dirtyTaskCodes || []);
    const conflicts = [];
    let added = 0;
    let updated = 0;
    let ignored = 0;

    for (const pkg of packages) {
      if (!pkg || pkg.fileType !== "MEE_CHANGES" || !Array.isArray(pkg.tasks)) {
        ignored += 1;
        continue;
      }
      const incomingDeleted = Array.isArray(pkg.deletedTasks)
        ? pkg.deletedTasks
        : (pkg.deletedTaskCodes || []).map(code => ({ code, deletedAt: pkg.exportedAt || new Date().toISOString(), deletedBy: pkg.exportedBy?.name || "Importado" }));
      for (const deletedItem of incomingDeleted) {
        const deletedCode = deletedItem?.code;
        if (!deletedCode) continue;
        if (byCode.has(deletedCode)) {
          byCode.delete(deletedCode);
          mergedDirty.add(deletedCode);
          updated += 1;
        }
        store.syncMeta.deletedTaskCodes = [...new Set([...(store.syncMeta.deletedTaskCodes || []), deletedCode])];
        const audit = new Map((store.syncMeta.deletedTasks || []).map(item => [item.code, item]));
        const previous = audit.get(deletedCode);
        if (!previous || new Date(deletedItem.deletedAt || 0) >= new Date(previous.deletedAt || 0)) {
          audit.set(deletedCode, JSON.parse(JSON.stringify(deletedItem)));
        }
        store.syncMeta.deletedTasks = [...audit.values()];
      }
      for (const incoming of pkg.tasks) {
        if (!incoming?.code) continue;
        const local = byCode.get(incoming.code);
        const incomingAt = new Date(incoming.sync?.lastModifiedAt || pkg.exportedAt || 0).getTime();
        const localAt = new Date(local?.sync?.lastModifiedAt || local?.createdAt || 0).getTime();
        if (!local) {
          byCode.set(incoming.code, JSON.parse(JSON.stringify(incoming)));
          mergedDirty.add(incoming.code);
          added += 1;
        } else if (incomingAt > localAt) {
          if (localDirty.has(incoming.code)) {
            conflicts.push({
              code: incoming.code,
              resolution: "Se conservó la versión más reciente",
              localAt: local?.sync?.lastModifiedAt || "",
              incomingAt: incoming?.sync?.lastModifiedAt || pkg.exportedAt || ""
            });
          }
          byCode.set(incoming.code, JSON.parse(JSON.stringify(incoming)));
          mergedDirty.add(incoming.code);
          updated += 1;
        } else {
          ignored += 1;
        }
      }
    }

    store.tasks = [...byCode.values()].sort((a, b) => String(b.code).localeCompare(String(a.code)));
    store.syncMeta.dirtyTaskCodes = [...mergedDirty];
    store.syncMeta.lastConsolidatedAt = new Date().toISOString();
    this.writeStore(store, { skipTracking: true });
    return { added, updated, ignored, conflicts };
  }

  async resetDemo() {
    localStorage.removeItem(this.storageKey);
    return this.mode === "file" ? this.initializeFile() : this.initializeDemo();
  }
}

window.DataAdapter = DataAdapter;
