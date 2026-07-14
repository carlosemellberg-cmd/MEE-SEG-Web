
// Diagnóstico no destructivo para Android/WebView.
window.addEventListener("error", event => {
  try {
    const detail = {
      at: new Date().toISOString(),
      type: "error",
      message: event.message || "Error desconocido",
      source: event.filename || "",
      line: event.lineno || 0,
      column: event.colno || 0
    };
    localStorage.setItem("mee_last_diagnostic", JSON.stringify(detail));
  } catch (_) {}
});
window.addEventListener("unhandledrejection", event => {
  try {
    const detail = {
      at: new Date().toISOString(),
      type: "unhandledrejection",
      message: String(event.reason?.message || event.reason || "Promesa rechazada")
    };
    localStorage.setItem("mee_last_diagnostic", JSON.stringify(detail));
  } catch (_) {}
});

window.DEMO_DATA = {
  currentUser: {
    name: "Usuario local",
    email: "",
    role: "OPERADOR",
    roleLabel: "Operador",
    cargo: "",
    group: "GENERAL",
    initials: "UL"
  },
  people: [],
  tasks: []
};

const app = {
  adapter: null,
  auth: null,
  fileSync: null,
  people: [],
  user: null,
  microsoftConnected: false,
  graphTest: null,
  route: "home",
  selectedTask: null,
  selectedWorkItem: null,
  taskTab: "summary",
  filter: "TODAS",
  pendingSubmission: null,
  notes: null,
  selectedNoteIdForTask: "",
  noteDraftTimer: null,
  syncRefreshTimer: null,
  syncRefreshRunning: false,

  async init() {
    this.adapter = new DataAdapter(window.APP_CONFIG);
    await this.adapter.initialize();
    this.people = await this.adapter.getPeople();
    this.fileSync = new FileSyncManager(this.adapter, window.APP_CONFIG);
    await this.fileSync.initialize();
    this.notes = new QuickNotesStore();
    this.populateLocalUserSelect();
    this.updateLoginState();
    this.bindStaticEvents();
    if ("serviceWorker" in navigator) {
      window.addEventListener("load", () => {
        navigator.serviceWorker.register("./service-worker.js").catch(console.warn);
      });
    }
  },

  updateLoginState() {
    const hasMaster = Array.isArray(this.people) && this.people.length > 0;
    const loginButton = document.querySelector("#btn-demo");
    const status = document.querySelector("#login-master-status");
    if (loginButton) loginButton.disabled = !hasMaster;
    if (status) {
      const store = this.adapter?.readStore?.() || {};
      const version = Number(store.syncMeta?.dataVersion || 0);
      status.textContent = hasMaster
        ? `Maestro cargado · versión ${version} · ${this.people.length} personas. Elegí tu usuario e ingresá.`
        : "La web no contiene datos corporativos. Cargá MEE_DATOS_COMITE_MASTER.json para comenzar.";
    }
  },

  populateLocalUserSelect() {
    const select = document.querySelector("#local-user-select");
    if (!select) return;
    const currentEmail = this.adapter.readStore().currentUser?.email || "";
    select.innerHTML = this.people
      .slice()
      .sort((a, b) => String(a.name).localeCompare(String(b.name), "es"))
      .map(person => `<option value="${this.escapeHtml(person.email)}" ${person.email === currentEmail ? "selected" : ""}>${this.escapeHtml(person.name)}</option>`)
      .join("");
  },

  async completeMicrosoftConnection(result) {
    this.microsoftConnected = true;
    this.graphTest = result;
    const email = result.profile.mail || result.profile.userPrincipalName || "";
    const name = result.profile.displayName || result.account.name || "Usuario Microsoft";
    const parts = name.trim().split(/\s+/).filter(Boolean);
    const initials = (parts[0]?.[0] || "U") + (parts[parts.length - 1]?.[0] || "");
    const demoUser = await this.adapter.getCurrentUser();
    await this.enterApp({ ...demoUser, name, email, initials: initials.toUpperCase() });
    this.openModal(`<h3>Conexión Microsoft correcta</h3>
      <div class="panel"><p><strong>${name}</strong><br>${email}</p></div>
      <div class="panel"><p><strong>Microsoft Graph:</strong> conectado<br>
      <strong>Sitio Operaciones Catamarca:</strong> ${result.site && !result.site.error ? "accesible" : "no confirmado"}</p></div>
      <div class="warning">Esta versión solamente valida la autenticación. Las tareas todavía continúan en modo demostración hasta configurar las direcciones de las listas.</div>`);
  },

  bindStaticEvents() {
    document.querySelector("#btn-load-master-login").addEventListener("click", () => {
      document.querySelector("#file-master-input").click();
    });
    document.querySelector("#btn-demo").addEventListener("click", async () => {
      try {
        if (!this.people.length) throw new Error("Primero cargá MEE_DATOS_COMITE_MASTER.json.");
        const email = document.querySelector("#local-user-select").value;
        await this.adapter.setCurrentUser(email);
        await this.enterApp();
      } catch (error) {
        this.openModal(`<h3>No se pudo ingresar</h3><div class="warning">${this.escapeHtml(error.message || error)}</div>`);
      }
    });
    document.querySelector("#btn-microsoft").addEventListener("click", () => {
      this.openModal(`<h3>Microsoft pendiente de aprobación</h3><div class="warning">Por ahora usá el modo archivo. La conexión directa con Microsoft Lists se habilitará después de la aprobación administrativa.</div>`);
    });
    document.querySelector("#file-master-input").addEventListener("change", async event => {
      const file = event.target.files?.[0];
      event.target.value = "";
      if (!file) return;
      if (this.fileSync.hasPendingSubmission()) {
        this.openModal(`<h3>Sincronización pendiente</h3><div class="warning">Confirmá el envío pendiente antes de importar otro archivo maestro.</div>`);
        return;
      }
      const currentSync = await this.adapter.getSyncInfo();
      if ((this.people.length || currentSync.dirtyCount) && !confirm("Actualizar desde la base maestra reemplazará los datos locales actuales. ¿Continuar?")) return;
      await this.runFileOperation(async () => {
        const result = await this.fileSync.importMasterFile(file);
        await this.afterFileDataChange();
        return `Base actualizada: ${result.tasks} tareas y ${result.people} personas.`;
      });
    });
    document.querySelector("#file-confirm-input").addEventListener("change", async event => {
      const file = event.target.files?.[0];
      event.target.value = "";
      if (!file) return;
      await this.runFileOperation(async () => {
        const result = await this.fileSync.confirmPendingSyncWithFile(file);
        if (result.confirmed) {
          await this.afterFileDataChange();
          return `Sincronización confirmada. Versión ${result.dataVersion}.`;
        }
        if (result.status === "conflict") {
          throw new Error(result.detail || "Existe un conflicto real con el maestro remoto.");
        }
        return result.detail || "El maestro todavía no contiene el ACK esperado.";
      });
    });
    document.querySelector("#file-changes-input").addEventListener("change", async event => {
      const files = [...(event.target.files || [])];
      event.target.value = "";
      if (!files.length) return;
      await this.runFileOperation(async () => {
        const result = await this.fileSync.consolidateFiles(files);
        await this.afterFileDataChange();
        return `Consolidación terminada: ${result.added} nuevas, ${result.updated} actualizadas, ${result.ignored} omitidas, ${result.conflicts.length} conflicto(s) resuelto(s).`;
      });
    });
    document.querySelector("#btn-back").addEventListener("click", () => {
      if (this.selectedTask) {
        this.selectedTask = null;
        this.route = "tasks";
        this.render();
      }
    });
    document.querySelector("#btn-profile").addEventListener("click", () => this.go("profile"));
    document.querySelectorAll(".nav-item").forEach(btn => {
      btn.addEventListener("click", () => this.go(btn.dataset.route));
    });
  },

  async runFileOperation(operation) {
    try {
      const message = await operation();
      if (message) this.toast(message);
    } catch (error) {
      console.error(error);
      this.openModal(`<h3>No se pudo completar</h3><div class="warning">${this.escapeHtml(error.message || error)}</div>`);
    }
  },

  schedulePendingAutoRefresh() {
    if (this.syncRefreshTimer) {
      clearTimeout(this.syncRefreshTimer);
      this.syncRefreshTimer = null;
    }
    if (!this.pendingSubmission?.hasPending || !window.MeeNative?.isNative?.()) return;

    this.syncRefreshTimer = setTimeout(() => {
      this.refreshSyncState({ silent: true });
    }, 12000);
  },

  async refreshSyncState(options = {}) {
    const silent = !!options.silent;
    if (this.syncRefreshRunning) return null;

    if (!window.MeeNative?.isNative?.()) {
      if (!this.fileSync.hasPendingSubmission()) return null;
      if (!silent) document.querySelector("#file-confirm-input")?.click();
      return null;
    }

    if (!window.MeeNative?.pendingSubmissionInfo?.()?.hasPending) return null;
    this.syncRefreshRunning = true;
    try {
      const result = await this.fileSync.refreshPendingState({
        attempts: silent ? 1 : 2,
        delayMs: silent ? 0 : 1500
      });

      if (result.confirmed) {
        await this.afterFileDataChange();
        this.toast(`Sincronización confirmada. Versión ${result.dataVersion}.`);
        return result;
      }

      if (result.status === "conflict") {
        throw new Error(result.detail || "Existe un conflicto con otra versión del maestro.");
      }

      if (!silent) {
        await this.render();
        this.toast(result.detail || "El ACK todavía no aparece en el maestro.");
      } else {
        this.pendingSubmission = (await this.fileSync.getStatus()).pendingSubmission || null;
        this.schedulePendingAutoRefresh();
      }
      return result;
    } catch (error) {
      console.error(error);
      if (!silent) {
        this.openModal(`<h3>No se pudo actualizar el estado</h3><div class="warning">${this.escapeHtml(error.message || error)}</div>`);
      } else {
        this.schedulePendingAutoRefresh();
      }
      return null;
    } finally {
      this.syncRefreshRunning = false;
    }
  },

  handleNativeResume() {
    if (document.querySelector("#main-shell")?.hidden) return;
    if (!window.MeeNative?.pendingSubmissionInfo?.()?.hasPending) return;
    setTimeout(() => this.refreshSyncState({ silent: true }), 1200);
  },

  ensureSyncDestinationEmail() {
    let current = this.fileSync.getSyncDestinationEmail() || window.APP_CONFIG.syncDestinationEmail || "";
    while (!current) {
      const entered = prompt(
        "Ingresá el correo corporativo que recibirá los envíos de MEE SEG para Power Automate:",
        ""
      );
      if (entered === null) return null;
      const value = String(entered || "").trim();
      if (!/^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$/.test(value)) {
        alert("Ingresá una dirección de correo válida.");
        continue;
      }
      if (!this.fileSync.setSyncDestinationEmail(value)) {
        alert("No se pudo guardar el correo. Revisá el formato.");
        continue;
      }
      current = value;
    }
    return current;
  },

  applyPendingUiLock() {
    if (!this.pendingSubmission?.hasPending) return;
    const mutationActions = new Set([
      "new-task", "delete-task", "add-contribution", "edit-task-text",
      "add-participant", "remove-participant", "add-work-item", "edit-work-item",
      "delete-work-item", "add-work-progress", "manage-work-participants",
      "add-pending", "edit-pending", "delete-pending", "toggle-pending",
      "close-task", "reset-demo"
    ]);
    document.querySelectorAll("[data-action]").forEach(element => {
      if (mutationActions.has(element.dataset.action)) {
        element.disabled = true;
        element.title = "Existe una sincronización pendiente. Confirmala antes de realizar nuevos cambios.";
      }
    });
    const form = document.querySelector("#new-task-form");
    if (form) {
      form.querySelectorAll("input, textarea, select, button[type='submit']").forEach(element => { element.disabled = true; });
    }
  },

  async afterFileDataChange() {
    this.people = await this.adapter.getPeople();
    this.user = await this.adapter.getCurrentUser();
    this.populateLocalUserSelect();
    this.updateLoginState();
    if (!document.querySelector("#main-shell").hidden) {
      document.querySelector("#btn-profile").textContent = this.user.initials;
      await this.render();
    }
  },

  async enterApp(userOverride = null) {
    this.people = await this.adapter.getPeople();
    this.user = userOverride || await this.adapter.getCurrentUser();
    document.querySelector("#login-view").hidden = true;
    document.querySelector("#main-shell").hidden = false;
    document.querySelector("#btn-profile").textContent = this.user.initials;
    await this.render();
  },

  async go(route) {
    this.route = route;
    this.selectedTask = null;
    if (route !== "newTask") this.selectedNoteIdForTask = "";
    await this.render();
  },

  statusClass(status) {
    return {
      EN_ANALISIS:"status-analysis",
      PRESENTADA:"status-analysis",
      EN_EJECUCION:"status-progress",
      PLANIFICADA:"status-planned",
      FINALIZADA:"status-done",
      PAUSADA:"status-paused"
    }[status] || "status-analysis";
  },

  groupLabel(value) {
    return {
      GRUPO_1: "Grupo 1",
      GRUPO_2: "Grupo 2",
      GRUPO_3: "Grupo 3",
      GRUPO_4: "Grupo 4",
      DOS_TURNOS: "Dos turnos",
      GENERAL: "General",
      CENTRAL: "Central"
    }[value] || String(value || "").replaceAll("_", " ");
  },

  typeLabel(value) {
    return {
      SEGURIDAD_ELECTRICA: "Seguridad eléctrica",
      MEJORA_TECNICA: "Mejora técnica",
      PROCEDIMIENTO: "Procedimiento",
      DOCUMENTACION: "Documentación",
      OTRO: "Otro"
    }[value] || value;
  },

  priorityLabel(value) {
    return {
      BAJA: "Baja",
      MEDIA: "Media",
      ALTA: "Alta",
      CRITICA: "Crítica"
    }[value] || value;
  },

  person(email) {
    return (this.people || []).find(p => p.email === email) || {email, name: email || "Sin asignar", group:"", cargo:""};
  },

  personName(email) {
    return this.person(email).name;
  },

  canManage() {
    return ["ADMINISTRADOR","JEFE"].includes(this.user.role);
  },

  isAdministrator() {
    return this.user.role === "ADMINISTRADOR";
  },

  canUpdateWorkItem(item) {
    return this.canManage();
  },

  escapeHtml(value) {
    return String(value ?? "").replace(/[&<>'"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[c]));
  },

  workStatusLabel(status) {
    return {PLANIFICADA:"Planificada",EN_EJECUCION:"En ejecución",COMPLETADA:"Completada",PAUSADA:"Pausada"}[status] || status;
  },

  async render() {
    const root = document.querySelector("#view-root");
    const back = document.querySelector("#btn-back");
    back.hidden = !this.selectedTask;
    const title = this.selectedTask ? this.selectedTask.code : ({
      home:"Inicio", tasks:"Tareas", pending:"Pendientes", notes:"Apuntes", documents:"Documentos", profile:"Perfil", fileSync:"Centro de archivos", newTask:"Nueva tarea"
    }[this.route] || "Inicio");
    document.querySelector("#topbar-subtitle").textContent = title;

    document.querySelectorAll(".nav-item").forEach(b => b.classList.toggle("active", b.dataset.route === this.route));

    const syncStatus = await this.fileSync.getStatus();
    this.pendingSubmission = syncStatus.pendingSubmission || null;

    let content = "";
    if (this.selectedTask) content = await this.renderTaskDetail(this.selectedTask);
    else if (this.route === "home") content = await this.renderHome();
    else if (this.route === "tasks") content = await this.renderTasks();
    else if (this.route === "pending") content = await this.renderPending();
    else if (this.route === "notes") content = this.renderNotes();
    else if (this.route === "documents") content = this.renderDocuments();
    else if (this.route === "profile") content = await this.renderProfile();
    else if (this.route === "fileSync") content = await this.renderFileSync();
    else if (this.route === "newTask") content = this.renderNewTask();

    if (this.pendingSubmission?.hasPending && this.route !== "fileSync") {
      content = `<div class="warning"><strong>Sincronización pendiente.</strong><br>
        Confirmala antes de realizar nuevos cambios. Podés navegar y consultar información, pero la edición está temporalmente bloqueada.
        <div class="action-row"><button class="btn btn-primary" data-route-action="fileSync">Abrir sincronización</button></div>
      </div>${content}`;
    }

    root.innerHTML = content;
    this.bindViewEvents();
    this.applyPendingUiLock();
    this.schedulePendingAutoRefresh();
  },

  async renderHome() {
    const tasks = await this.adapter.getTasks();
    const active = tasks.filter(t => t.status !== "FINALIZADA");
    const pending = active.reduce((n,t) => n + t.pending.filter(p => p.status !== "COMPLETADO").length, 0);
    const late = active.filter(t => t.status === "PAUSADA").length;
    const quickNotes = this.notes?.listPending(this.user).length || 0;
    return `
      <section class="hero">
        <div class="hero-top">
          <div>
            <h2>${this.escapeHtml(this.user.name)}</h2>
          </div>
        </div>
        <div class="kpi-grid">
          <div class="kpi"><strong>${active.length}</strong><small>Tareas activas</small></div>
          <div class="kpi"><strong>${pending}</strong><small>Pendientes</small></div>
          <div class="kpi"><strong>${quickNotes}</strong><small>Apuntes</small></div>
          <div class="kpi"><strong>${late}</strong><small>Bloqueadas</small></div>
        </div>
        <div class="connection-badge"><span class="dot file-dot"></span>Modo archivo · datos guardados localmente en este teléfono</div>
      </section>

      <div class="section-head"><h3>Acciones rápidas</h3></div>
      <section class="quick-grid">
        ${this.canManage() ? `<button class="quick-card primary" data-action="new-task"><span class="q-icon">＋</span><strong>Nueva tarea</strong><small>Registrar problema o mejora</small></button>` : ""}
        <button class="quick-card" data-route-action="notes"><span class="q-icon">✎</span><strong>Apunte rápido</strong><small>${this.canManage() ? "Anotar ahora y ordenar después" : "Consultar apuntes guardados"}</small></button>
        <button class="quick-card" data-route-action="tasks"><span class="q-icon">☷</span><strong>Ver tareas</strong><small>Seguimiento del comité</small></button>
        <button class="quick-card" data-route-action="pending"><span class="q-icon">✓</span><strong>Pendientes</strong><small>Acciones abiertas</small></button>
        <button class="quick-card" data-route-action="documents"><span class="q-icon">▤</span><strong>Documentos</strong><small>Archivos y evidencias</small></button>
        <button class="quick-card" data-route-action="fileSync"><span class="q-icon">⇄</span><strong>Archivo maestro</strong><small>Actualizar o guardar la base JSON</small></button>
      </section>

      <div class="section-head"><h3>Actividad reciente</h3><button class="text-btn" data-route-action="tasks">Ver todas</button></div>
      <section class="task-list-grid">
        ${active.slice(0,3).map(t => this.taskCard(t)).join("")}
      </section>

    `;
  },

  taskCard(t) {
    return `<article class="task-card" data-task="${t.code}">
      <div class="task-card-head">
        <div><div class="task-code">${t.code}</div><h4>${t.title}</h4></div>
        <span class="status-chip ${this.statusClass(t.status)}">${t.statusLabel}</span>
      </div>
      <div class="task-meta"><span>Responsable: ${t.responsible}</span><span>Ubicación: ${t.location}</span></div>
      <div class="progress"><span style="width:${t.progress}%"></span></div>
    </article>`;
  },

  async renderTasks() {
    const tasks = await this.adapter.getTasks();
    const filters = ["TODAS","EN_ANALISIS","PLANIFICADA","EN_EJECUCION","FINALIZADA"];
    const filtered = this.filter === "TODAS" ? tasks : tasks.filter(t => t.status === this.filter);
    return `
      <h2 class="page-title">Tareas del comité</h2>
      <input id="task-search" class="search-box" type="search" placeholder="Buscar por título, código o ubicación">
      <div class="filters">${filters.map(f => `<button class="filter-btn ${this.filter===f?"active":""}" data-filter="${f}">${({
        TODAS:"Todas",EN_ANALISIS:"En análisis",PLANIFICADA:"Planificadas",EN_EJECUCION:"En ejecución",FINALIZADA:"Finalizadas"
      })[f]}</button>`).join("")}</div>
      ${this.canManage() ? `<button class="btn btn-primary btn-full" data-action="new-task">＋ Nueva tarea</button><div style="height:12px"></div>` : ""}
      <section id="task-results" class="task-list-grid">
        ${filtered.length ? filtered.map(t => this.taskCard(t)).join("") : `<div class="empty">No hay tareas en este filtro.</div>`}
      </section>
    `;
  },

  async renderPending() {
    const tasks = await this.adapter.getTasks();
    const rows = tasks.flatMap(t => t.pending.filter(p => p.status !== "COMPLETADO").map(p => ({...p,task:t})));
    return `<h2 class="page-title">Pendientes abiertos</h2>
      ${rows.length ? rows.map(r => `<article class="panel" data-task="${r.task.code}">
        <div class="task-code">${r.task.code}</div>
        <h3>${r.description}</h3>
        <p>Responsable: ${r.responsible}</p>
      </article>`).join("") : `<div class="empty">No hay pendientes abiertos.</div>`}`;
  },

  renderNotes() {
    const notes = this.notes?.listForUser(this.user) || [];
    const pending = notes.filter(note => note.status !== "CONVERTIDO");
    const converted = notes.filter(note => note.status === "CONVERTIDO");
    const draft = this.notes?.getDraft(this.user) || { text: "", updatedAt: "" };
    const canEdit = this.canManage();

    const card = note => `<article class="note-card ${note.status === "CONVERTIDO" ? "converted" : ""}">
      <div class="note-card-head">
        <div>
          <strong>${note.status === "CONVERTIDO" ? "Apunte convertido" : "Apunte pendiente"}</strong>
          <div class="note-meta"><span>${this.date(note.updatedAt || note.createdAt)}</span><span>${this.escapeHtml(note.authorName || this.user.name)}</span></div>
        </div>
        <span class="note-status ${note.status === "CONVERTIDO" ? "converted" : ""}">${note.status === "CONVERTIDO" ? "CONVERTIDO" : "PENDIENTE"}</span>
      </div>
      <p>${this.escapeHtml(note.text)}</p>
      ${note.convertedTaskCode ? `<div class="note-meta"><span>Tarea: ${this.escapeHtml(note.convertedTaskCode)}</span></div>` : ""}
      ${canEdit ? `<div class="note-actions">
        ${note.status !== "CONVERTIDO" ? `<button class="btn btn-primary" data-action="create-task-from-note" data-note-id="${note.id}">Crear tarea con este apunte</button>` : `<button class="btn btn-secondary" data-action="reopen-note" data-note-id="${note.id}">Volver a pendiente</button>`}
        <button class="btn btn-secondary" data-action="edit-note" data-note-id="${note.id}">Editar</button>
        <button class="btn btn-danger" data-action="delete-note" data-note-id="${note.id}">Eliminar</button>
      </div>` : ""}
    </article>`;

    return `<h2 class="page-title">Apuntes rápidos</h2>
      ${canEdit ? `<section class="form-card notes-compose">
        <div>
          <h3>Escribí mientras se conversa</h3>
          <p class="panel-note">El borrador se guarda automáticamente en este dispositivo. Después podés convertirlo en una tarea sin salir de la aplicación.</p>
        </div>
        <div class="field">
          <label>Apunte en curso</label>
          <textarea id="quick-note-draft" placeholder="Anotá problemas, ideas, decisiones, sectores involucrados o cualquier dato para ordenar después.">${this.escapeHtml(draft.text || "")}</textarea>
        </div>
        <div class="autosave-line">
          <span><span class="autosave-dot"></span><span id="note-autosave-status">${draft.updatedAt ? `Guardado automáticamente ${this.date(draft.updatedAt)}` : "Borrador listo"}</span></span>
          <span>Solo local</span>
        </div>
        <div class="action-row">
          <button class="btn btn-primary" data-action="save-quick-note">Guardar como apunte pendiente</button>
          <button class="btn btn-secondary" data-action="clear-note-draft">Limpiar borrador</button>
        </div>
      </section>` : `<div class="warning">El rol Operador es de solo lectura. Puede consultar apuntes, pero no crearlos, modificarlos ni eliminarlos.</div>`}

      <div class="section-head"><h3>Pendientes de ordenar</h3><span>${pending.length}</span></div>
      <section class="note-list">
        ${pending.length ? pending.map(card).join("") : `<div class="empty">No hay apuntes pendientes.</div>`}
      </section>

      ${converted.length ? `<div class="section-head"><h3>Convertidos en tarea</h3><span>${converted.length}</span></div>
      <section class="note-list">${converted.map(card).join("")}</section>` : ""}

      <div class="location-note"><strong>Importante:</strong> los apuntes rápidos quedan guardados únicamente en este dispositivo y no se envían al maestro hasta crear una tarea formal.</div>`;
  },

  renderDocuments() {
    return `<h2 class="page-title">Documentos y reportes</h2>
      <section class="panel">
        <h3>Archivo maestro de datos</h3>
        <p>La aplicación trabaja siempre con <strong>MEE_DATOS_COMITE_MASTER.json</strong>. Para cargar la última versión, ingresá en Actualizar datos y elegí el archivo desde OneDrive, Descargas o Archivos del dispositivo.</p>
        <div class="action-row"><button class="btn btn-primary" data-route-action="fileSync">Actualizar datos</button></div>
      </section>
      <section class="panel">
        <h3>Minutas y reportes</h3>
        <p>Podés generar un reporte general en Excel o PDF. Dentro de cada tarea también está disponible la minuta PDF completa con problema, propuesta, participantes, aportes, plan de acción, avances y pendientes.</p>
        <div class="action-row"><button class="btn btn-secondary" data-action="generate-report">Reporte Excel</button><button class="btn btn-primary" data-action="generate-pdf-report">Reporte PDF</button></div>
      </section>`;
  },

  async renderProfile() {
    const sync = await this.fileSync.getStatus();
    return `<h2 class="page-title">Perfil</h2>
      <section class="panel">
        <div class="hero-top">
          <div><h3>${this.user.name}</h3><p>${this.user.email}</p></div>
          <span class="role-chip">${this.user.roleLabel}</span>
        </div>
        <div class="detail-grid">
          <div class="detail-item"><small>Cargo</small><strong>${this.user.cargo}</strong></div>
          <div class="detail-item"><small>Grupo</small><strong>${this.groupLabel(this.user.group)}</strong></div>
        </div>
        <div class="action-row"><button class="btn btn-secondary" data-action="change-local-user">Cambiar usuario local</button></div>
      </section>
      <section class="panel">
        <h3>Modo de trabajo</h3>
        <p>Versión ${APP_CONFIG.version}<br>Conexión: modo archivo<br>Cambios locales pendientes de exportar: ${sync.dirtyCount}</p>
        <div class="action-row"><button class="btn btn-primary" data-route-action="fileSync">Abrir centro de archivos</button>${this.canManage() ? `<button class="btn btn-secondary" data-action="reset-demo">Restablecer datos iniciales</button>` : ""}</div>
      </section>`;
  },



  async renderFileSync() {
    const sync = await this.fileSync.getStatus();
    const formatDate = value => value ? this.date(value) : "Nunca";
    const isAndroid = this.fileSync.isAndroidApp();
    const directFile = !isAndroid && !!window.showOpenFilePicker;

    if (isAndroid) {
      const remembered = !!sync.linkedFileName;
      const pending = sync.pendingSubmission;
      const email = sync.syncDestinationEmail || "Sin configurar";
      const pendingMode = pending?.submissionMode === "RECOVERY" ? "Recuperación del maestro" : "Actualización normal";
      const pendingId = pending?.submissionId ? `${pending.submissionId.slice(0, 8)}…${pending.submissionId.slice(-4)}` : "";

      return `<h2 class="page-title">Archivo maestro</h2>
        <section class="master-status-grid">
          <div class="master-status-card wide"><small>Archivo recordado</small><strong>${this.escapeHtml(sync.linkedFileName || "Todavía no se eligió")}</strong></div>
          <div class="master-status-card"><small>Acceso Android</small><strong>${remembered ? "Solo lectura segura" : "Sin archivo"}</strong></div>
          <div class="master-status-card"><small>Versión local</small><strong>${sync.dataVersion}</strong></div>
          <div class="master-status-card"><small>Cambios pendientes</small><strong>${sync.dirtyCount}</strong></div>
          <div class="master-status-card"><small>Última lectura</small><strong>${this.escapeHtml(formatDate(sync.lastReadAt || sync.lastImportedAt))}</strong></div>
          <div class="master-status-card"><small>Última confirmación servidor</small><strong>${this.escapeHtml(formatDate(sync.lastSavedAt))}</strong></div>
        </section>

        ${pending ? `<section class="panel">
          <h3>Sincronización pendiente</h3>
          <p><strong>${this.escapeHtml(pendingMode)}</strong><br>
          Versión objetivo: ${pending.targetDataVersion}<br>
          Identificador: ${this.escapeHtml(pendingId)}<br>
          Archivo: ${this.escapeHtml(pending.fileName || "")}</p>
          <div class="warning">La edición está bloqueada hasta que Power Automate actualice SharePoint y la aplicación confirme el ACK.</div>
          <div class="action-row">
            <button class="btn btn-primary" data-action="refresh-sync-state">Actualizar estado</button>
            <button class="btn btn-secondary" data-action="share-pending-outlook">Reenviar por Outlook</button>
          </div>
          <p class="sync-refresh-note">La aplicación vuelve a leer el maestro automáticamente al regresar desde Outlook y mientras este envío siga pendiente.</p>
          <div class="action-row">
            ${sync.hasSubmissionFolder
              ? '<button class="btn btn-primary" data-action="send-pending-to-folder">Enviar automáticamente a carpeta</button>'
              : '<button class="btn btn-secondary" data-action="configure-submission-folder">Autorizar carpeta de entrada</button>'}
            <button class="btn btn-secondary" data-action="save-pending-copy">Guardar copia local</button>
          </div>
          <div class="action-row">
            <button class="btn btn-secondary" data-action="confirm-sync-picked">Seleccionar maestro actualizado para confirmar</button>
          </div>
        </section>` : `<section class="panel master-actions-panel">
          <button class="btn btn-primary master-main-btn" data-action="${remembered ? "reload-master" : "choose-master-file"}">
            ${remembered ? "Actualizar desde el archivo recordado" : "Traer archivo maestro"}
          </button>
          <button class="btn btn-primary master-main-btn" data-action="save-master" ${remembered ? "" : "disabled"}>Preparar envío seguro</button>
          ${remembered ? '<button class="btn btn-secondary master-secondary-btn" data-action="choose-master-file">Elegir otro archivo</button>' : ""}
        </section>`}

        <section class="panel">
          <h3>Configuración de sincronización</h3>
          <p>Correo destino de Power Automate:<br><strong>${this.escapeHtml(email)}</strong></p>
          <p>Carpeta de entrada autorizada:<br><strong>${this.escapeHtml(sync.submissionFolderName || "Sin autorizar")}</strong></p>
          <div class="action-row">
            <button class="btn btn-secondary" data-action="configure-sync-email">Configurar correo</button>
            <button class="btn btn-secondary" data-action="configure-submission-folder">${sync.hasSubmissionFolder ? "Cambiar carpeta de entrada" : "Autorizar carpeta de entrada"}</button>
          </div>
        </section>

        <div class="location-note"><strong>Flujo seguro:</strong> Android solo lee el maestro. Al guardar, crea un archivo MEE_SUBMISSION nuevo, lo abre en Outlook y mantiene los cambios pendientes. Power Automate actualiza SharePoint. La aplicación intenta confirmar automáticamente el ACK; también podés usar <strong>Actualizar estado</strong>.</div>
        ${remembered ? `<div class="location-note"><strong>Diagnóstico:</strong> proveedor ${this.escapeHtml(sync.diagUriAuthority || "sin datos")} · MIME ${this.escapeHtml(sync.diagMimeType || "sin datos")} · última lectura ${sync.diagLastReadSize ? `${this.escapeHtml(String(sync.diagLastReadSize))} bytes (${this.escapeHtml(formatDate(sync.diagLastReadAt))})` : "sin datos"} · estado: ${this.escapeHtml(sync.diagLastWriteResult || "sin operaciones")}</div>` : ""}`;
    }

    const pending = sync.pendingSubmission;
    const pendingId = pending?.submissionId
      ? `${pending.submissionId.slice(0, 8)}…${pending.submissionId.slice(-4)}`
      : "";

    return `<h2 class="page-title">Archivo maestro</h2>
      <section class="sync-status-grid">
        <div class="sync-status-card"><small>Archivo cargado</small><strong>${this.escapeHtml(sync.linkedFileName || "Sin cargar")}</strong></div>
        <div class="sync-status-card"><small>Versión local</small><strong>${sync.dataVersion}</strong></div>
        <div class="sync-status-card"><small>Cambios pendientes</small><strong>${sync.dirtyCount}</strong></div>
        <div class="sync-status-card"><small>Última actualización</small><strong>${this.escapeHtml(formatDate(sync.lastImportedAt))}</strong></div>
      </section>

      ${pending ? `<section class="panel">
        <h3>Sincronización pendiente</h3>
        <p>Versión objetivo: <strong>${pending.targetDataVersion}</strong><br>
        Identificador: <strong>${this.escapeHtml(pendingId)}</strong><br>
        Archivo: <strong>${this.escapeHtml(pending.fileName || "")}</strong></p>
        <div class="warning">El submission se conserva en este navegador hasta confirmar el ACK. Reenviar mantiene el mismo UUID.</div>
        <div class="action-row">
          <button class="btn btn-primary" data-action="refresh-sync-state">Actualizar estado</button>
          <button class="btn btn-secondary" data-action="share-pending-outlook">Descargar y abrir Outlook</button>
        </div>
        <div class="action-row">
          <button class="btn btn-secondary" data-action="save-pending-copy">Guardar otra copia</button>
          <button class="btn btn-secondary" data-action="confirm-sync-picked">Seleccionar maestro actualizado</button>
        </div>
        <p class="sync-refresh-note">En la web, Actualizar estado abre el selector para cargar el maestro actualizado descargado de SharePoint.</p>
      </section>` : `<section class="panel sync-workflow">
        <h3>1. Cargar la última base</h3>
        <p>Seleccioná exactamente <strong>MEE_DATOS_COMITE_MASTER.json</strong>.</p>
        <div class="action-row"><button class="btn btn-primary" data-action="import-master-file">Cargar archivo maestro</button></div>
      </section>
      <section class="panel sync-workflow">
        <h3>2. Preparar el envío</h3>
        <p>La web genera un <strong>MEE_SUBMISSION</strong>. No sobrescribe el maestro ni publica datos en GitHub.</p>
        <div class="action-row"><button class="btn btn-primary" data-action="save-master" ${sync.dirtyCount ? "" : "disabled"}>Preparar envío seguro</button></div>
      </section>`}

      <section class="panel">
        <h3>Correo de Power Automate</h3>
        <p><strong>${this.escapeHtml(sync.syncDestinationEmail || window.APP_CONFIG.syncDestinationEmail || "Sin configurar")}</strong></p>
        <div class="action-row"><button class="btn btn-secondary" data-action="configure-sync-email">Configurar correo</button></div>
      </section>

      <div class="location-note"><strong>Versión web v1.10.1:</strong> los apuntes quedan solo en este navegador hasta convertirlos en tarea. El maestro y los datos corporativos no se alojan en GitHub Pages. El envío por Outlook continúa siendo manual hasta completar el nuevo automatismo.</div>`;
  },

  renderNewTask() {
    const notes = this.notes?.listPending(this.user) || [];
    const selectedId = notes.some(note => note.id === this.selectedNoteIdForTask)
      ? this.selectedNoteIdForTask
      : "";
    const selectedNote = selectedId ? this.notes.get(this.user, selectedId) : null;

    return `<h2 class="page-title">Nueva tarea</h2>
      <form id="new-task-form">
        <section class="form-card note-import-card">
          <div>
            <h3>Traer desde apunte</h3>
            <p class="panel-note">Elegí un apunte pendiente y cargalo directamente en “Problema o necesidad” sin salir de esta pantalla.</p>
          </div>
          <div class="field">
            <label>Apunte pendiente</label>
            <select id="note-source-select" name="sourceNoteId">
              <option value="">No usar apunte</option>
              ${notes.map(note => `<option value="${note.id}" ${note.id === selectedId ? "selected" : ""}>${this.escapeHtml(note.text.slice(0, 80))}${note.text.length > 80 ? "…" : ""}</option>`).join("")}
            </select>
          </div>
          <div id="note-source-preview" class="note-preview">${selectedNote ? this.escapeHtml(selectedNote.text) : "Seleccioná un apunte para ver el contenido."}</div>
          <div class="action-row">
            <button type="button" class="btn btn-primary" data-action="load-note-into-task" ${notes.length ? "" : "disabled"}>Cargar apunte en la tarea</button>
            <button type="button" class="btn btn-secondary" data-route-action="notes">Abrir apuntes</button>
          </div>
        </section>
        <section class="form-card">
          <div class="field"><label>Título *</label><input name="title" required placeholder="Ej.: Acondicionar tomas eléctricas"></div>
          <div class="field"><label>Problema o necesidad *</label><textarea name="problem" required placeholder="Describí qué se detectó y por qué debe tratarse">${selectedNote ? this.escapeHtml(selectedNote.text) : ""}</textarea></div>
          <div class="field"><label>Ubicación</label><input name="location" placeholder="Sala, equipo o sector"></div>
          <div class="field"><label>Tipo</label><select name="type"><option value="SEGURIDAD_ELECTRICA">Seguridad eléctrica</option><option value="MEJORA_TECNICA">Mejora técnica</option><option value="PROCEDIMIENTO">Procedimiento</option><option value="DOCUMENTACION">Documentación</option><option value="OTRO">Otro</option></select></div>
          <div class="field"><label>Prioridad</label><select name="priority"><option value="MEDIA">Media</option><option value="ALTA">Alta</option><option value="CRITICA">Crítica</option><option value="BAJA">Baja</option></select></div>
        </section>
        <section class="form-card">
          <div class="field"><label>Propuesta inicial</label><textarea name="proposal" placeholder="Qué solución se propone inicialmente"></textarea></div>
          <div class="field"><label>Grupo responsable</label><select name="group"><option value="GRUPO_1">Grupo 1</option><option value="GRUPO_2">Grupo 2</option><option value="GRUPO_3" selected>Grupo 3</option><option value="GRUPO_4">Grupo 4</option><option value="DOS_TURNOS">Dos turnos</option><option value="GENERAL">General</option></select></div>
          <div class="field"><label>Responsable</label><input name="responsible" placeholder="Nombre o grupo"></div>
        </section>
        <div class="form-actions"><button type="button" class="btn btn-secondary" data-route-action="home">Cancelar</button><button class="btn btn-primary" type="submit">Crear tarea</button></div>
      </form>`;
  },

  async renderTaskDetail(task) {
    const tabs = [
      ["summary","Resumen"],["actionPlan","Plan de acción"],["contributions","Aportes"],["pending","Pendientes"],
      ["participants","Participantes"],["documents","Documentos"],["history","Historial"]
    ];
    return `<section class="task-detail-head">
      <div class="task-code">${task.code}</div>
      <h2>${task.title}</h2>
      <span class="status-chip ${this.statusClass(task.status)}">${task.statusLabel}</span>
      <div class="progress"><span style="width:${task.progress}%"></span></div>
      <div class="detail-grid">
        <div class="detail-item"><small>Responsable</small><strong>${task.responsible}</strong></div>
        <div class="detail-item"><small>Ubicación</small><strong>${task.location}</strong></div>
        <div class="detail-item"><small>Prioridad</small><strong>${this.priorityLabel(task.priority)}</strong></div>
        <div class="detail-item"><small>Avance</small><strong>${task.progress} %</strong></div>
      </div>
    </section>
    <div class="tabs">${tabs.map(([id,label]) => `<button class="tab ${this.taskTab===id?"active":""}" data-tab="${id}">${label}</button>`).join("")}</div>
    ${this.renderTaskTab(task)}
    <div class="action-row">
      <button class="btn btn-secondary" data-action="minute">Vista previa</button>
      <button class="btn btn-primary" data-action="task-pdf">Descargar minuta PDF</button>
      ${task.status !== "FINALIZADA" && this.canManage() ? `<button class="btn btn-primary" data-action="close-task">Finalizar</button>` : ""}
      ${this.canManage() ? `<button class="btn btn-danger" data-action="delete-task">Eliminar tarea</button>` : ""}
    </div>
    `;
  },

  renderTaskTab(task) {
    if (this.taskTab === "summary") return `
      <section class="panel editable-panel">
        <div class="panel-title-row"><h3>Problema detectado</h3>${this.canManage() ? `<button class="panel-edit-btn" data-action="edit-task-text" data-field="problem">Editar problema</button>` : ""}</div>
        <p>${task.problem}</p>
      </section>
      <section class="panel editable-panel">
        <div class="panel-title-row"><h3>Propuesta general</h3>${this.canManage() ? `<button class="panel-edit-btn" data-action="edit-task-text" data-field="proposal">Editar propuesta</button>` : ""}</div>
        <p>${task.proposal || "Sin propuesta general."}</p>
      </section>
      <section class="panel plan-summary">
        <div class="panel-title-row"><h3>Plan de acción por sector</h3><button class="panel-edit-btn" data-tab-jump="actionPlan">Ver plan completo</button></div>
        <div class="plan-summary-grid">
          <div><strong>${(task.workItems || []).length}</strong><small>Tareas por sector</small></div>
          <div><strong>${(task.workItems || []).filter(w => w.progress === 100).length}</strong><small>Completadas</small></div>
          <div><strong>${task.progress} %</strong><small>Avance general</small></div>
        </div>
      </section>
      ${this.canManage() ? `<section class="panel"><h3>¿Querés sumar información sin reemplazar lo anterior?</h3><p class="panel-note">Usá un aporte para agregar una mejora, observación, decisión o avance. El problema original queda conservado.</p><div class="action-row"><button class="btn btn-primary" data-action="add-contribution">Agregar aporte o mejora</button></div></section>` : ""}`;
    if (this.taskTab === "actionPlan") return `
      <section class="panel">
        <div class="panel-title-row"><div><h3>Plan de acción</h3><p class="panel-note">La tarea general se divide por sector o frente de trabajo. Cada una tiene responsable, participantes y avances propios.</p></div>${this.canManage() ? `<button class="panel-edit-btn" data-action="add-work-item">＋ Nueva tarea por sector</button>` : ""}</div>
        <div class="work-grid">
          ${(task.workItems || []).length ? task.workItems.map(w => this.renderWorkItemCard(w)).join("") : `<div class="empty">Todavía no hay tareas por sector.</div>`}
        </div>
      </section>`;
    if (this.taskTab === "contributions") return `
      <section class="panel"><h3>Aportes y mejoras</h3>
        <p class="panel-note">Acá se cargan propuestas, mejoras, observaciones, decisiones y avances relacionados con esta tarea.</p>
        ${task.contributions.length ? task.contributions.map(a => `<div class="list-row"><div class="list-icon">✎</div><div class="list-main"><strong>${a.type} · ${a.author}</strong><small>${this.date(a.at)} · ${this.groupLabel(a.group)}</small><p>${a.text}</p></div></div>`).join("") : `<p>No hay aportes cargados.</p>`}
        ${this.canManage() ? `<div class="action-row"><button class="btn btn-primary" data-action="add-contribution">Agregar aporte o mejora</button></div>` : ""}
      </section>`;
    if (this.taskTab === "pending") return `
      <section class="panel"><h3>Pendientes</h3>
        <p class="panel-note">Cada pendiente se puede editar, completar, reabrir o eliminar.</p>
        ${task.pending.length ? task.pending.map(p => `<div class="list-row"><div class="list-icon">${p.status==="COMPLETADO"?"✓":"!"}</div><div class="list-main"><strong>${p.description}</strong><small>${p.responsible} · ${p.status === "COMPLETADO" ? "Completado" : "Abierto"}</small></div>${this.canManage() ? `<div class="list-actions"><button class="mini-btn" data-action="toggle-pending" data-pending-id="${p.id}">${p.status === "COMPLETADO" ? "Reabrir" : "Completar"}</button><button class="mini-btn" data-action="edit-pending" data-pending-id="${p.id}">Editar</button>${this.canManage() ? `<button class="mini-btn danger" data-action="delete-pending" data-pending-id="${p.id}">Eliminar</button>` : ""}</div>` : ""}</div>`).join("") : `<p>No hay pendientes.</p>`}
        ${this.canManage() ? `<div class="action-row"><button class="btn btn-primary" data-action="add-pending">Agregar pendiente</button></div>` : ""}
      </section>`;
    if (this.taskTab === "participants") return `<section class="panel">
      <div class="panel-title-row"><div><h3>Participantes generales</h3><p class="panel-note">Reciben seguimiento y minuta de la tarea general. Los ejecutores específicos se asignan dentro de cada tarea por sector.</p></div>${this.canManage() ? `<button class="panel-edit-btn" data-action="add-participant">＋ Agregar participante</button>` : ""}</div>
      ${(task.participants || []).length ? task.participants.map(email => { const p=this.person(email); return `<div class="person-row"><div class="person-avatar">${p.name.split(" ").map(x=>x[0]).slice(0,2).join("")}</div><div class="person-info"><strong>${this.escapeHtml(p.name)}</strong><small>${this.escapeHtml(p.cargo)} · ${this.groupLabel(p.group)}</small></div>${this.canManage() ? `<button class="mini-btn danger" data-action="remove-participant" data-email="${p.email}">Eliminar</button>` : ""}</div>`; }).join("") : `<p>Sin participantes generales.</p>`}
    </section>`;
    if (this.taskTab === "documents") return `<section class="panel"><h3>Documentos</h3>${task.documents.map(d => `<div class="list-row"><div class="list-icon">PDF</div><div class="list-main"><strong>${d.name}</strong><small>${d.type}</small></div></div>`).join("") || "<p>Sin documentos vinculados.</p>"}<div class="action-row"><button class="btn btn-primary" data-action="open-documents">Abrir carpeta</button></div></section>`;
    return `<section class="panel"><h3>Historial</h3>${task.history.map(h => `<div class="list-row"><div class="list-icon">↺</div><div class="list-main"><strong>${h.text}</strong><small>${h.by} · ${this.date(h.at)}</small></div></div>`).join("")}</section>`;
  },

  renderWorkItemCard(item) {
    const responsible = this.person(item.responsibleEmail);
    const latest = (item.updates || []).slice().sort((a,b)=>new Date(b.at)-new Date(a.at))[0];
    return `<article class="work-card">
      <div class="work-card-head">
        <div><span class="work-sector">${this.escapeHtml(item.sector)}</span><h4>${this.escapeHtml(item.title)}</h4></div>
        <span class="status-chip ${item.progress === 100 ? "status-done" : item.progress > 0 ? "status-progress" : "status-planned"}">${this.workStatusLabel(item.status)}</span>
      </div>
      <p class="work-description">${this.escapeHtml(item.description || "Sin descripción.")}</p>
      <div class="work-meta-grid">
        <div><small>Responsable</small><strong>${this.escapeHtml(responsible.name)}</strong></div>
        <div><small>Participantes</small><strong>${(item.participantEmails || []).length}</strong></div>
        <div><small>Fecha objetivo</small><strong>${item.dueDate ? new Intl.DateTimeFormat("es-AR").format(new Date(item.dueDate+"T12:00:00")) : "Sin fecha"}</strong></div>
        <div><small>Avance</small><strong>${item.progress} %</strong></div>
      </div>
      <div class="progress"><span style="width:${item.progress}%"></span></div>
      ${latest ? `<div class="latest-update"><small>Último avance — ${this.personName(latest.authorEmail)} · ${this.date(latest.at)}</small><p>${this.escapeHtml(latest.comment)}</p></div>` : `<div class="latest-update"><small>Sin avances informados.</small></div>`}
      <div class="work-actions">
        ${this.canUpdateWorkItem(item) ? `<button class="btn btn-primary" data-action="add-work-progress" data-work-id="${item.id}">Registrar avance</button>` : ""}
        ${this.canManage() ? `<button class="btn btn-secondary" data-action="manage-work-participants" data-work-id="${item.id}">Participantes</button><button class="btn btn-secondary" data-action="edit-work-item" data-work-id="${item.id}">Editar</button><button class="btn btn-danger" data-action="delete-work-item" data-work-id="${item.id}">Eliminar</button>` : ""}
      </div>
      ${(item.updates || []).length ? `<details class="updates-details"><summary>Ver historial de avances (${item.updates.length})</summary>${item.updates.slice().sort((a,b)=>new Date(b.at)-new Date(a.at)).map(u=>`<div class="update-row"><strong>${u.progress} % · ${this.personName(u.authorEmail)}</strong><small>${this.date(u.at)}</small><p>${this.escapeHtml(u.comment)}</p></div>`).join("")}</details>` : ""}
    </article>`;
  },

  bindViewEvents() {
    document.querySelectorAll("[data-route-action]").forEach(b => b.addEventListener("click", () => this.go(b.dataset.routeAction)));
    document.querySelectorAll("[data-task]").forEach(el => el.addEventListener("click", async () => {
      this.selectedTask = await this.adapter.getTask(el.dataset.task);
      this.taskTab = "summary";
      await this.render();
    }));
    document.querySelectorAll("[data-filter]").forEach(b => b.addEventListener("click", async () => {
      this.filter = b.dataset.filter; await this.render();
    }));
    document.querySelectorAll("[data-tab]").forEach(b => b.addEventListener("click", async () => {
      this.taskTab = b.dataset.tab; await this.render();
    }));
    document.querySelectorAll("[data-tab-jump]").forEach(b => b.addEventListener("click", async () => { this.taskTab=b.dataset.tabJump; await this.render(); }));
    document.querySelectorAll('[data-action="new-task"]').forEach(b => b.addEventListener("click", () => this.startNewTask()));
    document.querySelectorAll('[data-action="open-documents"]').forEach(b => b.addEventListener("click", () => this.toast("Abrí OneDrive desde el teléfono o la computadora para administrar los documentos.")));
    document.querySelectorAll('[data-action="choose-master-file"]').forEach(b => b.addEventListener("click", () => this.runFileOperation(async () => {
      if (window.MeeNative?.pendingSubmissionInfo?.()?.hasPending) {
        throw new Error("Existe una sincronización pendiente. Confirmala antes de elegir otro archivo maestro.");
      }
      const sync = await this.adapter.getSyncInfo();
      if (sync.dirtyCount && !confirm("Hay cambios locales pendientes. Elegir otro archivo los reemplazará. ¿Continuar?")) return "Operación cancelada.";
      const result = await this.fileSync.linkAndImportMaster();
      if (result.fallback) { document.querySelector("#file-master-input").click(); return ""; }
      await this.afterFileDataChange();
      return `Archivo maestro cargado en modo solo lectura: ${result.linkedFile}.`;
    })));

    document.querySelectorAll('[data-action="reload-master"]').forEach(b => b.addEventListener("click", () => this.runFileOperation(async () => {
      const sync = await this.adapter.getSyncInfo();
      const hasPendingSubmission = !!window.MeeNative?.pendingSubmissionInfo?.()?.hasPending;
      if (!hasPendingSubmission && sync.dirtyCount && !confirm("Hay cambios locales pendientes. Actualizar desde el archivo los reemplazará. ¿Continuar?")) return "Operación cancelada.";
      const result = await this.fileSync.reloadRememberedMaster();
      if (result.missing) return "No hay un archivo recordado. Elegí el archivo maestro.";
      if (result.confirmedSubmission) {
        await this.render();
        return `Sincronización confirmada por el servidor. Versión ${result.dataVersion}.`;
      }
      await this.afterFileDataChange();
      return `Base actualizada desde ${result.linkedFile}.`;
    })));

    document.querySelectorAll('[data-action="save-master"]').forEach(b => b.addEventListener("click", () => this.runFileOperation(async () => {
      const result = await this.fileSync.saveMasterToLinkedFile();
      if (result.noChanges) return "No hay cambios pendientes y el maestro no requiere recuperación.";
      await this.render();
      if (!["submissionPrepared", "pendingExists"].includes(result.mode)) {
        return "No se pudo preparar el envío seguro.";
      }
      const email = this.ensureSyncDestinationEmail();
      if (email === null) {
        return "El envío quedó preparado localmente. Configurá el correo y reenviá cuando quieras.";
      }
      await this.fileSync.sharePendingSubmission("outlook");
      await this.render();
      return result.recovery
        ? "Se abrió Outlook con una recuperación segura. Power Automate debe aplicar el archivo y luego tenés que confirmar la sincronización."
        : "Se abrió Outlook con el envío preparado. La sincronización todavía no está confirmada.";
    })));

    document.querySelectorAll('[data-action="share-pending-outlook"]').forEach(b => b.addEventListener("click", () => this.runFileOperation(async () => {
      const email = this.ensureSyncDestinationEmail();
      if (email === null) return "El envío sigue pendiente. No se abrió Outlook.";
      await this.fileSync.sharePendingSubmission("outlook");
      return "Se abrió Outlook con el mismo envío pendiente. La sincronización todavía no está confirmada.";
    })));

    document.querySelectorAll('[data-action="save-pending-copy"]').forEach(b => b.addEventListener("click", () => this.runFileOperation(async () => {
      await this.fileSync.sharePendingSubmission("save");
      return "Se guardó una copia del submission. Los cambios siguen pendientes hasta confirmar el ACK del servidor.";
    })));

    const describeConfirmResult = result => {
      if (result.confirmed) return `Sincronización confirmada por el servidor. Versión ${result.dataVersion}.`;
      const diagParts = [];
      if (result.masterFileName) diagParts.push(`archivo ${result.masterFileName}`);
      if (result.uriAuthority) diagParts.push(`proveedor ${result.uriAuthority}`);
      if (Number.isFinite(result.baseDataVersion) && result.baseDataVersion >= 0) diagParts.push(`base ${result.baseDataVersion}`);
      if (Number.isFinite(result.targetDataVersion) && result.targetDataVersion >= 0) diagParts.push(`objetivo ${result.targetDataVersion}`);
      if (Number.isFinite(result.remoteDataVersion) && result.remoteDataVersion >= 0) diagParts.push(`remoto ${result.remoteDataVersion}`);
      if (result.expectedSubmissionId) diagParts.push(`esperado ${String(result.expectedSubmissionId).slice(0, 8)}…`);
      diagParts.push(result.ackSubmissionId ? `ACK ${String(result.ackSubmissionId).slice(0, 8)}…` : "ACK vacío");
      diagParts.push(`resultado ${result.status || "pending"}`);
      const diag = ` (${diagParts.join(" · ")})`;
      if (result.status === "conflict") throw new Error((result.detail || "Existe un conflicto con otra versión del maestro.") + diag);
      return (result.detail || "Power Automate todavía no confirmó este envío.") + diag;
    };

    document.querySelectorAll('[data-action="refresh-sync-state"]').forEach(b => b.addEventListener("click", () => {
      this.refreshSyncState({ silent: false });
    }));

    document.querySelectorAll('[data-action="confirm-sync"]').forEach(b => b.addEventListener("click", () => {
      if (!this.fileSync.isAndroidApp()) {
        document.querySelector("#file-confirm-input").click();
        return;
      }
      this.runFileOperation(async () => {
        const result = await this.fileSync.confirmPendingSync();
        await this.render();
        return describeConfirmResult(result);
      });
    }));

    document.querySelectorAll('[data-action="confirm-sync-picked"]').forEach(b => b.addEventListener("click", () => {
      if (!this.fileSync.isAndroidApp()) {
        document.querySelector("#file-confirm-input").click();
        return;
      }
      this.runFileOperation(async () => {
        const result = await this.fileSync.confirmPendingSyncWithPickedFile();
        await this.render();
        return describeConfirmResult(result);
      });
    }));

    document.querySelectorAll('[data-action="configure-submission-folder"]').forEach(b => b.addEventListener("click", () => this.runFileOperation(async () => {
      const result = await this.fileSync.configureSubmissionFolder();
      await this.render();
      return `Carpeta de entrada autorizada: ${result.folderName || "carpeta seleccionada"}. Ahora podés enviar con un solo botón.`;
    })));

    document.querySelectorAll('[data-action="send-pending-to-folder"]').forEach(b => b.addEventListener("click", () => this.runFileOperation(async () => {
      if (!confirm("Esta opción requiere que ya esté activo el segundo flujo de Power Automate para MEE_SEG_ENTRADAS. ¿Continuar?")) {
        return "Operación cancelada.";
      }
      const result = await this.fileSync.sendPendingToFolder();
      await this.render();
      return result.mode === "alreadyInFolder"
        ? `El archivo ${result.fileName} ya estaba en la carpeta de entrada; no se duplicó. Falta el ACK de Power Automate.`
        : `Archivo ${result.fileName} depositado en la carpeta de entrada. La sincronización queda pendiente hasta confirmar el ACK.`;
    })));

    document.querySelectorAll('[data-action="configure-sync-email"]').forEach(b => b.addEventListener("click", async () => {
      const current = this.fileSync.getSyncDestinationEmail();
      const entered = prompt("Correo corporativo destino de Power Automate:", current);
      if (entered === null) return;
      if (!this.fileSync.setSyncDestinationEmail(String(entered).trim())) {
        this.toast("Correo inválido.");
        return;
      }
      this.toast(this.fileSync.isAndroidApp()
        ? "Correo de sincronización guardado."
        : "Correo de sincronización guardado en este navegador.");
      await this.render();
    }));

    document.querySelectorAll('[data-action="link-master-file"]').forEach(b => b.addEventListener("click", () => this.runFileOperation(async () => {
      if ((await this.adapter.getSyncInfo()).dirtyCount && !confirm("Hay cambios locales pendientes. Vincular una base maestra los reemplazará. ¿Continuar?")) return "Operación cancelada.";
      const result = await this.fileSync.linkAndImportMaster();
      if (result.fallback) { document.querySelector("#file-master-input").click(); return ""; }
      await this.afterFileDataChange();
      return `Archivo vinculado: ${result.linkedFile}.`;
    })));
    document.querySelectorAll('[data-action="import-master-file"]').forEach(b => b.addEventListener("click", () => document.querySelector("#file-master-input").click()));
    document.querySelectorAll('[data-action="save-linked-master"]').forEach(b => b.addEventListener("click", () => this.runFileOperation(async () => {
      const result = await this.fileSync.saveMasterToLinkedFile();
      await this.render();
      return result.mode === "overwrite" ? `Archivo maestro actualizado: ${result.linkedFile}.` : "El archivo maestro no fue reemplazado; los cambios continúan pendientes.";
    })));
    document.querySelectorAll('[data-action="export-master"]').forEach(b => b.addEventListener("click", () => this.runFileOperation(async () => {
      const result = await this.fileSync.exportMasterDownload();
      await this.render();
      if (result.noChanges) return "No hay cambios pendientes para guardar.";
      return result.mode === "overwrite" ? `Cambios guardados en ${result.linkedFile}.` : "El archivo maestro no fue reemplazado; los cambios continúan pendientes.";
    })));
    document.querySelectorAll('[data-action="export-changes"]').forEach(b => b.addEventListener("click", () => this.runFileOperation(async () => {
      const result = await this.fileSync.exportChangesDownload(); await this.render();
      return result.empty ? "No hay cambios pendientes para exportar." : (result.shared ? `${result.count} tarea(s) enviadas al menú Compartir.` : `${result.count} tarea(s) preparadas en ${result.fileName}.`);
    })));
    document.querySelectorAll('[data-action="consolidate-changes"]').forEach(b => b.addEventListener("click", () => document.querySelector("#file-changes-input").click()));
    document.querySelectorAll('[data-action="generate-report"]').forEach(b => b.addEventListener("click", () => this.runFileOperation(async () => {
      const result = await this.fileSync.generateExcelReport();
      return result.shared ? `Reporte Excel enviado al menú Compartir: ${result.fileName}.` : `Reporte Excel generado: ${result.fileName}.`;
    })));
    document.querySelectorAll('[data-action="generate-pdf-report"]').forEach(b => b.addEventListener("click", () => this.runFileOperation(async () => {
      const result = await this.fileSync.generatePdfReport();
      return result.shared ? `Reporte PDF enviado al menú Compartir: ${result.fileName}.` : `Reporte PDF generado: ${result.fileName}.`;
    })));
    document.querySelectorAll('[data-action="task-pdf"]').forEach(b => b.addEventListener("click", () => this.runFileOperation(async () => {
      const result = await this.fileSync.generateTaskPdf(this.selectedTask.code);
      return result.shared ? `Minuta enviada al menú Compartir: ${result.fileName}.` : `Minuta PDF generada: ${result.fileName}.`;
    })));
    document.querySelectorAll('[data-action="delete-task"]').forEach(b => b.addEventListener("click", () => this.promptDeleteTask()));
    document.querySelectorAll('[data-action="change-local-user"]').forEach(b => b.addEventListener("click", () => this.promptChangeLocalUser()));
    document.querySelectorAll('[data-action="add-contribution"]').forEach(b => b.addEventListener("click", () => this.promptContribution()));
    document.querySelectorAll('[data-action="edit-task-text"]').forEach(b => b.addEventListener("click", () => this.promptEditTaskText(b.dataset.field)));
    document.querySelectorAll('[data-action="add-participant"]').forEach(b => b.addEventListener("click", () => this.promptAddParticipant()));
    document.querySelectorAll('[data-action="remove-participant"]').forEach(b => b.addEventListener("click", () => this.promptRemoveParticipant(b.dataset.email)));
    document.querySelectorAll('[data-action="add-work-item"]').forEach(b => b.addEventListener("click", () => this.promptWorkItem()));
    document.querySelectorAll('[data-action="edit-work-item"]').forEach(b => b.addEventListener("click", () => this.promptWorkItem(b.dataset.workId)));
    document.querySelectorAll('[data-action="delete-work-item"]').forEach(b => b.addEventListener("click", () => this.promptDeleteWorkItem(b.dataset.workId)));
    document.querySelectorAll('[data-action="add-work-progress"]').forEach(b => b.addEventListener("click", () => this.promptWorkProgress(b.dataset.workId)));
    document.querySelectorAll('[data-action="manage-work-participants"]').forEach(b => b.addEventListener("click", () => this.promptWorkParticipants(b.dataset.workId)));
    document.querySelectorAll('[data-action="add-pending"]').forEach(b => b.addEventListener("click", () => this.promptPending()));
    document.querySelectorAll('[data-action="edit-pending"]').forEach(b => b.addEventListener("click", () => this.promptEditPending(b.dataset.pendingId)));
    document.querySelectorAll('[data-action="delete-pending"]').forEach(b => b.addEventListener("click", () => this.promptDeletePending(b.dataset.pendingId)));
    document.querySelectorAll('[data-action="toggle-pending"]').forEach(b => b.addEventListener("click", () => this.togglePendingStatus(b.dataset.pendingId))); 
    document.querySelectorAll('[data-action="close-task"]').forEach(b => b.addEventListener("click", () => this.promptClose()));
    document.querySelectorAll('[data-action="minute"]').forEach(b => b.addEventListener("click", () => this.previewMinute()));
    document.querySelectorAll('[data-action="save-quick-note"]').forEach(b => b.addEventListener("click", async () => {
      try {
        this.notes.createFromDraft(this.user);
        this.toast("Apunte guardado.");
        await this.render();
      } catch (error) {
        this.toast(error.message || String(error));
      }
    }));
    document.querySelectorAll('[data-action="clear-note-draft"]').forEach(b => b.addEventListener("click", async () => {
      if (!confirm("¿Limpiar el borrador actual?")) return;
      this.notes.clearDraft(this.user);
      await this.render();
    }));
    document.querySelectorAll('[data-action="edit-note"]').forEach(b => b.addEventListener("click", () => this.promptEditNote(b.dataset.noteId)));
    document.querySelectorAll('[data-action="delete-note"]').forEach(b => b.addEventListener("click", () => this.promptDeleteNote(b.dataset.noteId)));
    document.querySelectorAll('[data-action="reopen-note"]').forEach(b => b.addEventListener("click", async () => {
      this.notes.reopen(this.user, b.dataset.noteId);
      this.toast("Apunte vuelto a pendiente.");
      await this.render();
    }));
    document.querySelectorAll('[data-action="create-task-from-note"]').forEach(b => b.addEventListener("click", () => this.startTaskFromNote(b.dataset.noteId)));
    document.querySelectorAll('[data-action="load-note-into-task"]').forEach(b => b.addEventListener("click", () => this.loadSelectedNoteIntoTask()));

    const quickDraft = document.querySelector("#quick-note-draft");
    if (quickDraft) {
      quickDraft.addEventListener("input", event => {
        clearTimeout(this.noteDraftTimer);
        const status = document.querySelector("#note-autosave-status");
        if (status) status.textContent = "Guardando…";
        this.noteDraftTimer = setTimeout(() => {
          const draft = this.notes.saveDraft(this.user, event.target.value);
          const label = document.querySelector("#note-autosave-status");
          if (label) label.textContent = `Guardado automáticamente ${this.date(draft.updatedAt)}`;
        }, 250);
      });
    }

    const noteSource = document.querySelector("#note-source-select");
    if (noteSource) {
      noteSource.addEventListener("change", event => {
        this.selectedNoteIdForTask = event.target.value || "";
        const note = this.selectedNoteIdForTask ? this.notes.get(this.user, this.selectedNoteIdForTask) : null;
        const preview = document.querySelector("#note-source-preview");
        if (preview) preview.textContent = note?.text || "Seleccioná un apunte para ver el contenido.";
      });
    }

    document.querySelectorAll('[data-action="reset-demo"]').forEach(b => b.addEventListener("click", async () => {
      if (!confirm("Esto eliminará los datos locales de esta PC y restaurará la base inicial. ¿Continuar?")) return;
      await this.adapter.resetDemo();
      this.people = await this.adapter.getPeople();
      this.user = await this.adapter.getCurrentUser();
      document.querySelector("#btn-profile").textContent = this.user.initials;
      this.toast("Datos iniciales restablecidos.");
      this.go("home");
    }));
    const form = document.querySelector("#new-task-form");
    if (form) form.addEventListener("submit", async e => {
      e.preventDefault();
      const d = Object.fromEntries(new FormData(form));
      const task = await this.adapter.createTask(d);
      if (d.sourceNoteId) {
        this.notes.markConverted(this.user, d.sourceNoteId, task.code);
      }
      this.selectedNoteIdForTask = "";
      this.toast(`Tarea ${task.code} creada.`);
      this.selectedTask = task;
      this.route = "tasks";
      await this.render();
    });
    const search = document.querySelector("#task-search");
    if (search) search.addEventListener("input", async e => {
      const q = e.target.value.toLowerCase();
      const tasks = (await this.adapter.getTasks()).filter(t => this.filter==="TODAS" || t.status===this.filter)
        .filter(t => `${t.code} ${t.title} ${t.location}`.toLowerCase().includes(q));
      document.querySelector("#task-results").innerHTML = tasks.map(t => this.taskCard(t)).join("") || `<div class="empty">Sin resultados.</div>`;
      document.querySelectorAll("#task-results [data-task]").forEach(el => el.addEventListener("click", async () => {
        this.selectedTask = await this.adapter.getTask(el.dataset.task); this.taskTab="summary"; await this.render();
      }));
    });
  },

  startNewTask() {
    if (!this.canManage()) {
      this.toast("El rol Operador es de solo lectura.");
      return;
    }
    this.selectedNoteIdForTask = "";
    this.go("newTask");
  },

  startTaskFromNote(noteId) {
    if (!this.canManage()) {
      this.toast("El rol Operador es de solo lectura.");
      return;
    }
    const note = this.notes.get(this.user, noteId);
    if (!note || note.status === "CONVERTIDO") {
      this.toast("El apunte ya no está disponible.");
      return;
    }
    this.selectedNoteIdForTask = noteId;
    this.go("newTask");
  },

  loadSelectedNoteIntoTask() {
    const select = document.querySelector("#note-source-select");
    const noteId = select?.value || "";
    if (!noteId) {
      this.toast("Elegí un apunte.");
      return;
    }

    const note = this.notes.get(this.user, noteId);
    if (!note) {
      this.toast("Apunte no encontrado.");
      return;
    }

    const problem = document.querySelector('#new-task-form textarea[name="problem"]');
    if (!problem) return;

    const current = problem.value.trim();
    if (!current) {
      problem.value = note.text;
    } else if (!current.includes(note.text)) {
      problem.value = `${current}\n\n${note.text}`;
    }

    this.selectedNoteIdForTask = noteId;
    problem.focus();
    this.toast("Apunte cargado en la tarea.");
  },

  promptEditNote(noteId) {
    if (!this.canManage()) return;
    const note = this.notes.get(this.user, noteId);
    if (!note) return;

    this.openModal(`<h3>Editar apunte</h3>
      <div class="field"><label>Texto</label><textarea id="modal-note-text">${this.escapeHtml(note.text)}</textarea></div>
      <button id="modal-save" type="button" class="btn btn-primary btn-full">Guardar cambios</button>`);

    document.querySelector("#modal-save").onclick = async () => {
      try {
        this.notes.update(this.user, noteId, document.querySelector("#modal-note-text").value);
        document.querySelector("#modal").close();
        this.toast("Apunte actualizado.");
        await this.render();
      } catch (error) {
        this.toast(error.message || String(error));
      }
    };
  },

  promptDeleteNote(noteId) {
    if (!this.canManage()) return;
    const note = this.notes.get(this.user, noteId);
    if (!note) return;

    this.openModal(`<h3>Eliminar apunte</h3>
      <div class="warning">Se eliminará este apunte guardado únicamente en este dispositivo.</div>
      <div class="panel" style="margin-top:12px"><p>${this.escapeHtml(note.text)}</p></div>
      <button id="modal-save" type="button" class="btn btn-danger btn-full">Eliminar apunte</button>`);

    document.querySelector("#modal-save").onclick = async () => {
      this.notes.remove(this.user, noteId);
      document.querySelector("#modal").close();
      this.toast("Apunte eliminado.");
      await this.render();
    };
  },

  peopleOptions(selectedEmail="", excludeEmails=[], preferredGroup="") {
    const available = (this.people || []).filter(person => !excludeEmails.includes(person.email) || person.email === selectedEmail);
    const suggested = preferredGroup ? available.filter(person => person.group === preferredGroup) : [];
    const others = preferredGroup ? available.filter(person => person.group !== preferredGroup) : available;
    const makeOptions = list => list.map(person => `<option value="${person.email}" ${person.email===selectedEmail?"selected":""}>${this.escapeHtml(person.name)}</option>`).join("");
    if (!preferredGroup) return makeOptions(others);
    return `${suggested.length ? `<optgroup label="${this.groupLabel(preferredGroup)} — sugeridos">${makeOptions(suggested)}</optgroup>` : ""}<optgroup label="Otras personas">${makeOptions(others)}</optgroup>`;
  },

  participantCheckboxes(selected=[], preferredGroup="", autoSelectPreferred=false) {
    const people = (this.people || []).slice().sort((a,b) => {
      const ap = preferredGroup && a.group === preferredGroup ? 0 : 1;
      const bp = preferredGroup && b.group === preferredGroup ? 0 : 1;
      return ap - bp || String(a.name).localeCompare(String(b.name), "es");
    });
    const effectiveSelected = selected.length ? selected : (autoSelectPreferred && preferredGroup ? people.filter(person => person.group === preferredGroup).map(person => person.email) : []);
    return `${preferredGroup ? `<div class="group-suggestion-note">Sugeridos automáticamente: ${this.groupLabel(preferredGroup)}. Podés marcar personas de cualquier otro grupo.</div>` : ""}<div class="check-grid">${people.map(person => `<label class="check-person ${person.group === preferredGroup ? "suggested-person" : ""}"><input type="checkbox" name="participantEmail" value="${person.email}" ${effectiveSelected.includes(person.email)?"checked":""}><span><strong>${this.escapeHtml(person.name)}</strong><small>${this.groupLabel(person.group)} · ${this.escapeHtml(person.cargo)}</small></span></label>`).join("")}</div>`;
  },

  promptAddParticipant() {
    const existing=this.selectedTask.participants || [];
    const available=(this.people || []).filter(p=>!existing.includes(p.email));
    if (!available.length) { this.toast("No quedan personas disponibles."); return; }
    this.openModal(`<h3>Agregar participante general</h3><div class="field"><label>Persona</label><select id="modal-person">${this.peopleOptions("",existing,this.selectedTask.group)}</select></div><div class="warning">Primero aparecen las personas del ${this.groupLabel(this.selectedTask.group)}. También podés elegir cualquier persona de otro grupo.</div><button id="modal-save" type="button" class="btn btn-primary btn-full">Agregar participante</button>`);
    document.querySelector("#modal-save").onclick=async()=>{ const email=document.querySelector("#modal-person").value; this.selectedTask=await this.adapter.addParticipant(this.selectedTask.code,email); document.querySelector("#modal").close(); this.toast("Participante agregado."); await this.render(); };
  },

  promptRemoveParticipant(email) {
    const p=this.person(email);
    this.openModal(`<h3>Eliminar participante</h3><div class="warning">Se quitará a <strong>${this.escapeHtml(p.name)}</strong> de los participantes generales. No se elimina de las tareas por sector donde ya esté asignado.</div><button id="modal-save" type="button" class="btn btn-danger btn-full">Eliminar participante</button>`);
    document.querySelector("#modal-save").onclick=async()=>{ this.selectedTask=await this.adapter.removeParticipant(this.selectedTask.code,email); document.querySelector("#modal").close(); this.toast("Participante eliminado."); await this.render(); };
  },

  promptWorkItem(id=null) {
    const item=id ? (this.selectedTask.workItems || []).find(w=>w.id===id) : null;
    const selectedParticipants=item?.participantEmails || [];
    this.openModal(`<h3>${item?"Editar":"Nueva"} tarea por sector</h3><div class="work-form-grid"><div class="field"><label>Título *</label><input id="work-title" value="${item?this.escapeHtml(item.title):""}" placeholder="Ej.: Cambio de tomas — ER17"></div><div class="field"><label>Sector / ubicación *</label><input id="work-sector" value="${item?this.escapeHtml(item.sector):""}" placeholder="Ej.: ER17"></div><div class="field field-span"><label>Descripción</label><textarea id="work-description" placeholder="Alcance concreto de esta tarea">${item?this.escapeHtml(item.description||""):""}</textarea></div><div class="field"><label>Responsable *</label><select id="work-responsible"><option value="">Seleccionar…</option>${this.peopleOptions(item?.responsibleEmail||"",[],this.selectedTask.group)}</select></div><div class="field"><label>Fecha objetivo</label><input id="work-due" type="date" value="${item?.dueDate||""}"></div></div><div class="field"><label>Participantes de esta tarea</label>${this.participantCheckboxes(selectedParticipants,this.selectedTask.group,!item)}</div><button id="modal-save" type="button" class="btn btn-primary btn-full">${item?"Guardar cambios":"Crear tarea por sector"}</button>`);
    document.querySelector("#modal-save").onclick=async()=>{
      const title=document.querySelector("#work-title").value.trim(), sector=document.querySelector("#work-sector").value.trim(), responsibleEmail=document.querySelector("#work-responsible").value;
      if(!title||!sector||!responsibleEmail){this.toast("Completá título, sector y responsable.");return;}
      const participantEmails=[...document.querySelectorAll('input[name="participantEmail"]:checked')].map(x=>x.value);
      if(!participantEmails.includes(responsibleEmail)) participantEmails.push(responsibleEmail);
      const payload={title,sector,description:document.querySelector("#work-description").value.trim(),responsibleEmail,dueDate:document.querySelector("#work-due").value,participantEmails};
      this.selectedTask=item ? await this.adapter.editWorkItem(this.selectedTask.code,id,payload) : await this.adapter.addWorkItem(this.selectedTask.code,payload);
      document.querySelector("#modal").close(); this.toast(item?"Tarea por sector actualizada.":"Tarea por sector creada."); await this.render();
    };
  },

  promptDeleteWorkItem(id) {
    const item=(this.selectedTask.workItems || []).find(w=>w.id===id); if(!item)return;
    this.openModal(`<h3>Eliminar tarea por sector</h3><div class="warning">Se eliminará <strong>${this.escapeHtml(item.title)}</strong> junto con su historial de avances de esta demostración.</div><button id="modal-save" type="button" class="btn btn-danger btn-full">Eliminar tarea</button>`);
    document.querySelector("#modal-save").onclick=async()=>{this.selectedTask=await this.adapter.deleteWorkItem(this.selectedTask.code,id);document.querySelector("#modal").close();this.toast("Tarea por sector eliminada.");await this.render();};
  },

  promptWorkParticipants(id) {
    const item=(this.selectedTask.workItems || []).find(w=>w.id===id); if(!item)return;
    this.openModal(`<h3>Participantes — ${this.escapeHtml(item.title)}</h3><div class="field"><label>Responsable</label><select id="work-responsible">${this.peopleOptions(item.responsibleEmail,[],this.selectedTask.group)}</select></div><div class="field"><label>Participantes</label>${this.participantCheckboxes(item.participantEmails||[],this.selectedTask.group,false)}</div><button id="modal-save" type="button" class="btn btn-primary btn-full">Guardar asignación</button>`);
    document.querySelector("#modal-save").onclick=async()=>{const responsibleEmail=document.querySelector("#work-responsible").value;const participantEmails=[...document.querySelectorAll('input[name="participantEmail"]:checked')].map(x=>x.value);if(!participantEmails.includes(responsibleEmail))participantEmails.push(responsibleEmail);this.selectedTask=await this.adapter.setWorkItemParticipants(this.selectedTask.code,id,responsibleEmail,participantEmails);document.querySelector("#modal").close();this.toast("Participantes actualizados.");await this.render();};
  },

  promptWorkProgress(id) {
    const item=(this.selectedTask.workItems || []).find(w=>w.id===id); if(!item)return;
    this.openModal(`<h3>Registrar avance</h3><div class="panel compact-panel"><strong>${this.escapeHtml(item.title)}</strong><p>${this.escapeHtml(item.sector)} · Avance actual ${item.progress} %</p></div><div class="field"><label>Nuevo avance (%)</label><input id="progress-value" type="number" min="0" max="100" step="5" value="${item.progress}"></div><div class="field"><label>Detalle del avance *</label><textarea id="progress-comment" placeholder="Qué se realizó, resultado y próximos pasos"></textarea></div><button id="modal-save" type="button" class="btn btn-primary btn-full">Guardar avance</button>`);
    document.querySelector("#modal-save").onclick=async()=>{const progress=Math.max(0,Math.min(100,Number(document.querySelector("#progress-value").value)));const comment=document.querySelector("#progress-comment").value.trim();if(!comment){this.toast("Escribí el detalle del avance.");return;}this.selectedTask=await this.adapter.addWorkProgress(this.selectedTask.code,id,progress,comment);document.querySelector("#modal").close();this.toast("Avance registrado.");await this.render();};
  },

  promptEditTaskText(field) {
    const isProblem = field === "problem";
    const title = isProblem ? "Editar problema detectado" : "Editar propuesta inicial";
    const label = isProblem ? "Problema o necesidad" : "Propuesta inicial";
    const current = isProblem ? this.selectedTask.problem : (this.selectedTask.proposal || "");
    this.openModal(`<h3>${title}</h3><div class="warning">Este cambio reemplaza el texto principal y quedará registrado en el historial. Para sumar información sin reemplazarlo, usá “Agregar aporte o mejora”.</div><div class="field" style="margin-top:14px"><label>${label}</label><textarea id="modal-text">${current}</textarea></div><button id="modal-save" type="button" class="btn btn-primary btn-full">Guardar cambios</button>`);
    document.querySelector("#modal-save").onclick = async () => {
      const text = document.querySelector("#modal-text").value.trim();
      if (!text) {
        this.toast("El texto no puede quedar vacío.");
        return;
      }
      this.selectedTask = await this.adapter.editTaskText(this.selectedTask.code, field, text);
      document.querySelector("#modal").close();
      this.toast(isProblem ? "Problema actualizado." : "Propuesta actualizada.");
      await this.render();
    };
  },



  promptChangeLocalUser() {
    const options = (this.people || []).slice().sort((a,b) => String(a.name).localeCompare(String(b.name), "es"))
      .map(person => `<option value="${person.email}" ${person.email === this.user.email ? "selected" : ""}>${this.escapeHtml(person.name)}</option>`).join("");
    this.openModal(`<h3>Cambiar usuario local</h3><div class="warning">Este modo no valida identidad con Microsoft. Seleccioná correctamente quién está usando la aplicación para que el historial quede identificado.</div><div class="field" style="margin-top:14px"><label>Usuario</label><select id="local-user-modal">${options}</select></div><button id="modal-save" type="button" class="btn btn-primary btn-full">Usar este usuario</button>`);
    document.querySelector("#modal-save").onclick = async () => {
      await this.adapter.setCurrentUser(document.querySelector("#local-user-modal").value);
      this.user = await this.adapter.getCurrentUser();
      document.querySelector("#btn-profile").textContent = this.user.initials;
      document.querySelector("#modal").close();
      this.toast(`Usuario activo: ${this.user.name}.`);
      await this.render();
    };
  },

  promptContribution() {
    this.openModal(`<h3>Agregar aporte o mejora</h3><div class="field"><label>Tipo</label><select id="modal-type"><option value="MEJORA">Mejora</option><option value="PROPUESTA">Propuesta</option><option value="AVANCE">Avance</option><option value="OBSERVACION">Observación</option><option value="DECISION">Decisión</option></select></div><div class="field"><label>Detalle</label><textarea id="modal-text" placeholder="Escribí el aporte, mejora o avance"></textarea></div><button id="modal-save" type="button" class="btn btn-primary btn-full">Guardar aporte</button>`);
    document.querySelector("#modal-save").onclick = async () => {
      const text = document.querySelector("#modal-text").value.trim();
      if (!text) return;
      this.selectedTask = await this.adapter.addContribution(this.selectedTask.code,text,document.querySelector("#modal-type").value);
      document.querySelector("#modal").close(); this.toast("Aporte guardado."); await this.render();
    };
  },

  promptPending() {
    this.openModal(`<h3>Agregar pendiente</h3><div class="field"><label>Descripción</label><textarea id="modal-text" placeholder="Acción concreta a realizar"></textarea></div><button id="modal-save" type="button" class="btn btn-primary btn-full">Guardar pendiente</button>`);
    document.querySelector("#modal-save").onclick = async () => {
      const text = document.querySelector("#modal-text").value.trim();
      if (!text) return;
      this.selectedTask = await this.adapter.addPending(this.selectedTask.code,text);
      document.querySelector("#modal").close(); this.toast("Pendiente guardado."); await this.render();
    };
  },

  promptEditPending(id) {
    const item = this.selectedTask.pending.find(p => p.id === id);
    if (!item) return;
    this.openModal(`<h3>Editar pendiente</h3><div class="field"><label>Descripción</label><textarea id="modal-text">${item.description}</textarea></div><button id="modal-save" type="button" class="btn btn-primary btn-full">Guardar cambios</button>`);
    document.querySelector("#modal-save").onclick = async () => {
      const text = document.querySelector("#modal-text").value.trim();
      if (!text) return;
      this.selectedTask = await this.adapter.editPending(this.selectedTask.code,id,text);
      document.querySelector("#modal").close(); this.toast("Pendiente actualizado."); await this.render();
    };
  },

  promptDeletePending(id) {
    const item = this.selectedTask.pending.find(p => p.id === id);
    if (!item) return;
    this.openModal(`<h3>Eliminar pendiente</h3><div class="warning">Se eliminará este pendiente:</div><div class="panel" style="margin-top:12px"><p>${item.description}</p></div><button id="modal-save" type="button" class="btn btn-danger btn-full">Eliminar pendiente</button>`);
    document.querySelector("#modal-save").onclick = async () => {
      this.selectedTask = await this.adapter.deletePending(this.selectedTask.code,id);
      document.querySelector("#modal").close(); this.toast("Pendiente eliminado."); await this.render();
    };
  },

  async togglePendingStatus(id) {
    this.selectedTask = await this.adapter.togglePending(this.selectedTask.code,id);
    this.toast("Estado del pendiente actualizado.");
    await this.render();
  },

  promptDeleteTask() {
    if (!this.canManage()) return;
    const task = this.selectedTask;
    this.openModal(`<h3>Eliminar tarea completa</h3><div class="warning">Esta acción eliminará definitivamente la tarea <strong>${this.escapeHtml(task.code)} — ${this.escapeHtml(task.title)}</strong>, incluyendo pendientes, aportes, participantes, acciones por sector e historial.</div><button id="modal-save" type="button" class="btn btn-danger btn-full">Eliminar tarea definitivamente</button>`);
    document.querySelector("#modal-save").onclick = async () => {
      await this.adapter.deleteTask(task.code);
      document.querySelector("#modal").close();
      this.selectedTask = null;
      this.route = "tasks";
      this.toast("Tarea eliminada.");
      await this.render();
    };
  },

  promptClose() {
    const open = this.selectedTask.pending.filter(p => p.status !== "COMPLETADO").length;
    const openWork = (this.selectedTask.workItems || []).filter(w => w.progress < 100).length;
    this.openModal(`<h3>Finalizar tarea</h3>${open || openWork ? `<div class="warning">La tarea tiene ${open} pendiente(s) abierto(s) y ${openWork} tarea(s) por sector sin completar. Podés cerrar en esta demostración, pero la versión real aplicará la regla configurada.</div>` : ""}<div class="field"><label>Resultado final *</label><textarea id="modal-text" placeholder="Indicá qué se realizó y el resultado obtenido"></textarea></div><button id="modal-save" type="button" class="btn btn-primary btn-full">Confirmar cierre</button>`);
    document.querySelector("#modal-save").onclick = async () => {
      const text = document.querySelector("#modal-text").value.trim();
      if (!text) return;
      this.selectedTask = await this.adapter.closeTask(this.selectedTask.code,text);
      document.querySelector("#modal").close(); this.toast(`Cerrada por ${this.user.name}.`); await this.render();
    };
  },

  previewMinute() {
    const t = this.selectedTask;
    this.openModal(`<h3>Vista previa de minuta</h3><div class="panel"><p><strong>${t.code} — ${t.title}</strong><br><br>Estado: ${t.statusLabel}<br>Avance: ${t.progress} %<br>Responsable: ${t.responsible}<br><br>Tareas por sector: ${(t.workItems || []).length}<br>Aportes generales: ${t.contributions.length}<br>Pendientes: ${t.pending.length}<br>Participantes generales: ${(t.participants || []).length}<br>Documentos: ${t.documents.length}</p></div><div class="warning">Usá “Descargar minuta PDF” para generar el documento completo. El envío automático por Outlook se conectará más adelante.</div>`);
  },

  openModal(html) {
    document.querySelector("#modal-content").innerHTML = html;
    document.querySelector("#modal").showModal();
  },

  date(v) {
    return new Intl.DateTimeFormat("es-AR",{dateStyle:"short",timeStyle:"short"}).format(new Date(v));
  },

  toast(text) {
    const t = document.querySelector("#toast");
    t.textContent = text; t.classList.add("show");
    clearTimeout(this.toastTimer);
    this.toastTimer = setTimeout(() => t.classList.remove("show"),2200);
  }
};

window.addEventListener("mee-native-resume", () => app.handleNativeResume());
app.init();
