const auth = firebase.auth();

const YEARS = [2020, 2021, 2022, 2023, 2024, 2025, 2026];
const INITIAL_VISIBLE = 15;

const RD_POR_ANIO = {
  2020: 966,
  2021: 1460,
  2022: 1615,
  2023: 2030,
  2024: 1560,
  2025: 1690,
  2026: 30,
};

const MODULES = ["RD UGEL", "RD Expedientes", "RD Profesores"];
const PROFESSOR_SHEETS = ["2026", "2025", "ASCENSO", "DATA", "CONTRATOS"];

const STORAGE_KEYS = {
  rdData: "ugel_rd_data_v5_profesores_secciones",
  loanData: "ugel_loan_data_v1",
};

let currentLoanFilter = "todos";
let appInitialized = false;

/* ========================= AUTH ========================= */
function setupAuth() {
  const loginForm = document.getElementById("loginForm");
  const logoutBtn = document.getElementById("logoutBtn");

  if (loginForm && loginForm.dataset.bound !== "true") {
    loginForm.addEventListener("submit", async (e) => {
      e.preventDefault();

      const email = document.getElementById("loginEmail").value.trim();
      const password = document.getElementById("loginPassword").value;

      if (!email || !password) {
        showLoginError("Completa correo y contraseña.");
        return;
      }

      clearLoginError();

      try {
        await auth.signInWithEmailAndPassword(email, password);
      } catch (error) {
        console.error(error);
        showLoginError(getFirebaseErrorMessage(error));
      }
    });

    loginForm.dataset.bound = "true";
  }

  if (logoutBtn && logoutBtn.dataset.bound !== "true") {
    logoutBtn.addEventListener("click", async () => {
      try {
        await auth.signOut();
      } catch (error) {
        console.error(error);
        showToast("No se pudo cerrar sesión.");
      }
    });

    logoutBtn.dataset.bound = "true";
  }

  auth.onAuthStateChanged((user) => {
    if (user) {
      showApp(user);

      if (!appInitialized) {
        initializeMainApp();
        appInitialized = true;
      }
    } else {
      showLogin();
    }
  });
}

function initializeMainApp() {
  setupYearSelects();
  setupSidebar();
  setupNavigation();
  setupModuleButtons();
  setupSearch();
  setupModal();
  setupLoans();
  setupExcelImport();
  renderAllModules();
  renderLoans();
  updateStats();
  setTodayByDefault();
}

function getAppShell() {
  return document.getElementById("appShell") || document.querySelector(".app-shell");
}

function showApp(user) {
  const loginScreen = document.getElementById("loginScreen");
  const appShell = getAppShell();
  const userBox = document.getElementById("userBox");

  if (loginScreen) loginScreen.style.display = "none";
  if (appShell) appShell.style.display = "flex";

  if (userBox) {
    const name = user.displayName || user.email || "Usuario";
    userBox.textContent = name;
  }

  clearLoginError();
}

function showLogin() {
  const loginScreen = document.getElementById("loginScreen");
  const appShell = getAppShell();
  const loginForm = document.getElementById("loginForm");
  const userBox = document.getElementById("userBox");

  if (loginScreen) loginScreen.style.display = "flex";
  if (appShell) appShell.style.display = "none";
  if (userBox) userBox.textContent = "";

  if (loginForm) loginForm.reset();

  clearLoginError();
}

function showLoginError(message) {
  const loginError = document.getElementById("loginError");
  if (loginError) loginError.textContent = message;
}

function clearLoginError() {
  const loginError = document.getElementById("loginError");
  if (loginError) loginError.textContent = "";
}

function getFirebaseErrorMessage(error) {
  const code = error?.code || "";

  switch (code) {
    case "auth/invalid-email":
      return "El correo no es válido.";
    case "auth/user-disabled":
      return "Esta cuenta ha sido deshabilitada.";
    case "auth/user-not-found":
      return "No existe una cuenta con ese correo.";
    case "auth/wrong-password":
      return "La contraseña es incorrecta.";
    case "auth/invalid-credential":
      return "Credenciales inválidas.";
    case "auth/popup-closed-by-user":
      return "Se cerró la ventana de inicio con Google.";
    case "auth/network-request-failed":
      return "Error de red. Verifica tu conexión.";
    case "auth/too-many-requests":
      return "Demasiados intentos. Intenta más tarde.";
    default:
      return "No se pudo iniciar sesión.";
  }
}

/* ========================= DATOS / STORAGE ========================= */
const rdData = loadRDData();
const loans = loadLoans();

const expandedYears = {
  "RD UGEL": Object.fromEntries(YEARS.map((y) => [y, false])),
  "RD Expedientes": Object.fromEntries(YEARS.map((y) => [y, false])),
};

const visibleByYear = {
  "RD UGEL": Object.fromEntries(YEARS.map((y) => [y, INITIAL_VISIBLE])),
  "RD Expedientes": Object.fromEntries(YEARS.map((y) => [y, INITIAL_VISIBLE])),
};

const expandedProfessorSheets = Object.fromEntries(
  PROFESSOR_SHEETS.map((sheet) => [sheet, false])
);

const visibleProfessorBySheet = Object.fromEntries(
  PROFESSOR_SHEETS.map((sheet) => [sheet, INITIAL_VISIBLE])
);

function loadRDData() {
  const saved = localStorage.getItem(STORAGE_KEYS.rdData);

  if (saved) {
    try {
      return migrateRDData(JSON.parse(saved));
    } catch (error) {
      console.error("Error leyendo datos guardados:", error);
    }
  }

  const base = {};

  MODULES.forEach((moduleName) => {
    base[moduleName] = {};

    if (moduleName === "RD Profesores") {
      PROFESSOR_SHEETS.forEach((sheet) => {
        base[moduleName][sheet] = {
          loaded: true,
          count: 0,
          items: [],
        };
      });
    } else {
      YEARS.forEach((year) => {
        base[moduleName][year] = {
          loaded: false,
          count: RD_POR_ANIO[year] || 0,
          items: [],
        };
      });
    }
  });

  localStorage.setItem(STORAGE_KEYS.rdData, JSON.stringify(base));
  return base;
}

function migrateRDData(data) {
  const migrated = {};

  MODULES.forEach((moduleName) => {
    migrated[moduleName] = {};

    if (moduleName === "RD Profesores") {
      PROFESSOR_SHEETS.forEach((sheet) => {
        const oldValue = data?.[moduleName]?.[sheet];
        const items = extractItems(oldValue).map(normalizeProfessorRecord);

        migrated[moduleName][sheet] = {
          loaded: true,
          count: items.length,
          items,
        };
      });
    } else {
      YEARS.forEach((year) => {
        const oldValue = data?.[moduleName]?.[year];

        if (Array.isArray(oldValue)) {
          migrated[moduleName][year] = {
            loaded: true,
            count: oldValue.length,
            items: oldValue.map(normalizeGenericRecord),
          };
        } else if (oldValue && typeof oldValue === "object") {
          migrated[moduleName][year] = {
            loaded: Boolean(oldValue.loaded),
            count:
              typeof oldValue.count === "number"
                ? oldValue.count
                : Array.isArray(oldValue.items)
                ? oldValue.items.length
                : RD_POR_ANIO[year] || 0,
            items: Array.isArray(oldValue.items)
              ? oldValue.items.map(normalizeGenericRecord)
              : [],
          };
        } else {
          migrated[moduleName][year] = {
            loaded: false,
            count: RD_POR_ANIO[year] || 0,
            items: [],
          };
        }
      });
    }
  });

  localStorage.setItem(STORAGE_KEYS.rdData, JSON.stringify(migrated));
  return migrated;
}

function extractItems(value) {
  if (Array.isArray(value)) return value;
  if (value && typeof value === "object" && Array.isArray(value.items)) return value.items;
  return [];
}

function normalizeGenericRecord(item) {
  return {
    id: item?.id || generateId(),
    numero: String(item?.numero || "").trim(),
    asunto: String(item?.asunto || "").trim(),
    fecha: String(item?.fecha || "").trim(),
    responsable: String(item?.responsable || "").trim(),
    estado: normalizeGenericStatus(item?.estado || "Archivado"),
  };
}

function normalizeProfessorRecord(item) {
  return {
    id: item?.id || generateId(),
    numero: String(item?.numero || "").trim(),
    asunto: String(item?.asunto || "").trim(),
    fecha: String(item?.fecha || "").trim(),
    responsable: String(item?.responsable || item?.nombre || "").trim(),
    estado: normalizeProfessorStatus(item?.estado || "Pendiente"),
  };
}

function normalizeGenericStatus(status) {
  const value = String(status || "").trim().toLowerCase();

  if (value === "activo") return "Activo";
  if (value === "en revisión" || value === "en revision") return "En revisión";
  if (value === "prestado") return "Prestado";

  return "Archivado";
}

function normalizeProfessorStatus(status) {
  const value = String(status || "").trim().toLowerCase();

  if (
    value === "recogido" ||
    value === "entregado" ||
    value === "devuelto" ||
    value === "activo"
  ) {
    return "Recogido";
  }

  return "Pendiente";
}

function loadLoans() {
  const saved = localStorage.getItem(STORAGE_KEYS.loanData);

  if (saved) {
    try {
      return JSON.parse(saved);
    } catch (error) {
      console.error("Error leyendo préstamos:", error);
    }
  }

  localStorage.setItem(STORAGE_KEYS.loanData, JSON.stringify([]));
  return [];
}

function saveRDData() {
  localStorage.setItem(STORAGE_KEYS.rdData, JSON.stringify(rdData));
}

function saveLoans() {
  localStorage.setItem(STORAGE_KEYS.loanData, JSON.stringify(loans));
}

function createInitialRecords(moduleName, year) {
  const records = [];
  const cantidad = RD_POR_ANIO[year] || 0;

  for (let i = 1; i <= cantidad; i++) {
    records.push({
      id: generateId(),
      numero: `${moduleCode(moduleName)}-${year}-${String(i).padStart(3, "0")}`,
      asunto: `${moduleName} - Registro administrativo ${i} del año ${year}`,
      fecha: `${year}-${String((i % 12) + 1).padStart(2, "0")}-${String(
        (i % 28) + 1
      ).padStart(2, "0")}`,
      responsable:
        moduleName === "RD Expedientes" ? "Mesa de Partes" : "Secretaría General",
      estado: "Archivado",
    });
  }

  applyLoanStatusesToRecords(moduleName, year, records);
  return records;
}

function ensureYearLoaded(moduleName, key) {
  const blockData = rdData[moduleName]?.[key];
  if (!blockData) return;

  if (moduleName === "RD Profesores") {
    blockData.loaded = true;
    blockData.items = Array.isArray(blockData.items)
      ? blockData.items.map(normalizeProfessorRecord)
      : [];
    blockData.count = blockData.items.length;
    saveRDData();
    return;
  }

  if (blockData.loaded) return;

  blockData.items = createInitialRecords(moduleName, key);
  blockData.loaded = true;
  blockData.count = blockData.items.length;
  saveRDData();
}

function getYearItems(moduleName, key) {
  return rdData[moduleName]?.[key]?.items || [];
}

function getYearCount(moduleName, key) {
  return rdData[moduleName]?.[key]?.count || 0;
}

function moduleCode(moduleName) {
  if (moduleName === "RD UGEL") return "UGEL";
  if (moduleName === "RD Expedientes") return "EXP";
  return "PROF";
}

function generateId() {
  return `${Date.now()}-${Math.random().toString(16).slice(2, 10)}`;
}

/* ========================= CONFIGURACIÓN GENERAL ========================= */
function setupYearSelects() {
  const rdYear = document.getElementById("rdYear");
  const loanYear = document.getElementById("loanYear");

  if (rdYear && !rdYear.dataset.loaded) {
    YEARS.forEach((year) => {
      const opt = document.createElement("option");
      opt.value = year;
      opt.textContent = year;
      rdYear.appendChild(opt);
    });

    rdYear.dataset.loaded = "true";
    rdYear.value = "2026";
  }

  if (loanYear && !loanYear.dataset.loaded) {
    YEARS.forEach((year) => {
      const opt = document.createElement("option");
      opt.value = year;
      opt.textContent = year;
      loanYear.appendChild(opt);
    });

    loanYear.dataset.loaded = "true";
    loanYear.value = "2026";
  }
}

function getLocalDateISO() {
  const date = new Date();
  const offset = date.getTimezoneOffset();
  const local = new Date(date.getTime() - offset * 60000);
  return local.toISOString().split("T")[0];
}

function setTodayByDefault() {
  const today = getLocalDateISO();

  const loanDate = document.getElementById("loanDate");
  const rdDate = document.getElementById("rdDate");

  if (loanDate) loanDate.value = today;
  if (rdDate) rdDate.value = today;
}

function setupSidebar() {
  const menuBtn = document.getElementById("menuBtn");
  const sidebar = document.getElementById("sidebar");
  const overlay = document.getElementById("sidebarOverlay");

  if (!menuBtn || !sidebar || !overlay || menuBtn.dataset.bound === "true") return;

  menuBtn.addEventListener("click", () => {
    sidebar.classList.toggle("open");
    overlay.classList.toggle("show");
  });

  overlay.addEventListener("click", () => {
    sidebar.classList.remove("open");
    overlay.classList.remove("show");
  });

  menuBtn.dataset.bound = "true";
}

function setupNavigation() {
  const navItems = document.querySelectorAll(".nav-item");
  const sections = document.querySelectorAll(".content-section");

  navItems.forEach((item) => {
    if (item.dataset.bound === "true") return;

    item.addEventListener("click", () => {
      const target = item.dataset.section;

      navItems.forEach((n) => n.classList.remove("active"));
      item.classList.add("active");

      sections.forEach((section) => {
        section.classList.toggle("active", section.id === target);
      });

      document.getElementById("sidebar")?.classList.remove("open");
      document.getElementById("sidebarOverlay")?.classList.remove("show");
    });

    item.dataset.bound = "true";
  });

  document.querySelectorAll(".go-module-btn").forEach((btn) => {
    if (btn.dataset.bound === "true") return;

    btn.addEventListener("click", () => {
      const target = btn.dataset.target;
      if (target) goToSection(target);
    });

    btn.dataset.bound = "true";
  });
}

function goToSection(sectionId) {
  document.querySelectorAll(".nav-item").forEach((n) => {
    n.classList.toggle("active", n.dataset.section === sectionId);
  });

  document.querySelectorAll(".content-section").forEach((section) => {
    section.classList.toggle("active", section.id === sectionId);
  });

  document.getElementById("sidebar")?.classList.remove("open");
  document.getElementById("sidebarOverlay")?.classList.remove("show");
}

/* ========================= RENDER GENERAL ========================= */
function renderAllModules() {
  renderStandardModule("RD UGEL", "ugel-year-blocks", getSearchValueByModule("RD UGEL"));
  renderStandardModule(
    "RD Expedientes",
    "expedientes-year-blocks",
    getSearchValueByModule("RD Expedientes")
  );
  renderProfessorsModule(getSearchValueByModule("RD Profesores"));
}

function renderStandardModule(moduleName, containerId, searchTerm = "") {
  const container = document.getElementById(containerId);
  if (!container) return;

  container.innerHTML = "";

  const sortedYears = [...YEARS].sort((a, b) => {
    const countA = getFilteredCount(moduleName, a, searchTerm);
    const countB = getFilteredCount(moduleName, b, searchTerm);
    return countB - countA;
  });

  sortedYears.forEach((year) => {
    const isExpanded = expandedYears[moduleName][year];
    const visibleCount = visibleByYear[moduleName][year];
    const filteredCount = getFilteredCount(moduleName, year, searchTerm);

    let bodyHTML = "";

    if (isExpanded) {
      ensureYearLoaded(moduleName, year);
      const data = getYearItems(moduleName, year);
      const filtered = data.filter((item) => matchesStandardSearch(item, searchTerm));
      const visibleItems = filtered.slice(0, visibleCount);

      bodyHTML =
        visibleItems.length > 0
          ? visibleItems.map((item) => createRDItemHTML(item)).join("")
          : `<div class="empty-state">No se encontraron registros en este año.</div>`;

      if (filtered.length > visibleCount) {
        bodyHTML += `
          <div style="margin-top:12px;">
            <button class="add-more-btn load-more-btn" data-module="${moduleName}" data-year="${year}">
              Ver 15 más
            </button>
          </div>
        `;
      }
    }

    const yearCard = document.createElement("div");
    yearCard.className = "year-card glass-card";
    yearCard.innerHTML = `
      <div class="year-card-head">
        <div class="year-card-head-left">
          <div class="year-pill">${year}</div>
          <div class="year-title">
            <h4>${moduleName} - ${year}</h4>
            <p>Registros disponibles para consulta y administración</p>
          </div>
        </div>
        <div class="year-actions">
          <button type="button" class="toggle-rd-btn ${
            isExpanded ? "rotate" : ""
          }" data-module="${moduleName}" data-year="${year}">
            <span>▼</span>
          </button>
          <div class="year-count">${filteredCount} registros</div>
          <button type="button" class="add-more-btn" data-module="${moduleName}" data-year="${year}">
            Agregar 5 registros
          </button>
        </div>
      </div>
      <div class="rd-list ${isExpanded ? "" : "hidden"}" id="rd-${sanitizeKey(
      moduleName
    )}-${year}">
        ${bodyHTML}
      </div>
    `;

    container.appendChild(yearCard);
  });

  attachAddMoreEvents();
  attachLoadMoreEvents();
}

function renderProfessorsModule(searchTerm = "") {
  const container = document.getElementById("profesores-year-blocks");
  if (!container) return;

  container.innerHTML = "";

  PROFESSOR_SHEETS.forEach((sheetName) => {
    if (!rdData["RD Profesores"][sheetName]) {
      rdData["RD Profesores"][sheetName] = {
        loaded: true,
        count: 0,
        items: [],
      };
    }

    ensureYearLoaded("RD Profesores", sheetName);

    const isExpanded = expandedProfessorSheets[sheetName];
    const visibleCount = visibleProfessorBySheet[sheetName];
    const allItems = rdData["RD Profesores"][sheetName].items || [];
    const filtered = allItems.filter((item) => matchesProfessorSearch(item, searchTerm));
    const filteredCount = filtered.length;
    const pickedCount = filtered.filter((item) => item.estado === "Recogido").length;

    let bodyHTML = "";

    if (isExpanded) {
      const visibleItems = filtered.slice(0, visibleCount);

      bodyHTML =
        visibleItems.length > 0
          ? visibleItems.map((item) => createProfessorItemHTML(item, sheetName)).join("")
          : `<div class="empty-state">No se encontraron registros en esta sección.</div>`;

      if (filtered.length > visibleCount) {
        bodyHTML += `
          <div style="margin-top:12px;">
            <button class="add-more-btn load-more-prof-btn" data-sheet="${sheetName}">
              Ver 15 más
            </button>
          </div>
        `;
      }
    }

    const card = document.createElement("div");
    card.className = "year-card glass-card prof-year-card";
    card.innerHTML = `
      <div class="year-card-head">
        <div class="year-card-head-left">
          <div class="year-pill">${escapeHTML(sheetName)}</div>
          <div class="year-title">
            <h4>RD Profesores - ${escapeHTML(sheetName)}</h4>
            <p>Control exclusivo de entrega a docentes</p>
          </div>
        </div>
        <div class="year-actions">
          <button type="button" class="toggle-prof-btn toggle-rd-btn ${
            isExpanded ? "rotate" : ""
          }" data-sheet="${escapeHTML(sheetName)}">
            <span>▼</span>
          </button>
          <div class="year-count prof-count-badge">${filteredCount} registros</div>
          <div class="prof-note-badge">Recogidos: ${pickedCount}</div>
        </div>
      </div>
      <div class="rd-list prof-rd-list ${
        isExpanded ? "" : "hidden"
      }" id="prof-sheet-${sanitizeKey(sheetName)}">
        ${bodyHTML}
      </div>
    `;

    container.appendChild(card);
  });

  attachProfessorLoadMoreEvents();
}

function matchesStandardSearch(item, searchTerm) {
  if (!searchTerm) return true;

  const text =
    `${item.numero} ${item.asunto} ${item.fecha} ${item.responsable} ${item.estado}`.toLowerCase();

  return text.includes(searchTerm);
}

function matchesProfessorSearch(item, searchTerm) {
  if (!searchTerm) return true;

  const text = `${item.numero} ${item.responsable} ${item.asunto} ${item.estado}`.toLowerCase();
  return text.includes(searchTerm);
}

function getFilteredCount(moduleName, key, searchTerm) {
  const blockData = rdData[moduleName]?.[key];
  if (!blockData) return 0;

  if (moduleName === "RD Profesores") {
    ensureYearLoaded(moduleName, key);
    if (!searchTerm) return blockData.items.length;
    return blockData.items.filter((item) => matchesProfessorSearch(item, searchTerm)).length;
  }

  if (!searchTerm) return blockData.count;
  if (!blockData.loaded) ensureYearLoaded(moduleName, key);

  return blockData.items.filter((item) => matchesStandardSearch(item, searchTerm)).length;
}

function createRDItemHTML(item) {
  let statusClass = "archivado";

  if (item.estado === "Prestado") {
    statusClass = "prestado";
  } else if (item.estado === "Activo") {
    statusClass = "activo";
  } else if (item.estado === "En revisión") {
    statusClass = "revision";
  }

  return `
    <div class="rd-item">
      <div class="rd-main">
        <strong>${escapeHTML(item.numero)}</strong>
        <span>${escapeHTML(item.asunto)}</span>
      </div>
      <div class="rd-meta">
        <span><strong>Fecha:</strong> ${escapeHTML(item.fecha || "-")}</span><br>
        <span><strong>Estado:</strong> ${escapeHTML(item.estado)}</span>
      </div>
      <div class="rd-responsable">
        <span><strong>Responsable:</strong> ${escapeHTML(
          item.responsable || "No especificado"
        )}</span>
      </div>
      <div>
        <span class="rd-status ${statusClass}">${escapeHTML(item.estado)}</span>
      </div>
    </div>
  `;
}

function createProfessorItemHTML(item, sheetName) {
  const isPicked = item.estado === "Recogido";

  return `
    <div class="prof-rd-item">
      <div class="prof-rd-main">
        <strong>${escapeHTML(item.numero)}</strong>
        <span class="prof-name">${escapeHTML(item.responsable || "Sin nombre")}</span>
        <span class="prof-subject">${escapeHTML(item.asunto || "Sin asunto")}</span>
      </div>
      <div class="prof-rd-meta">
        <div class="prof-meta-line"><strong>Sección:</strong> ${escapeHTML(sheetName)}</div>
        <div class="prof-meta-line"><strong>Estado:</strong> ${escapeHTML(item.estado)}</div>
      </div>
      <div class="prof-actions">
        ${
          isPicked
            ? `<button type="button" class="prof-status-fixed picked">Recogido</button>`
            : `
              <button
                type="button"
                class="prof-btn prof-btn-pending"
                data-prof-status="Pendiente"
                data-prof-id="${item.id}"
                data-prof-sheet="${escapeHTML(sheetName)}"
              >
                Pendiente
              </button>
              <button
                type="button"
                class="prof-btn prof-btn-picked"
                data-prof-status="Recogido"
                data-prof-id="${item.id}"
                data-prof-sheet="${escapeHTML(sheetName)}"
              >
                Marcar recogido
              </button>
            `
        }
      </div>
    </div>
  `;
}

function attachAddMoreEvents() {
  document
    .querySelectorAll(".add-more-btn:not(.load-more-btn):not(.load-more-prof-btn)")
    .forEach((btn) => {
      btn.onclick = () => {
        const moduleName = btn.dataset.module;
        const year = Number(btn.dataset.year);

        if (!moduleName || moduleName === "RD Profesores") return;

        ensureYearLoaded(moduleName, year);

        const items = getYearItems(moduleName, year);
        const currentLength = items.length;
        const start = currentLength + 1;

        for (let i = start; i < start + 5; i++) {
          items.push({
            id: generateId(),
            numero: `${moduleCode(moduleName)}-${year}-${String(i).padStart(3, "0")}`,
            asunto: `${moduleName} - Nuevo registro añadido ${i}`,
            fecha: `${year}-${String(((i + 2) % 12) + 1).padStart(2, "0")}-${String(
              ((i + 5) % 28) + 1
            ).padStart(2, "0")}`,
            responsable:
              moduleName === "RD Expedientes" ? "Mesa de Partes" : "Secretaría General",
            estado: "Archivado",
          });
        }

        rdData[moduleName][year].count = items.length;
        saveRDData();
        renderAllModules();
        updateStats();
        showToast(`Se agregaron 5 registros archivados en ${moduleName} ${year}.`);
      };
    });
}

function attachLoadMoreEvents() {
  document.querySelectorAll(".load-more-btn").forEach((btn) => {
    btn.onclick = () => {
      const moduleName = btn.dataset.module;
      const year = Number(btn.dataset.year);

      visibleByYear[moduleName][year] += INITIAL_VISIBLE;

      renderStandardModule(
        moduleName,
        getContainerIdByModule(moduleName),
        getSearchValueByModule(moduleName)
      );
    };
  });
}

function attachProfessorLoadMoreEvents() {
  document.querySelectorAll(".load-more-prof-btn").forEach((btn) => {
    btn.onclick = () => {
      const sheetName = btn.dataset.sheet;
      if (!sheetName) return;

      visibleProfessorBySheet[sheetName] += INITIAL_VISIBLE;
      renderProfessorsModule(getSearchValueByModule("RD Profesores"));
    };
  });
}

function getContainerIdByModule(moduleName) {
  if (moduleName === "RD UGEL") return "ugel-year-blocks";
  if (moduleName === "RD Expedientes") return "expedientes-year-blocks";
  return "profesores-year-blocks";
}

function getSearchValueByModule(moduleName) {
  if (moduleName === "RD UGEL") {
    return (document.getElementById("search-rd-ugel")?.value || "").trim().toLowerCase();
  }

  if (moduleName === "RD Expedientes") {
    return (document.getElementById("search-rd-expedientes")?.value || "")
      .trim()
      .toLowerCase();
  }

  return (document.getElementById("search-rd-profesores")?.value || "")
    .trim()
    .toLowerCase();
}

/* ========================= BÚSQUEDA ========================= */
function setupSearch() {
  const searchUgel = document.getElementById("search-rd-ugel");
  const searchExp = document.getElementById("search-rd-expedientes");
  const searchProf = document.getElementById("search-rd-profesores");

  if (searchUgel && searchUgel.dataset.bound !== "true") {
    searchUgel.addEventListener("input", (e) => {
      renderStandardModule("RD UGEL", "ugel-year-blocks", e.target.value.trim().toLowerCase());
    });
    searchUgel.dataset.bound = "true";
  }

  if (searchExp && searchExp.dataset.bound !== "true") {
    searchExp.addEventListener("input", (e) => {
      renderStandardModule(
        "RD Expedientes",
        "expedientes-year-blocks",
        e.target.value.trim().toLowerCase()
      );
    });
    searchExp.dataset.bound = "true";
  }

  if (searchProf && searchProf.dataset.bound !== "true") {
    searchProf.addEventListener("input", (e) => {
      renderProfessorsModule(e.target.value.trim().toLowerCase());
    });
    searchProf.dataset.bound = "true";
  }
}

/* ========================= MODAL AGREGAR RD ========================= */
function setupModuleButtons() {
  const btnUgel = document.getElementById("add-rd-ugel");
  const btnExp = document.getElementById("add-rd-expedientes");

  if (btnUgel && btnUgel.dataset.bound !== "true") {
    btnUgel.addEventListener("click", () => openRDModal("RD UGEL"));
    btnUgel.dataset.bound = "true";
  }

  if (btnExp && btnExp.dataset.bound !== "true") {
    btnExp.addEventListener("click", () => openRDModal("RD Expedientes"));
    btnExp.dataset.bound = "true";
  }
}

function setupModal() {
  const modal = document.getElementById("rdModal");
  const closeModalBtn = document.getElementById("closeModalBtn");
  const rdForm = document.getElementById("rdForm");

  if (!modal || !closeModalBtn || !rdForm) return;
  if (rdForm.dataset.bound === "true") return;

  closeModalBtn.addEventListener("click", closeRDModal);

  modal.addEventListener("click", (e) => {
    if (e.target === modal) closeRDModal();
  });

  rdForm.addEventListener("submit", (e) => {
    e.preventDefault();

    requireAuthAction(() => {
      const moduleName = document.getElementById("rdModule").value;
      const year = Number(document.getElementById("rdYear").value);
      const numero = document.getElementById("rdNumber").value.trim();
      const fecha = document.getElementById("rdDate").value;
      const asunto = document.getElementById("rdSubject").value.trim();
      const responsable = document.getElementById("rdResponsible").value.trim();
      const estado = document.getElementById("rdStatus").value || "Archivado";

      if (!moduleName || !year || !numero || !fecha || !asunto) {
        showToast("Completa los campos obligatorios del registro RD.");
        return;
      }

      if (moduleName === "RD Profesores") {
        showToast("RD Profesores solo se carga desde Excel.");
        return;
      }

      ensureYearLoaded(moduleName, year);

      const exists = rdData[moduleName][year].items.some(
        (item) => item.numero.trim().toLowerCase() === numero.trim().toLowerCase()
      );

      if (exists) {
        showToast("Ya existe una RD con ese número en ese módulo y año.");
        return;
      }

      rdData[moduleName][year].items.unshift({
        id: generateId(),
        numero,
        asunto,
        fecha,
        responsable,
        estado,
      });

      rdData[moduleName][year].count = rdData[moduleName][year].items.length;

      saveRDData();
      renderAllModules();
      updateStats();
      closeRDModal();
      rdForm.reset();
      document.getElementById("rdYear").value = "2026";
      document.getElementById("rdDate").value = getLocalDateISO();
      showToast("Registro RD agregado correctamente.");
    });
  });

  rdForm.dataset.bound = "true";
}

function openRDModal(moduleName) {
  const modal = document.getElementById("rdModal");
  if (!modal) return;

  modal.classList.add("show");
  document.getElementById("rdModule").value = moduleName;
  document.getElementById("rdStatus").value = "Archivado";
}

function closeRDModal() {
  const modal = document.getElementById("rdModal");
  if (!modal) return;

  modal.classList.remove("show");
}

/* ========================= PRÉSTAMOS ========================= */
function setupLoans() {
  const loanForm = document.getElementById("loanForm");
  if (!loanForm || loanForm.dataset.bound === "true") return;

  loanForm.addEventListener("submit", (e) => {
    e.preventDefault();

    requireAuthAction(() => {
      const person = document.getElementById("loanPerson").value.trim();
      const moduleName = document.getElementById("loanModule").value;
      const year = Number(document.getElementById("loanYear").value);
      const rdNumber = document.getElementById("loanRdNumber").value.trim();
      const loanDate = document.getElementById("loanDate").value;
      const returnDate = document.getElementById("loanReturnDate").value;
      const notes = document.getElementById("loanNotes").value.trim();

      if (!person || !moduleName || !year || !rdNumber || !loanDate || !returnDate) {
        showToast("Completa todos los campos obligatorios del préstamo.");
        return;
      }

      if (moduleName === "RD Profesores") {
        showToast(
          "RD Profesores se controla desde su propio módulo con Pendiente o Recogido."
        );
        return;
      }

      ensureYearLoaded(moduleName, year);

      const existingPendingLoan = loans.find(
        (loan) =>
          loan.moduleName === moduleName &&
          Number(loan.year) === year &&
          loan.rdNumber === rdNumber &&
          loan.status === "pendiente"
      );

      if (existingPendingLoan) {
        showToast("Ese archivo ya está prestado.");
        return;
      }

      let rdItem = rdData[moduleName][year].items.find((item) => item.numero === rdNumber);

      if (!rdItem) {
        rdItem = {
          id: generateId(),
          numero: rdNumber,
          asunto: `${moduleName} - Registro agregado por préstamo`,
          fecha: loanDate,
          responsable:
            moduleName === "RD Expedientes" ? "Mesa de Partes" : "Secretaría General",
          estado: "Prestado",
        };

        rdData[moduleName][year].items.unshift(rdItem);
        rdData[moduleName][year].count = rdData[moduleName][year].items.length;
      } else {
        rdItem.estado = "Prestado";
      }

      loans.unshift({
        id: generateId(),
        person,
        moduleName,
        year,
        rdNumber,
        loanDate,
        promisedReturnDate: returnDate,
        notes,
        status: "pendiente",
        returnedAt: null,
        createdAt: new Date().toISOString(),
      });

      saveLoans();
      saveRDData();
      renderLoans();
      renderAllModules();
      updateStats();
      loanForm.reset();
      document.getElementById("loanDate").value = getLocalDateISO();
      document.getElementById("loanYear").value = "2026";
      showToast("Préstamo registrado correctamente.");
    });
  });

  document.querySelectorAll(".loan-tab").forEach((tab) => {
    if (tab.dataset.bound === "true") return;

    tab.addEventListener("click", () => {
      document.querySelectorAll(".loan-tab").forEach((t) => t.classList.remove("active"));
      tab.classList.add("active");
      currentLoanFilter = tab.dataset.loanFilter;
      renderLoans();
    });

    tab.dataset.bound = "true";
  });

  loanForm.dataset.bound = "true";
}

function renderLoans() {
  const tbody = document.getElementById("loanTableBody");
  if (!tbody) return;

  tbody.innerHTML = "";

  let filtered = [...loans];

  if (currentLoanFilter === "prestados" || currentLoanFilter === "pendientes") {
    filtered = filtered.filter((loan) => loan.status === "pendiente");
  } else if (currentLoanFilter === "devueltos") {
    filtered = filtered.filter((loan) => loan.status === "devuelto");
  }

  if (filtered.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="9">
          <div class="empty-state">No hay préstamos para mostrar en esta categoría.</div>
        </td>
      </tr>
    `;
    updateLoanMiniStats();
    return;
  }

  const today = getLocalDateISO();

  filtered.forEach((loan) => {
    const tr = document.createElement("tr");
    const overdue = loan.status === "pendiente" && loan.promisedReturnDate < today;

    if (overdue) tr.classList.add("overdue");

    tr.innerHTML = `
      <td>${escapeHTML(loan.person)}</td>
      <td>${escapeHTML(loan.moduleName)}</td>
      <td>${escapeHTML(String(loan.year))}</td>
      <td>${escapeHTML(loan.rdNumber)}</td>
      <td>${escapeHTML(loan.loanDate)}</td>
      <td>${escapeHTML(loan.promisedReturnDate)}</td>
      <td>
        <span class="status-pill ${
          loan.status === "pendiente"
            ? overdue
              ? "pending overdue-pill"
              : "pending"
            : "returned"
        }">
          ${
            loan.status === "pendiente" ? (overdue ? "Vencido" : "Pendiente") : "Devuelto"
          }
        </span>
      </td>
      <td>${loan.returnedAt ? escapeHTML(formatDateTime(loan.returnedAt)) : "-"}</td>
      <td>
        ${
          loan.status === "pendiente"
            ? `<button type="button" class="action-btn return-btn" data-return-id="${loan.id}">Marcar devolución</button>`
            : `<button type="button" class="action-btn disabled-btn" disabled>Devuelto</button>`
        }
      </td>
    `;

    tbody.appendChild(tr);
  });

  document.querySelectorAll("[data-return-id]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const id = btn.dataset.returnId;
      markAsReturned(id);
    });
  });

  updateLoanMiniStats();
}

function markAsReturned(id) {
  requireAuthAction(() => {
    const loan = loans.find((item) => item.id === id);
    if (!loan) return;

    loan.status = "devuelto";
    loan.returnedAt = new Date().toISOString();

    const moduleName = loan.moduleName;
    const year = Number(loan.year);
    const rdNumber = loan.rdNumber;

    ensureYearLoaded(moduleName, year);

    const rdItem = rdData[moduleName][year].items.find((item) => item.numero === rdNumber);
    if (rdItem) {
      rdItem.estado = "Archivado";
    }

    saveLoans();
    saveRDData();
    renderLoans();
    renderAllModules();
    updateStats();
    showToast("Devolución registrada con fecha y hora exacta.");
  });
}

function updateLoanMiniStats() {
  const pending = loans.filter((l) => l.status === "pendiente").length;
  const returned = loans.filter((l) => l.status === "devuelto").length;
  const total = loans.length;

  const miniPending = document.getElementById("miniPending");
  const miniReturned = document.getElementById("miniReturned");
  const miniTotalLoans = document.getElementById("miniTotalLoans");

  if (miniPending) miniPending.textContent = pending;
  if (miniReturned) miniReturned.textContent = returned;
  if (miniTotalLoans) miniTotalLoans.textContent = total;
}

function applyLoanStatusesToRecords(moduleName, year, records) {
  const pendingLoans = loans.filter(
    (loan) =>
      loan.moduleName === moduleName &&
      Number(loan.year) === Number(year) &&
      loan.status === "pendiente"
  );

  const map = new Map(records.map((item) => [item.numero, item]));

  pendingLoans.forEach((loan) => {
    if (map.has(loan.rdNumber)) {
      map.get(loan.rdNumber).estado = "Prestado";
    }
  });
}

/* ========================= RD PROFESORES - ESTADOS ========================= */
function updateProfessorStatus(id, sheetName, newStatus) {
  requireAuthAction(() => {
    const items = rdData["RD Profesores"][sheetName]?.items || [];
    const item = items.find((row) => row.id === id);
    if (!item) return;

    item.estado = normalizeProfessorStatus(newStatus);
    rdData["RD Profesores"][sheetName].count = items.length;

    saveRDData();
    renderProfessorsModule(getSearchValueByModule("RD Profesores"));
    updateStats();

    if (item.estado === "Recogido") {
      showToast("Registro marcado como recogido.");
    } else {
      showToast("Registro marcado como pendiente.");
    }
  });
}

/* ========================= IMPORTAR EXCEL ========================= */
function setupExcelImport() {
  const importBtn = document.getElementById("importExcelBtn");
  const fileInput = document.getElementById("excelFile");
  const fileNameBox = document.getElementById("selectedFileName");

  if (!importBtn || !fileInput || importBtn.dataset.bound === "true") return;

  fileInput.addEventListener("change", () => {
    const file = fileInput.files?.[0];
    if (fileNameBox) {
      fileNameBox.textContent = file
        ? `Archivo seleccionado: ${file.name}`
        : "Ningún archivo seleccionado.";
    }
  });

  importBtn.addEventListener("click", () => {
    requireAuthAction(() => {
      const file = fileInput.files?.[0];

      if (!file) {
        showToast("Selecciona un archivo Excel antes de importar.");
        return;
      }

      importExcel(file);
    });
  });

  importBtn.dataset.bound = "true";
}

function importExcel(file) {
  const reader = new FileReader();

  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      let importedCount = 0;
      let duplicateCount = 0;
      let skippedCount = 0;

      PROFESSOR_SHEETS.forEach((sheetName) => {
        if (!rdData["RD Profesores"][sheetName]) {
          rdData["RD Profesores"][sheetName] = {
            loaded: true,
            count: 0,
            items: [],
          };
        }

        rdData["RD Profesores"][sheetName].items = [];
        rdData["RD Profesores"][sheetName].count = 0;
      });

      workbook.SheetNames.forEach((rawSheetName) => {
        const normalizedSheetName = normalizeProfessorSheetName(rawSheetName);
        if (!normalizedSheetName) return;

        const worksheet = workbook.Sheets[rawSheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

        if (!rows || rows.length === 0) return;

        rows.forEach((row) => {
          const numero = String(row[0] || "").trim();
          const nombre = String(row[1] || "").trim();
          const asunto = String(row[2] || "").trim();

          if (!numero && !nombre && !asunto) {
            skippedCount++;
            return;
          }

          if (isHeaderLikeRow(numero, nombre, asunto)) {
            skippedCount++;
            return;
          }

          if (!numero) {
            skippedCount++;
            return;
          }

          const exists = rdData["RD Profesores"][normalizedSheetName].items.some(
            (item) => item.numero.trim().toLowerCase() === numero.trim().toLowerCase()
          );

          if (exists) {
            duplicateCount++;
            return;
          }

          rdData["RD Profesores"][normalizedSheetName].items.push({
            id: generateId(),
            numero,
            asunto: asunto || "Sin asunto",
            fecha: "",
            responsable: nombre || "No especificado",
            estado: "Pendiente",
          });

          importedCount++;
        });

        rdData["RD Profesores"][normalizedSheetName].count =
          rdData["RD Profesores"][normalizedSheetName].items.length;
      });

      saveRDData();
      renderProfessorsModule(getSearchValueByModule("RD Profesores"));
      updateStats();
      goToSection("rd-profesores");
      showToast(
        `Importación completada. Importados: ${importedCount}. Duplicados: ${duplicateCount}. Omitidos: ${skippedCount}.`
      );
    } catch (error) {
      console.error(error);
      showToast("Ocurrió un error al leer el archivo Excel.");
    }
  };

  reader.onerror = function () {
    showToast("No se pudo leer el archivo seleccionado.");
  };

  reader.readAsArrayBuffer(file);
}

function normalizeProfessorSheetName(name) {
  const value = String(name || "").trim().toUpperCase();

  if (value === "2026") return "2026";
  if (value === "2025") return "2025";
  if (value === "ASCENSO") return "ASCENSO";
  if (value === "DATA") return "DATA";
  if (value === "CONTRATOS") return "CONTRATOS";

  return null;
}

function isHeaderLikeRow(numero, nombre, asunto) {
  const a = String(numero || "").trim().toLowerCase();
  const b = String(nombre || "").trim().toLowerCase();
  const c = String(asunto || "").trim().toLowerCase();

  return (
    a.includes("numero") ||
    a.includes("número") ||
    b.includes("nombre") ||
    b.includes("docente") ||
    c.includes("asunto") ||
    c.includes("detalle")
  );
}

/* ========================= ESTADÍSTICAS ========================= */
function updateStats() {
  const totalUgel = getModuleTotal("RD UGEL");
  const totalExpedientes = getModuleTotal("RD Expedientes");
  const totalProfesores = getModuleTotal("RD Profesores");
  const pendingLoans = loans.filter((l) => l.status === "pendiente").length;

  const statUgel = document.getElementById("statUgel");
  const statExpedientes = document.getElementById("statExpedientes");
  const statProfesores = document.getElementById("statProfesores");
  const statPendientes = document.getElementById("statPendientes");

  if (statUgel) statUgel.textContent = totalUgel;
  if (statExpedientes) statExpedientes.textContent = totalExpedientes;
  if (statProfesores) statProfesores.textContent = totalProfesores;
  if (statPendientes) statPendientes.textContent = pendingLoans;

  updateLoanMiniStats();
}

function getModuleTotal(moduleName) {
  if (moduleName === "RD Profesores") {
    return PROFESSOR_SHEETS.reduce((acc, sheetName) => {
      return acc + (rdData["RD Profesores"][sheetName]?.items?.length || 0);
    }, 0);
  }

  return YEARS.reduce((acc, year) => acc + getYearCount(moduleName, year), 0);
}

/* ========================= UTILIDADES ========================= */
function formatDateTime(isoString) {
  const date = new Date(isoString);

  return date.toLocaleString("es-PE", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });
}

function showToast(message) {
  const toast = document.getElementById("toast");
  if (!toast) return;

  toast.textContent = message;
  toast.classList.add("show");

  clearTimeout(showToast._timer);
  showToast._timer = setTimeout(() => {
    toast.classList.remove("show");
  }, 3200);
}

function sanitizeKey(text) {
  return text.toLowerCase().replace(/\s+/g, "-");
}

function escapeHTML(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

/* ========================= SEGURIDAD BÁSICA DE VISTA ========================= */
function requireAuthAction(callback) {
  if (!auth.currentUser) {
    showLogin();
    showLoginError("Debes iniciar sesión para continuar.");
    return;
  }

  callback();
}

/* ========================= EVENTOS GLOBALES ========================= */
document.addEventListener("click", function (e) {
  const toggleProfBtn = e.target.closest(".toggle-prof-btn");

  if (toggleProfBtn) {
    const sheetName = toggleProfBtn.dataset.sheet;
    if (!sheetName) return;

    expandedProfessorSheets[sheetName] = !expandedProfessorSheets[sheetName];
    renderProfessorsModule(getSearchValueByModule("RD Profesores"));
    return;
  }

  const toggleBtn = e.target.closest(".toggle-rd-btn");

  if (toggleBtn && !toggleBtn.classList.contains("toggle-prof-btn")) {
    const moduleName = toggleBtn.dataset.module;
    const year = Number(toggleBtn.dataset.year);

    if (!moduleName || !year) return;

    expandedYears[moduleName][year] = !expandedYears[moduleName][year];

    if (expandedYears[moduleName][year]) {
      ensureYearLoaded(moduleName, year);
    }

    renderStandardModule(
      moduleName,
      getContainerIdByModule(moduleName),
      getSearchValueByModule(moduleName)
    );

    return;
  }

  const profBtn = e.target.closest("[data-prof-id]");

  if (profBtn) {
    const id = profBtn.dataset.profId;
    const sheetName = profBtn.dataset.profSheet;
    const newStatus = profBtn.dataset.profStatus;

    if (!id || !sheetName || !newStatus) return;

    updateProfessorStatus(id, sheetName, newStatus);
  }
});

/* ========================= INICIO ========================= */
document.addEventListener("DOMContentLoaded", () => {
  showLogin();
  setupAuth();
});
