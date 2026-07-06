// app.js - Interfaz del Migrador SLEP.
// La logica de negocio vive en scripts/slep/*.py (paquete Python real).

const SLEP_MODULES = [
  "__init__.py",
  "constants.py",
  "text_utils.py",
  "core.py",
];

const SLEP_PY_BASE = "scripts/slep";
const ESTABLECIMIENTOS_URL = "../assets/tables/establecimientos.xlsx";

let files = { licencias: null };
let pyodideInstance = null;

function log(msg, isError = false) {
  const el = document.getElementById("log");
  el.style.display = "block";
  const line = document.createElement("div");
  line.textContent = msg;
  if (isError) line.className = "err";
  el.appendChild(line);
  el.scrollTop = el.scrollHeight;
}

function updateBtn() {
  const ready = !!files.licencias;
  const btn = document.getElementById("btn-procesar");
  btn.disabled = !ready;
  btn.textContent = ready ? "Procesar y descargar" : "Selecciona el archivo de licencias";
}

function setupDrop(boxId, key) {
  const box = document.getElementById(boxId);
  const input = box.querySelector('input[type="file"]');
  const tag = box.querySelector('div[id^="tag"]');

  const setFile = (file) => {
    files[key] = file;
    tag.innerHTML = `<span class="file-tag">${file.name}</span>`;
    updateBtn();
  };

  input.addEventListener("change", (e) => {
    if (e.target.files.length) setFile(e.target.files[0]);
  });

  box.addEventListener("dragover", (e) => {
    e.preventDefault();
    box.classList.add("active");
  });
  box.addEventListener("dragleave", () => box.classList.remove("active"));
  box.addEventListener("drop", (e) => {
    e.preventDefault();
    box.classList.remove("active");
    if (e.dataTransfer.files.length) setFile(e.dataTransfer.files[0]);
  });
}

// Carga Pyodide via script tag (evita bug de require.js con pyodide.mjs)
async function getPyodide() {
  if (pyodideInstance) return pyodideInstance;

  log("Cargando Python en el navegador...");

  // Si ya está cargado (por ejemplo, por otro script), reusar
  if (typeof loadPyodide === "function") {
    pyodideInstance = await loadPyodide();
    return pyodideInstance;
  }

  await new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "https://cdn.jsdelivr.net/pyodide/v0.26.2/full/pyodide.js";
    script.onload = resolve;
    script.onerror = () => reject(new Error("No se pudo cargar pyodide.js"));
    document.head.appendChild(script);
  });

  const pyodide = await loadPyodide();

  await pyodide.loadPackage("micropip");
  const micropip = pyodide.pyimport("micropip");

  log("Instalando openpyxl (~20-30 segundos)...");
  await micropip.install("openpyxl");

  log("Descargando modulos del migrador...");
  pyodide.FS.mkdirTree("/home/pyodide/slep");
  for (const modulo of SLEP_MODULES) {
    const res = await fetch(`${SLEP_PY_BASE}/${modulo}`);
    if (!res.ok) throw new Error(`No se pudo cargar ${modulo} (HTTP ${res.status})`);
    const texto = await res.text();
    pyodide.FS.writeFile(`/home/pyodide/slep/${modulo}`, texto);
  }

  pyodide.runPython(`
import sys
if "/home/pyodide" not in sys.path:
    sys.path.insert(0, "/home/pyodide")
import slep
  `);

  pyodideInstance = pyodide;
  return pyodide;
}

function descargarBlob(nombre, data, tipo) {
  const blob = new Blob([data], { type: tipo });
  const url = URL.createObjectURL(blob);
  const li = document.createElement("li");
  const a = document.createElement("a");
  a.href = url;
  a.download = nombre;
  a.textContent = nombre;
  a.style.fontWeight = "bold";
  li.appendChild(a);
  document.getElementById("download-list").appendChild(li);
}

async function procesarArchivos() {
  const btn = document.getElementById("btn-procesar");
  btn.disabled = true;
  btn.textContent = "Iniciando...";
  document.getElementById("log").innerHTML = "";
  document.getElementById("downloads").style.display = "none";
  document.getElementById("download-list").innerHTML = "";

  try {
    const pyodide = await getPyodide();

    log("Leyendo archivo de licencias...");
    const licenciasBuf = await files.licencias.arrayBuffer();

    log("Cargando maestro de establecimientos...");
    const estRes = await fetch(ESTABLECIMIENTOS_URL);
    if (!estRes.ok) {
      throw new Error(`No se pudo cargar el maestro de establecimientos (${ESTABLECIMIENTOS_URL})`);
    }
    const estBuf = await estRes.arrayBuffer();

    pyodide.globals.set("licencias_bytes", new Uint8Array(licenciasBuf));
    pyodide.globals.set("establecimientos_bytes", new Uint8Array(estBuf));

    log("Procesando datos...");
    pyodide.runPython(`
resultados = slep.procesar(bytes(licencias_bytes), bytes(establecimientos_bytes))
    `);

    const resultados = pyodide.globals.get("resultados").toJs();
    const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    for (const [nombre, data] of resultados) {
      if (nombre === "SLEP_files.zip") continue;
      descargarBlob(nombre, data, XLSX_MIME);
    }
    descargarBlob("SLEP_files.zip", resultados.get("SLEP_files.zip"), "application/zip");

    document.getElementById("downloads").style.display = "block";
    log("Listo! Descarga los archivos arriba.");
    btn.textContent = "Procesar y descargar";
    btn.disabled = false;
  } catch (e) {
    log("ERROR: " + e.message, true);
    console.error(e);
    btn.textContent = "Reintentar";
    btn.disabled = false;
  }
}

export function initSlepApp() {
  setupDrop("drop-licencias", "licencias");
  document.getElementById("btn-procesar").addEventListener("click", procesarArchivos);
}