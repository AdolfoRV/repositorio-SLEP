// assets/slep/pyodide-loader.js
// Carga Pyodide, instala dependencias y sube el paquete Python `slep`.

import { SLEP_MODULES, SLEP_PY_BASE, PYODIDE_SCRIPT_URL } from "./config.js";
import { log } from "./log.js";

let pyodideInstance = null;

function loadPyodideScript() {
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = PYODIDE_SCRIPT_URL;
    script.onload = resolve;
    script.onerror = () => reject(new Error("No se pudo cargar pyodide.js"));
    document.head.appendChild(script);
  });
}

async function fetchSlepModules(pyodide) {
  log("Descargando modulos del migrador...");
  pyodide.FS.mkdirTree("/home/pyodide/slep");

  for (const modulo of SLEP_MODULES) {
    const url = new URL(modulo, SLEP_PY_BASE).href;
    const res = await fetch(url);
    if (!res.ok) {
      throw new Error(`No se pudo cargar ${modulo} (HTTP ${res.status}) desde ${url}`);
    }
    const texto = await res.text();
    pyodide.FS.writeFile(`/home/pyodide/slep/${modulo}`, texto);
  }

  pyodide.runPython(`
import sys
if "/home/pyodide" not in sys.path:
    sys.path.insert(0, "/home/pyodide")
import slep
  `);
}

// Devuelve la instancia de Pyodide, cargándola y preparándola solo la primera vez.
export async function getPyodide() {
  if (pyodideInstance) return pyodideInstance;

  log("Cargando Python en el navegador...");

  if (typeof loadPyodide !== "function") {
    await loadPyodideScript();
  }

  const pyodide = await loadPyodide();

  await pyodide.loadPackage("micropip");
  const micropip = pyodide.pyimport("micropip");

  log("Instalando openpyxl (~20-30 segundos)...");
  await micropip.install("openpyxl");

  await fetchSlepModules(pyodide);

  pyodideInstance = pyodide;
  return pyodide;
}
