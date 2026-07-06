// assets/slep/processor.js
// Orquesta el flujo completo: leer archivos, correr slep.procesar y descargar resultados.

import { ESTABLECIMIENTOS_URL, XLSX_MIME, ZIP_MIME } from "./config.js";
import { log, clearLog } from "./log.js";
import { getPyodide } from "./pyodide-loader.js";
import { descargarBlob } from "./download.js";
import { files } from "./upload.js";

async function cargarEstablecimientos() {
  log("Cargando maestro de establecimientos...");
  log("(URL: " + ESTABLECIMIENTOS_URL + ")");
  const estRes = await fetch(ESTABLECIMIENTOS_URL);
  if (!estRes.ok) {
    throw new Error(
      `No se pudo cargar el maestro de establecimientos (${ESTABLECIMIENTOS_URL}) — HTTP ${estRes.status}`
    );
  }
  return estRes.arrayBuffer();
}

function descargarResultados(resultados) {
  for (const [nombre, data] of resultados) {
    if (nombre === "SLEP_files.zip") continue;
    descargarBlob(nombre, data, XLSX_MIME);
  }
  descargarBlob("SLEP_files.zip", resultados.get("SLEP_files.zip"), ZIP_MIME);
}

export async function procesarArchivos() {
  const btn = document.getElementById("btn-procesar");
  const downloadsPanel = document.getElementById("downloads");
  const downloadList = document.getElementById("download-list");

  btn.disabled = true;
  btn.textContent = "Iniciando...";
  clearLog();
  downloadsPanel.style.display = "none";
  downloadList.innerHTML = "";

  try {
    const pyodide = await getPyodide();

    log("Leyendo archivo de licencias...");
    const licenciasBuf = await files.licencias.arrayBuffer();

    const estBuf = await cargarEstablecimientos();

    pyodide.globals.set("licencias_bytes", new Uint8Array(licenciasBuf));
    pyodide.globals.set("establecimientos_bytes", new Uint8Array(estBuf));

    log("Procesando datos...");
    pyodide.runPython(`
resultados = slep.procesar(bytes(licencias_bytes), bytes(establecimientos_bytes))
    `);

    const resultados = pyodide.globals.get("resultados").toJs();
    descargarResultados(resultados);

    downloadsPanel.style.display = "block";
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
