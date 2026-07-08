// assets/slep/processor.js

import { log, clearLog, descargarBlob, files } from "./ui.js";

const WORKER_TIMEOUT_MS = 120000;
const XLSX_MIME =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const ZIP_MIME = "application/zip";
const PBIT_MIME = "application/octet-stream";

function descargarResultados(resultados) {
  for (const [nombre, data] of resultados) {
    if (nombre === "SLEP_files.zip") continue;
    const mime = nombre.endsWith(".pbit") ? PBIT_MIME : XLSX_MIME;
    descargarBlob(nombre, data, mime);
  }
  const zipEntry = resultados.find(([n]) => n === "SLEP_files.zip");
  if (zipEntry) {
    descargarBlob("SLEP_files.zip", zipEntry[1], ZIP_MIME);
  }
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

  let worker = null;
  let timeoutId = null;

  const cleanup = () => {
    if (timeoutId) clearTimeout(timeoutId);
    if (worker) {
      worker.terminate();
      worker = null;
    }
  };

  try {
    log("Leyendo archivo de licencias...");
    const licenciasBuf = await files.licencias.arrayBuffer();

    log("Cargando maestro de establecimientos...");
    const siteRoot = window.location.href.substring(
      0,
      window.location.href.lastIndexOf("/") + 1
    );
    const estUrl = new URL("assets/tables/establecimientos.xlsx", siteRoot).href;
    const estRes = await fetch(estUrl);
    if (!estRes.ok) {
      throw new Error(
        `No se pudo cargar el maestro de establecimientos (${estUrl}) — HTTP ${estRes.status}`
      );
    }
    const estBuf = await estRes.arrayBuffer();

    // URL del .pbit para que el worker lo descargue directamente
    const pbitUrl = new URL("assets/Licencias_Medicas.pbit", siteRoot).href;

    log("Iniciando procesamiento en segundo plano...");
    worker = new Worker("assets/slep/worker.js");

    timeoutId = setTimeout(() => {
      log("ERROR: El worker no respondió dentro del tiempo límite (2 min).", true);
      btn.textContent = "Reintentar";
      btn.disabled = false;
      cleanup();
    }, WORKER_TIMEOUT_MS);

    worker.onmessage = (e) => {
      const { type, msg, resultados } = e.data;

      if (type === "ready") {
        log("Worker listo.");
        return;
      }
      if (type === "log") {
        log(msg);
        return;
      }
      if (type === "error") {
        log("ERROR: " + msg, true);
        console.error(msg);
        btn.textContent = "Reintentar";
        btn.disabled = false;
        cleanup();
        return;
      }
      if (type === "done") {
        clearTimeout(timeoutId);
        timeoutId = null;
        descargarResultados(resultados);
        downloadsPanel.style.display = "block";
        btn.textContent = "Procesar y descargar";
        btn.disabled = false;
        cleanup();
        return;
      }
    };

    worker.onerror = (err) => {
      log("ERROR del Worker: " + err.message, true);
      console.error(err);
      btn.textContent = "Reintentar";
      btn.disabled = false;
      cleanup();
    };

    worker.onmessageerror = (err) => {
      log("ERROR de mensaje del Worker: " + err.message, true);
      console.error(err);
      btn.textContent = "Reintentar";
      btn.disabled = false;
      cleanup();
    };

    worker.postMessage({
      licencias_bytes: new Uint8Array(licenciasBuf),
      establecimientos_bytes: new Uint8Array(estBuf),
      pbit_url: pbitUrl,           // ← URL en vez de bytes
      siteRoot,
      slepModules: ["__init__.py", "constants.py", "utils.py", "core.py"],
      pyodideIndexUrl: "https://cdn.jsdelivr.net/pyodide/v0.26.2/full/",
    });
  } catch (e) {
    log("ERROR: " + e.message, true);
    console.error(e);
    btn.textContent = "Reintentar";
    btn.disabled = false;
    cleanup();
  }
}