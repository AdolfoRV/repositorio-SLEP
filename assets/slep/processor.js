/**
 * @fileoverview Orquestador del pipeline de migración SLEP.
 *
 * Coordina la comunicación entre la interfaz de usuario (UI) y el
 * {@link Worker} que ejecuta el motor Python (Pyodide/WASM). Su responsabilidad
 * es:
 *
 * 1. Leer el archivo de licencias seleccionado por el usuario.
 * 2. Descargar vía `fetch()` los recursos estáticos del sitio (maestro de
 *    establecimientos y plantilla Power BI).
 * 3. Instanciar y configurar el {@link Worker} con los datos binarios y las
 *    URLs necesarias.
 * 4. Gestionar el ciclo de vida del worker (timeout, terminación, manejo de
 *    errores) y reaccionar a los mensajes recibidos (log, error, done).
 * 5. Al finalizar exitosamente, disparar la descarga de cada resultado
 *    individual y del ZIP consolidado.
 *
 * @module processor
 * @see {@link module:ui} - Módulo de interfaz de usuario que provee `log`,
 *      `clearLog`, `descargarBlob` y el estado `files`.
 */

import { log, clearLog, descargarBlob, files } from "./ui.js";

//  Constantes de configuración

/**
 * Tiempo máximo de espera (ms) antes de considerar que el worker no respondió.
 */
const WORKER_TIMEOUT_MS = 120000;

/**
 * MIME type para archivos Excel (.xlsx).
 */
const XLSX_MIME =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/**
 * MIME type para archivos ZIP.
 */
const ZIP_MIME = "application/zip";

/**
 * MIME type genérico para archivos binarios no reconocidos (usado para `.pbit`).
 */
const PBIT_MIME = "application/octet-stream";

//  Funciones auxiliares

/**
 * Itera sobre los resultados devueltos por el worker y genera los enlaces de
 * descarga correspondientes en el DOM.
 *
 * Descarga primero cada archivo individual (excepto el ZIP) y, al final,
 * el archivo `SLEP_files.zip` que contiene el paquete completo. El orden
 * asegura que el usuario vea inmediatamente los archivos desglosados y luego
 * el consolidado.
 *
 * @function descargarResultados
 * @param {Array<[string, Uint8Array]>} resultados - Lista de tuplas
 *   `[nombreArchivo, datosBinarios]` devuelta por el worker al finalizar.
 * @returns {void}
 * @sideEffect Llama a {@link module:ui.descargarBlob} por cada archivo,
 *   generando elementos `<a>` en el DOM.
 *
 * @example
 * descargarResultados([
 *   ["01_Dim_Funcionario.xlsx", uint8Array1],
 *   ["SLEP_files.zip", uint8ArrayZip]
 * ]);
 */
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

//  Función principal de orquestación

/**
 * Punto de entrada principal del procesamiento. Se invoca típicamente desde
 * el evento `click` del botón "Procesar y descargar" en la interfaz Quarto.
 *
 * Ejecuta el siguiente flujo:
 *
 * 1. **Preparación de UI**: deshabilita el botón, limpia logs y oculta el
 *    panel de descargas previo.
 * 2. **Lectura de archivos**: convierte el `File` de licencias del usuario a
 *    `ArrayBuffer`.
 * 3. **Carga de recursos estáticos**: descarga mediante `fetch()` el maestro
 *    de establecimientos (`assets/tables/establecimientos.xlsx`) y la
 *    plantilla Power BI (`assets/Dashboard_Licencias.pbit`).
 * 4. **Instanciación del worker**: crea un {@link Worker} desde
 *    `assets/slep/worker.js` y configura listeners para los mensajes
 *    `ready`, `log`, `error` y `done`.
 * 5. **Timeout de seguridad**: si el worker no envía `done` dentro de
 *    {@link WORKER_TIMEOUT_MS}, se considera fallo y se ofrece reintentar.
 * 6. **Post-procesamiento**: al recibir `done`, descarga los resultados,
 *    habilita el botón y limpia recursos.
 *
 * Cualquier excepción en los pasos 2 o 3 (lectura de archivo, `fetch` fallido,
 * etc.) se captura, se muestra en el log de errores y se restaura el botón.
 *
 * @function procesarArchivos
 * @returns {Promise<void>}
 * @async
 * @throws {Error} Propagada internamente; nunca escapa al caller porque se
 *   captura en el bloque `catch` y se transforma en un mensaje de log.
 * @sideEffect
 *   - Modifica el DOM (`#btn-procesar`, `#downloads`, `#download-list`, `#log`).
 *   - Crea y termina un {@link Worker}.
 *   - Genera object URLs para descargas.
 *
 * @example
 * document.getElementById("btn-procesar")
 *   .addEventListener("click", procesarArchivos);
 */
export async function procesarArchivos() {
  // Referencias DOM
  const btn = document.getElementById("btn-procesar");
  const downloadsPanel = document.getElementById("downloads");
  const downloadList = document.getElementById("download-list");

  // Reset de UI
  btn.disabled = true;
  btn.textContent = "Iniciando...";
  clearLog();
  downloadsPanel.style.display = "none";
  downloadList.innerHTML = "";

  let worker = null;
  let timeoutId = null;

  /**
   * Libera recursos del worker y del timeout de forma idempotente.
   *
   * @function cleanup
   * @private
   * @returns {void}
   */
  const cleanup = () => {
    if (timeoutId) clearTimeout(timeoutId);
    if (worker) {
      worker.terminate();
      worker = null;
    }
  };

  try {
    // Lectura del archivo de licencias
    log("Leyendo archivo de licencias...");
    const licenciasBuf = await files.licencias.arrayBuffer();

    // Carga del maestro de establecimientos (recurso estático)
    log("Cargando maestro de establecimientos...");
    const siteRoot = window.location.href.substring(
      0,
      window.location.href.lastIndexOf("/") + 1
    );

    const estUrl = new URL("assets/tables/establecimientos.xlsx", siteRoot).href;
    const estRes = await fetch(estUrl);
    if (!estRes.ok) {
      throw new Error(
        `No se pudo cargar el maestro de establecimientos (${estUrl}) - HTTP ${estRes.status}`
      );
    }
    const estBuf = await estRes.arrayBuffer();

    // URL de la plantilla Power BI (el worker la descarga directamente)
    const pbitUrl = new URL("assets/Dashboard_Licencias.pbit", siteRoot).href;

    // Instanciación del Web Worker
    log("Iniciando procesamiento en segundo plano...");
    worker = new Worker("assets/slep/worker.js");

    // Timeout de seguridad 
    timeoutId = setTimeout(() => {
      log(
        "ERROR: El worker no respondió dentro del tiempo límite (2 min).",
        true
      );
      btn.textContent = "Reintentar";
      btn.disabled = false;
      cleanup();
    }, WORKER_TIMEOUT_MS);

    // Handlers de mensajes del worker 
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

    // Envío de datos al worker
    worker.postMessage({
      licencias_bytes: new Uint8Array(licenciasBuf),
      establecimientos_bytes: new Uint8Array(estBuf),
      pbit_url: pbitUrl,           // URL en vez de bytes para reducir payload inicial
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