/**
 * @fileoverview Web Worker para el pipeline de migración SLEP.
 *
 * Este worker corre en un hilo independiente del navegador y es responsable de:
 *
 * 1. Cargar el runtime de Pyodide (Python compilado a WebAssembly).
 * 2. Instalar la dependencia `openpyxl` vía `micropip`.
 * 3. Descargar los módulos Python del paquete `slep` desde el sitio estático
 *    y montarlos en el filesystem virtual de Pyodide (`/home/pyodide/slep/`).
 * 4. Descargar la plantilla Power BI (`.pbit`) desde el sitio estático.
 * 5. Ejecutar `slep.procesar(...)` con los bytes de entrada y capturar el log
 *    del pipeline para reenviarlo a la UI principal.
 * 6. Convertir los resultados del diccionario Python a un array JavaScript
 *    tipado y enviarlos al hilo principal mediante `postMessage`.
 *
 * **Protocolo de mensajes** (entre worker y hilo principal):
 *
 * | Dirección | Tipo     | Payload                        | Significado                           |
 * |-----------|----------|--------------------------------|---------------------------------------|
 * | → Worker  | -        | `{ licencias_bytes, ... }`     | Inicio del procesamiento              |
 * | ← Main    | `ready`  | -                              | Worker inicializado y escuchando      |
 * | ← Main    | `log`    | `{ msg: string }`              | Línea de log del pipeline             |
 * | ← Main    | `error`  | `{ msg: string }`              | Error irrecuperable                   |
 * | ← Main    | `done`   | `{ resultados: Array }`        | Procesamiento exitoso, con archivos   |
 *
 * @module worker
 */

//  Utilidades de serialización de errores

/**
 * Normaliza cualquier valor de error a una cadena legible y segura para
 * `postMessage`.
 *
 * Maneja casos edge case como `null`, `undefined`, objetos sin `message`,
 * objetos que no son serializables con `JSON.stringify`, y el string
 * `"[object Object]"`.
 *
 * @function safeErrorMsg
 * @param {*} err - Valor capturado en un bloque `catch` o en `self.onerror`.
 * @returns {string} Representación textual segura del error.
 *
 * @example
 * safeErrorMsg(null);                 // "Error desconocido (null/undefined)"
 * safeErrorMsg(new Error("foo"));     // "foo"
 * safeErrorMsg({ custom: true });     // '{"custom":true}'
 */
function safeErrorMsg(err) {
  if (err == null) return "Error desconocido (null/undefined)";
  if (typeof err === "string") return err;
  if (err.message) return String(err.message);
  if (err.toString && err.toString() !== "[object Object]") return err.toString();
  try {
    return JSON.stringify(err);
  } catch {
    return "Error no serializable";
  }
}

//  Manejadores globales de errores del worker

/**
 * Handler global para errores no capturados dentro del worker.
 *
 * Se dispara cuando una excepción escapa de cualquier `try/catch` o cuando
 * ocurre un error de runtime del worker. Reenvía el error al hilo principal
 * con tipo `"error"` para que la UI lo muestre al usuario.
 *
 * @event self#onerror
 * @param {ErrorEvent} event
 * @sideEffect Envía `postMessage({ type: "error", msg: ... })`.
 */
self.onerror = function (event) {
  self.postMessage({
    type: "error",
    msg: "Worker error: " + safeErrorMsg(event.error || event.message),
  });
};

// Notificar al hilo principal que el worker está listo para recibir mensajes.
self.postMessage({ type: "ready" });

//  Handler principal del mensaje de inicio

/**
 * Punto de entrada del procesamiento. Recibe los datos binarios del archivo
 * de licencias, la URL del maestro de establecimientos (ya convertido a bytes
 * por el hilo principal), la URL de la plantilla Power BI, y los metadatos
 * necesarios para montar el entorno Pyodide.
 *
 * @event self#onmessage
 * @param {MessageEvent} e
 * @param {Object} e.data
 * @param {Uint8Array} e.data.licencias_bytes - Bytes de la planilla madre de licencias.
 * @param {Uint8Array} e.data.establecimientos_bytes - Bytes del maestro de establecimientos.
 * @param {string} e.data.pbit_url - URL absoluta de la plantilla `.pbit`.
 * @param {string} e.data.siteRoot - URL base del sitio (para resolver rutas relativas de módulos).
 * @param {string[]} e.data.slepModules - Lista de nombres de archivo del paquete `slep`
 *   (ej. `["__init__.py", "constants.py", "utils.py", "core.py"]`).
 * @param {string} e.data.pyodideIndexUrl - URL base del índice de Pyodide
 *   (ej. `"https://cdn.jsdelivr.net/pyodide/v0.26.2/full/"`).
 *
 * @returns {Promise<void>}
 * @async
 * @sideEffect
 *   - Carga Pyodide, `openpyxl` y los módulos `slep` en memoria.
 *   - Ejecuta `slep.procesar(...)`.
 *   - Envía múltiples mensajes `log` al hilo principal.
 *   - Al finalizar, envía `done` con los resultados o `error` si falla.
 */
self.onmessage = async function (e) {
  const {
    licencias_bytes,
    establecimientos_bytes,
    pbit_url,
    siteRoot,
    slepModules,
    pyodideIndexUrl,
  } = e.data;

  const pyodideScriptUrl = pyodideIndexUrl + "pyodide.js";

  try {
    // Carga de Pyodide
    self.postMessage({ type: "log", msg: "Cargando Pyodide..." });

    if (typeof loadPyodide !== "function") {
      self.postMessage({
        type: "log",
        msg: "Descargando pyodide.js vía importScripts...",
      });
      try {
        self.importScripts(pyodideScriptUrl);
      } catch (importErr) {
        throw new Error("importScripts falló: " + safeErrorMsg(importErr));
      }
    }

    if (typeof loadPyodide !== "function") {
      throw new Error(
        "loadPyodide no está disponible después de importScripts."
      );
    }

    const pyodide = await loadPyodide({ indexURL: pyodideIndexUrl });

    // Instalación de dependencias Python
    self.postMessage({
      type: "log",
      msg: "Instalando openpyxl (~20-30 segundos)...",
    });
    await pyodide.loadPackage("micropip");
    const micropip = pyodide.pyimport("micropip");
    await micropip.install("openpyxl");

    // Descarga y montaje de los módulos del paquete slep
    self.postMessage({ type: "log", msg: "Descargando módulos del migrador..." });

    pyodide.FS.mkdirTree("/home/pyodide/slep");
    const slepBase = new URL("scripts/slep/", siteRoot).href;

    for (const mod of slepModules) {
      const url = new URL(mod, slepBase).href;
      const res = await fetch(url);
      if (!res.ok) {
        throw new Error(
          `No se pudo cargar ${mod} (HTTP ${res.status}) desde ${url}`
        );
      }
      const text = await res.text();
      pyodide.FS.writeFile(`/home/pyodide/slep/${mod}`, text);
    }

    // Inicializar el paquete Python en el path de Pyodide.
    pyodide.runPython(`
import sys
if "/home/pyodide" not in sys.path:
    sys.path.insert(0, "/home/pyodide")
import slep
    `);

    // Descarga de la plantilla Power BI
    self.postMessage({ type: "log", msg: "Descargando plantilla Power BI..." });
    const pbitRes = await fetch(pbit_url);
    if (!pbitRes.ok) {
      throw new Error(
        `No se pudo cargar la plantilla Power BI (${pbit_url}) - HTTP ${pbitRes.status}`
      );
    }
    const pbitBuf = await pbitRes.arrayBuffer();

    //  Ejecución del pipeline slep.procesar(...)
    self.postMessage({ type: "log", msg: "Procesando datos..." });

    // Callback de log: cada línea emitida por el pipeline Python se reenvía
    // al hilo principal para mostrarla en tiempo real en la UI.
    pyodide.globals.set("slep_log", (msg) => {
      self.postMessage({ type: "log", msg: String(msg) });
    });

    // Transferir los bytes de entrada al entorno Python.
    pyodide.globals.set("licencias_bytes", licencias_bytes);
    pyodide.globals.set("establecimientos_bytes", establecimientos_bytes);
    pyodide.globals.set("pbit_bytes", new Uint8Array(pbitBuf));

    // Invocar la función orquestadora del paquete slep.
    pyodide.runPython(`
resultados = slep.procesar(
    bytes(licencias_bytes),
    bytes(establecimientos_bytes),
    slep_log,
    pbit_data=bytes(pbit_bytes)
)

resultados_list = []
for k, v in resultados.items():
    name = str(k)
    data = bytes(v) if isinstance(v, (bytes, bytearray)) else bytes(v)
    resultados_list.append((name, data))
    `);

    // Validación y conversión de resultados
    const resultados = pyodide.globals.get("resultados_list").toJs();

    if (!Array.isArray(resultados)) {
      throw new Error(
        `resultados no es un array, es ${typeof resultados}`
      );
    }
    if (resultados.length > 0) {
      const first = resultados[0];
      if (
        !Array.isArray(first) ||
        first.length !== 2 ||
        typeof first[0] !== "string" ||
        !(first[1] instanceof Uint8Array)
      ) {
        throw new Error(
          `Formato inesperado. Primera entrada: tipo=${typeof first}, ` +
          `len=${first?.length}, nombre_tipo=${typeof first?.[0]}, ` +
          `data_tipo=${first?.[1]?.constructor?.name}`
        );
      }
    }

    // Entrega exitosa
    self.postMessage({ type: "done", resultados });
  } catch (err) {
    // Cualquier excepción en los pasos anteriores se normaliza y se envía
    // al hilo principal como error irrecuperable.
    self.postMessage({ type: "error", msg: safeErrorMsg(err) });
  }
};
