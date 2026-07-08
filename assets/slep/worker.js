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

self.onerror = function(event) {
  self.postMessage({
    type: "error",
    msg: "Worker error: " + safeErrorMsg(event.error || event.message),
  });
};

self.postMessage({ type: "ready" });

self.onmessage = async function(e) {
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
      throw new Error("loadPyodide no está disponible después de importScripts.");
    }

    const pyodide = await loadPyodide({ indexURL: pyodideIndexUrl });

    self.postMessage({
      type: "log",
      msg: "Instalando openpyxl (~20-30 segundos)...",
    });
    await pyodide.loadPackage("micropip");
    const micropip = pyodide.pyimport("micropip");
    await micropip.install("openpyxl");

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

    pyodide.runPython(`
import sys
if "/home/pyodide" not in sys.path:
    sys.path.insert(0, "/home/pyodide")
import slep
    `);

    self.postMessage({ type: "log", msg: "Descargando plantilla Power BI..." });
    const pbitRes = await fetch(pbit_url);
    if (!pbitRes.ok) {
      throw new Error(
        `No se pudo cargar la plantilla Power BI (${pbit_url}) — HTTP ${pbitRes.status}`
      );
    }
    const pbitBuf = await pbitRes.arrayBuffer();

    self.postMessage({ type: "log", msg: "Procesando datos..." });

    pyodide.globals.set("slep_log", (msg) => {
      self.postMessage({ type: "log", msg: String(msg) });
    });

    pyodide.globals.set("licencias_bytes", licencias_bytes);
    pyodide.globals.set("establecimientos_bytes", establecimientos_bytes);
    pyodide.globals.set("pbit_bytes", new Uint8Array(pbitBuf));

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

    const resultados = pyodide.globals.get("resultados_list").toJs();

    if (!Array.isArray(resultados)) {
      throw new Error(`resultados no es un array, es ${typeof resultados}`);
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
          `Formato inesperado. Primera entrada: tipo=${typeof first}, len=${first?.length}, nombre_tipo=${typeof first?.[0]}, data_tipo=${first?.[1]?.constructor?.name}`
        );
      }
    }

    self.postMessage({ type: "done", resultados });
  } catch (err) {
    self.postMessage({ type: "error", msg: safeErrorMsg(err) });
  }
};