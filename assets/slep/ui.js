/**
 * @fileoverview Módulo de interfaz de usuario (UI) para el migrador SLEP.
 *
 * Gestiona todo lo relacionado con el DOM: carga de archivos mediante drag-and-drop
 * o selector nativo, visualización de logs de procesamiento, y generación de
 * enlaces de descarga para los resultados del pipeline.
 *
 * Este módulo es puramente presentacional; no contiene lógica de negocio ni
 * comunicación con el worker. Su responsabilidad única es mediar entre el usuario
 * y el resto de la aplicación.
 *
 * @module ui
 */

//  Estado global de archivos

/**
 * Estado mutable de los archivos seleccionados por el usuario.
 *
 * @type {{ licencias: (File|null) }}
 * @description Solo se expone la propiedad `licencias` porque el maestro de
 * establecimientos y la plantilla Power BI se cargan automáticamente desde
 * `assets/` mediante `fetch()` en el módulo {@link processor}.
 */
export const files = {
  licencias: null,
};

//  Controles de la interfaz

/**
 * Actualiza el estado del botón principal de procesamiento en función de si
 * el usuario ha seleccionado un archivo de licencias.
 *
 * Si el archivo está presente, habilita el botón y cambia el texto a
 * "Procesar y descargar". En caso contrario, lo deshabilita y solicita la
 * selección del archivo.
 *
 * @function updateBtn
 * @returns {void}
 * @sideEffect Modifica `disabled` y `textContent` del elemento `#btn-procesar`.
 */
export function updateBtn() {
  const ready = !!files.licencias;
  const btn = document.getElementById("btn-procesar");
  if (!btn) return;
  btn.disabled = !ready;
  btn.textContent = ready
    ? "Procesar y descargar"
    : "Selecciona el archivo de licencias";
}

/**
 * Configura un área de drop (drag-and-drop) para recibir archivos.
 *
 * Vincula el elemento contenedor identificado por `boxId` con:
 * - El input nativo `<input type="file">` (evento `change`).
 * - Eventos de drag-and-drop (`dragover`, `dragleave`, `drop`).
 *
 * Al recibir un archivo, lo almacena en {@link files} bajo la clave indicada
 * y actualiza la etiqueta visual correspondiente.
 *
 * @function setupDrop
 * @param {string} boxId - Identificador del contenedor DOM (ej. `"drop-licencias"`).
 * @param {string} key   - Clave dentro de {@link files} donde se almacenará el archivo (ej. `"licencias"`).
 * @returns {void}
 * @sideEffect
 *   - Escribe en `files[key]`.
 *   - Modifica el HTML del elemento cuyo `id` comienza con `"tag"`.
 *   - Llama a {@link updateBtn} para refrescar el estado del botón.
 *
 * @example
 * setupDrop("drop-licencias", "licencias");
 */
export function setupDrop(boxId, key) {
  const box = document.getElementById(boxId);
  if (!box) return;

  const input = box.querySelector('input[type="file"]');
  const tag = box.querySelector('div[id^="tag"]');

  /**
   * Almacena el archivo recibido, actualiza la etiqueta visual y refresca
   * el estado del botón principal.
   *
   * @param {File} file - Archivo seleccionado por el usuario.
   * @private
   */
  const setFile = (file) => {
    files[key] = file;
    if (tag) tag.innerHTML = `<span class="file-tag">${file.name}</span>`;
    updateBtn();
  };

  // Input nativo 
  input.addEventListener("change", (e) => {
    if (e.target.files.length) setFile(e.target.files[0]);
  });

  // Drag & Drop
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

//  Log de procesamiento

/**
 * Agrega una línea al panel de log en pantalla.
 *
 * El mensaje se renderiza respetando saltos de línea (`\n`) gracias a
 * `white-space: pre-line`. Si `isError` es `true`, se aplica la clase CSS
 * `err` para resaltar visualmente la entrada.
 *
 * @function log
 * @param {string} msg     - Texto a mostrar. Puede contener `\n`.
 * @param {boolean} [isError=false] - Si es `true`, marca la línea como error.
 * @returns {void}
 * @sideEffect
 *   - Crea y anexa un `<div>` dentro del contenedor `#log`.
 *   - Hace scroll automático al final del panel.
 *   - Si `#log` no existe en el DOM, emite una advertencia por `console.warn`.
 */
export function log(msg, isError = false) {
  const el = document.getElementById("log");
  if (!el) {
    console.warn("No se encontró #log en el DOM:", msg);
    return;
  }
  el.style.display = "block";
  const line = document.createElement("div");
  line.textContent = msg;
  line.style.whiteSpace = "pre-line"; // respeta \n
  if (isError) line.className = "err";
  el.appendChild(line);
  el.scrollTop = el.scrollHeight;
}

/**
 * Limpia completamente el contenido del panel de log.
 *
 * @function clearLog
 * @returns {void}
 * @sideEffect Vacía el `innerHTML` del contenedor `#log`.
 */
export function clearLog() {
  const el = document.getElementById("log");
  if (el) el.innerHTML = "";
}

//  Descarga de resultados

/**
 * Crea un enlace de descarga para un blob de datos y lo anexa a la lista
 * de descargas del DOM.
 *
 * Genera un objeto URL temporal (`URL.createObjectURL`) que se revoca
 * automáticamente 10 segundos después del clic para evitar fugas de memoria.
 *
 * @function descargarBlob
 * @param {string} nombre - Nombre sugerido para el archivo descargado (incluye extensión).
 * @param {BlobPart} data - Contenido binario o de texto del archivo.
 * @param {string} tipo   - MIME type del blob (ej. `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`).
 * @returns {HTMLAnchorElement} - El elemento `<a>` creado, por si se necesita manipularlo posteriormente.
 * @sideEffect
 *   - Crea un `<li>` con un `<a download>` dentro de `#download-list`.
 *   - Registra un `setTimeout` para revocar el object URL tras 10 000 ms.
 */
export function descargarBlob(nombre, data, tipo) {
  const blob = new Blob([data], { type: tipo });
  const url = URL.createObjectURL(blob);

  const li = document.createElement("li");
  const a = document.createElement("a");
  a.href = url;
  a.download = nombre;
  a.textContent = nombre;
  a.style.fontWeight = "bold";

  a.addEventListener("click", () => {
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  });

  li.appendChild(a);
  document.getElementById("download-list")?.appendChild(li);
  return a;
}
