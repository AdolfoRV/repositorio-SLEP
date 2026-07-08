// Todo lo que manipula el DOM: carga de archivos, log y descargas.

// ── Estado de archivos ──
export const files = {
  licencias: null,
};

export function updateBtn() {
  const ready = !!files.licencias;
  const btn = document.getElementById("btn-procesar");
  if (!btn) return;
  btn.disabled = !ready;
  btn.textContent = ready ? "Procesar y descargar" : "Selecciona el archivo de licencias";
}

export function setupDrop(boxId, key) {
  const box = document.getElementById(boxId);
  if (!box) return;

  const input = box.querySelector('input[type="file"]');
  const tag = box.querySelector('div[id^="tag"]');

  const setFile = (file) => {
    files[key] = file;
    if (tag) tag.innerHTML = `<span class="file-tag">${file.name}</span>`;
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

// ── Log ──
export function log(msg, isError = false) {
  const el = document.getElementById("log");
  if (!el) {
    console.warn("No se encontró #log en el DOM:", msg);
    return;
  }
  el.style.display = "block";
  const line = document.createElement("div");
  line.textContent = msg;
  line.style.whiteSpace = "pre-line";   // ← respeta \n
  if (isError) line.className = "err";
  el.appendChild(line);
  el.scrollTop = el.scrollHeight;
}

export function clearLog() {
  const el = document.getElementById("log");
  if (el) el.innerHTML = "";
}

// ── Descargas ──
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