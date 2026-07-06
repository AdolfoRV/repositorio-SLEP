// assets/slep/upload.js
// Estado de archivos seleccionados y manejo de drag&drop.

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
