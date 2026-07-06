// assets/slep/log.js
// Escribe mensajes en el panel #log de la interfaz.

export function log(msg, isError = false) {
  const el = document.getElementById("log");
  if (!el) {
    console.warn("No se encontró #log en el DOM:", msg);
    return;
  }
  el.style.display = "block";
  const line = document.createElement("div");
  line.textContent = msg;
  if (isError) line.className = "err";
  el.appendChild(line);
  el.scrollTop = el.scrollHeight;
}

export function clearLog() {
  const el = document.getElementById("log");
  if (el) el.innerHTML = "";
}
