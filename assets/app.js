// assets/app.js
// Punto de entrada. Conecta los módulos de lógica con los elementos del DOM.

import { setupDrop } from "./slep/ui.js";
import { procesarArchivos } from "./slep/processor.js";

export function initSlepApp() {
  setupDrop("drop-licencias", "licencias");
  document.getElementById("btn-procesar").addEventListener("click", procesarArchivos);
}
