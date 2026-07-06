// assets/app.js
// Punto de entrada. Toda la lógica vive en assets/slep/*.js; este archivo
// solo conecta esos módulos con los elementos del DOM definidos en el .qmd.

import { setupDrop } from "./slep/upload.js";
import { procesarArchivos } from "./slep/processor.js";

export function initSlepApp() {
  setupDrop("drop-licencias", "licencias");
  document.getElementById("btn-procesar").addEventListener("click", procesarArchivos);
}
