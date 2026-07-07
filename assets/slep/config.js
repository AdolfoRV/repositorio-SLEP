// assets/slep/config.js
// Rutas y constantes compartidas por el migrador SLEP.

export const SLEP_MODULES = [
  "__init__.py",
  "constants.py",
  "utils.py",
  "core.py",
];

const SITE_ROOT = window.location.href.substring(
  0,
  window.location.href.lastIndexOf("/") + 1
);

export const SLEP_PY_BASE = new URL("scripts/slep/", SITE_ROOT).href;
export const ESTABLECIMIENTOS_URL = new URL(
  "assets/tables/establecimientos.xlsx",
  SITE_ROOT
).href;

export const XLSX_MIME =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
export const ZIP_MIME = "application/zip";

export const PYODIDE_SCRIPT_URL =
  "https://cdn.jsdelivr.net/pyodide/v0.26.2/full/pyodide.js";
