// assets/slep/download.js
// Crea enlaces de descarga para los archivos resultantes.

export function descargarBlob(nombre, data, tipo) {
  const blob = new Blob([data], { type: tipo });
  const url = URL.createObjectURL(blob);

  const li = document.createElement("li");
  const a = document.createElement("a");
  a.href = url;
  a.download = nombre;
  a.textContent = nombre;
  a.style.fontWeight = "bold";
  li.appendChild(a);

  const list = document.getElementById("download-list");
  if (list) list.appendChild(li);

  return a;
}
