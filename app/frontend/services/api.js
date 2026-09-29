/**
 * Cliente HTTP para la API del Iberostar Inventory Synchronizer.
 * Todas las funciones devuelven promesas y lanzan un Error con un
 * mensaje legible cuando la respuesta no es correcta.
 */
const Api = (() => {
  const NETWORK_ERROR =
    "No se pudo contactar con la aplicación. Ciérrala y vuelve a abrirla.";

  async function request(url, options) {
    try {
      return await fetch(url, options);
    } catch (error) {
      throw new Error(NETWORK_ERROR);
    }
  }

  async function parseJsonOrThrow(response) {
    let payload = null;

    try {
      payload = await response.json();
    } catch (error) {
      // La respuesta no era JSON (por ejemplo, un error HTML del servidor).
    }

    if (!response.ok) {
      const message =
        (payload && payload.error) ||
        `Error ${response.status}: ${response.statusText}`;
      throw new Error(message);
    }

    return payload;
  }

  async function getJson(url) {
    return parseJsonOrThrow(await request(url));
  }

  async function postJson(url, body) {
    return parseJsonOrThrow(await request(url, { method: "POST", body }));
  }

  const getInfo = () => getJson("/api/info");
  const getSalesPoints = () => getJson("/api/sales-points");
  const getPeriods = () => getJson("/api/periods");
  const openFolder = () => postJson("/api/open-folder");

  function ping() {
    return request("/api/ping").catch(() => null);
  }

  function syncFiles(files) {
    const formData = new FormData();

    for (const file of files) {
      formData.append("files", file);
    }

    return postJson("/api/sync", formData);
  }

  /**
   * Exporta un Excel mensual.
   *
   * Dentro de la ventana de escritorio se usa el diálogo nativo
   * "Guardar como"; en el navegador, una descarga normal.
   *
   * Devuelve un texto describiendo el resultado, o null si el usuario
   * canceló el diálogo.
   */
  async function exportExcel(salesPoint, year, month) {
    const desktop = window.pywebview && window.pywebview.api;

    if (desktop && desktop.save_export) {
      const result = await desktop.save_export(salesPoint, Number(year), Number(month));

      if (result.error) throw new Error(result.error);
      if (result.cancelled) return null;
      return `Guardado en ${result.saved}`;
    }

    const params = new URLSearchParams({ sales_point: salesPoint, year, month });
    const response = await request(`/api/export?${params.toString()}`);

    if (!response.ok) {
      await parseJsonOrThrow(response);
    }

    const disposition = response.headers.get("Content-Disposition") || "";
    const match = disposition.match(/filename="?([^"]+)"?/);
    const filename = match ? match[1] : `${salesPoint}.xlsx`;

    const url = URL.createObjectURL(await response.blob());
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);

    return `Descargado ${filename}`;
  }

  return {
    getInfo,
    getSalesPoints,
    getPeriods,
    openFolder,
    ping,
    syncFiles,
    exportExcel,
  };
})();
