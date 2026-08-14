/**
 * Cliente HTTP para la API del Iberostar Inventory Synchronizer.
 * Todas las funciones devuelven promesas y lanzan un Error con un
 * mensaje legible cuando la respuesta no es correcta.
 */
const Api = (() => {
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

  async function getSalesPoints() {
    const response = await fetch("/api/sales-points");
    return parseJsonOrThrow(response);
  }

  async function getPeriods() {
    const response = await fetch("/api/periods");
    return parseJsonOrThrow(response);
  }

  async function syncFiles(files) {
    const formData = new FormData();

    for (const file of files) {
      formData.append("files", file);
    }

    const response = await fetch("/api/sync", {
      method: "POST",
      body: formData,
    });

    return parseJsonOrThrow(response);
  }

  async function downloadExport(salesPoint, year, month) {
    const params = new URLSearchParams({
      sales_point: salesPoint,
      year,
      month,
    });

    const response = await fetch(`/api/export?${params.toString()}`);

    if (!response.ok) {
      const payload = await response.json().catch(() => null);
      throw new Error(
        (payload && payload.error) ||
          `Error ${response.status}: ${response.statusText}`
      );
    }

    const disposition = response.headers.get("Content-Disposition") || "";
    const match = disposition.match(/filename="?([^"]+)"?/);
    const filename = match ? match[1] : `${salesPoint}.xlsx`;

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);

    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  return {
    getSalesPoints,
    getPeriods,
    syncFiles,
    downloadExport,
  };
})();
