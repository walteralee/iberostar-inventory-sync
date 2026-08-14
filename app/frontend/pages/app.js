(() => {
  const MONTHS = [
    "ENERO",
    "FEBRERO",
    "MARZO",
    "ABRIL",
    "MAYO",
    "JUNIO",
    "JULIO",
    "AGOSTO",
    "SEPTIEMBRE",
    "OCTUBRE",
    "NOVIEMBRE",
    "DICIEMBRE",
  ];

  const MONTH_LABELS = [
    "Enero",
    "Febrero",
    "Marzo",
    "Abril",
    "Mayo",
    "Junio",
    "Julio",
    "Agosto",
    "Septiembre",
    "Octubre",
    "Noviembre",
    "Diciembre",
  ];

  let selectedFiles = [];

  // Rellenada por initExportPanel(); la llama initSyncPanel() al terminar
  // una sincronización para que la pestaña de exportación se actualice
  // sin necesidad de recargar la página.
  let refreshExportOptions = async () => {};

  // ================= TABS =================

  function initTabs() {
    const buttons = document.querySelectorAll(".tabs__button");
    const panels = {
      sync: document.getElementById("panel-sync"),
      export: document.getElementById("panel-export"),
    };

    buttons.forEach((button) => {
      button.addEventListener("click", () => {
        buttons.forEach((other) => {
          other.classList.remove("is-active");
          other.setAttribute("aria-selected", "false");
        });
        Object.values(panels).forEach((panel) => panel.classList.remove("is-active"));

        button.classList.add("is-active");
        button.setAttribute("aria-selected", "true");
        panels[button.dataset.tab].classList.add("is-active");

        if (button.dataset.tab === "export") {
          refreshExportOptions(true);
        }
      });
    });
  }

  // ================= SYNC PANEL =================

  function initSyncPanel() {
    const dropzone = document.getElementById("dropzone");
    const fileInput = document.getElementById("file-input");
    const fileList = document.getElementById("file-list");
    const syncButton = document.getElementById("sync-button");
    const clearButton = document.getElementById("clear-files-button");
    const statusBox = document.getElementById("sync-status");
    const resultsBox = document.getElementById("sync-results");

    function renderFileList() {
      fileList.innerHTML = "";

      selectedFiles.forEach((file, index) => {
        const item = document.createElement("li");
        item.className = "file-list__item";

        const name = document.createElement("span");
        name.className = "file-list__name";
        name.textContent = file.name;

        const removeButton = document.createElement("button");
        removeButton.className = "file-list__remove";
        removeButton.type = "button";
        removeButton.textContent = "✕";
        removeButton.addEventListener("click", () => {
          selectedFiles.splice(index, 1);
          renderFileList();
        });

        item.appendChild(name);
        item.appendChild(removeButton);
        fileList.appendChild(item);
      });

      const hasFiles = selectedFiles.length > 0;
      syncButton.disabled = !hasFiles;
      clearButton.disabled = !hasFiles;
    }

    function addFiles(fileListLike) {
      const incoming = Array.from(fileListLike).filter((file) =>
        file.name.toLowerCase().endsWith(".xlsx")
      );

      const existingKeys = new Set(
        selectedFiles.map((file) => `${file.name}__${file.size}`)
      );

      incoming.forEach((file) => {
        const key = `${file.name}__${file.size}`;
        if (!existingKeys.has(key)) {
          selectedFiles.push(file);
          existingKeys.add(key);
        }
      });

      renderFileList();
    }

    fileInput.addEventListener("change", (event) => {
      addFiles(event.target.files);
      fileInput.value = "";
    });

    ["dragover", "dragenter"].forEach((eventName) => {
      dropzone.addEventListener(eventName, (event) => {
        event.preventDefault();
        dropzone.classList.add("is-dragover");
      });
    });

    ["dragleave", "dragend", "drop"].forEach((eventName) => {
      dropzone.addEventListener(eventName, (event) => {
        event.preventDefault();
        dropzone.classList.remove("is-dragover");
      });
    });

    dropzone.addEventListener("drop", (event) => {
      if (event.dataTransfer && event.dataTransfer.files) {
        addFiles(event.dataTransfer.files);
      }
    });

    clearButton.addEventListener("click", () => {
      selectedFiles = [];
      renderFileList();
    });

    syncButton.addEventListener("click", async () => {
      if (selectedFiles.length === 0) return;

      syncButton.disabled = true;
      clearButton.disabled = true;
      resultsBox.innerHTML = "";
      showStatus(statusBox, "info", "Sincronizando, esto puede tardar unos segundos…");

      try {
        const result = await Api.syncFiles(selectedFiles);
        hideStatus(statusBox);
        renderSyncResults(resultsBox, result);
        selectedFiles = [];
        renderFileList();

        // La sincronización puede haber creado o modificado Excel
        // mensuales: refrescamos la pestaña de exportación en caliente.
        refreshExportOptions(true);
      } catch (error) {
        showStatus(statusBox, "error", error.message);
      } finally {
        syncButton.disabled = selectedFiles.length === 0;
        clearButton.disabled = selectedFiles.length === 0;
      }
    });

    renderFileList();
  }

  function showStatus(box, type, message) {
    box.hidden = false;
    box.className = `status status--${type}`;
    box.textContent = message;
  }

  function hideStatus(box) {
    box.hidden = true;
  }

  // ================= SYNC RESULTS =================
  //
  // Diseño deliberadamente minimalista: una frase con el resultado, las
  // dos o tres cifras que de verdad importan, y el detalle completo (si
  // lo hay) escondido detrás de un desplegable para no abrumar cuando
  // todo ha ido bien.

  function renderSyncResults(container, result) {
    const {
      import_summary: importSummary,
      pending_warnings: pendingWarnings,
      synchronization_totals: syncTotals,
    } = result;

    const errorMessages = [
      ...importSummary.file_error_messages,
      ...importSummary.row_error_messages,
      ...importSummary.conflict_messages,
      ...syncTotals.error_messages,
    ];

    const warningMessages = [
      ...ignoredRowBullets(importSummary),
      ...importSummary.discrepancy_messages,
      ...importSummary.duplicate_messages,
      ...importSummary.thousands_format_messages,
      ...pendingWarnings,
    ];

    const deliveriesApplied =
      syncTotals.synchronized_deliveries + syncTotals.recovered_deliveries;

    const summary = document.createElement("div");
    summary.className = "summary";

    const tone = errorMessages.length > 0
      ? "is-error"
      : warningMessages.length > 0
      ? "is-warning"
      : "is-ok";

    const bannerText = {
      "is-ok": "✅ Sincronización completada sin incidencias.",
      "is-warning": "⚠️ Sincronización completada con avisos.",
      "is-error": "❌ Sincronización completada con errores.",
    }[tone];

    const banner = document.createElement("div");
    banner.className = `summary__banner ${tone}`;
    banner.textContent = bannerText;
    summary.appendChild(banner);

    const meta = document.createElement("p");
    meta.className = "summary__meta";
    meta.textContent = buildMetaLine(deliveriesApplied, syncTotals);
    summary.appendChild(meta);

    const incidentCount = errorMessages.length + warningMessages.length;

    if (incidentCount > 0) {
      const details = document.createElement("details");
      details.className = "summary__details";

      const toggle = document.createElement("summary");
      toggle.textContent = `Ver detalle (${incidentCount})`;
      details.appendChild(toggle);

      if (errorMessages.length > 0) {
        details.appendChild(buildIncidentGroup("Errores", errorMessages, "error"));
      }

      if (warningMessages.length > 0) {
        details.appendChild(buildIncidentGroup("Avisos", warningMessages, "warning"));
      }

      summary.appendChild(details);
    }

    container.appendChild(summary);
  }

  function buildMetaLine(deliveriesApplied, syncTotals) {
    const parts = [
      `${deliveriesApplied} ${pluralize(deliveriesApplied, "entrega sincronizada", "entregas sincronizadas")}`,
      `${syncTotals.products_written} ${pluralize(syncTotals.products_written, "producto escrito", "productos escritos")}`,
    ];

    if (syncTotals.skipped_deliveries > 0) {
      parts.push(
        `${syncTotals.skipped_deliveries} ya ${pluralize(
          syncTotals.skipped_deliveries,
          "estaba sincronizada",
          "estaban sincronizadas"
        )}`
      );
    }

    return parts.join(" · ");
  }

  function ignoredRowBullets(importSummary) {
    const bullets = [];

    if (importSummary.ignored_group_count > 0) {
      bullets.push(
        `${importSummary.ignored_group_count} ${pluralize(
          importSummary.ignored_group_count,
          "fila ignorada",
          "filas ignoradas"
        )} por grupo de producto no admitido.`
      );
    }

    if (importSummary.ignored_sales_point_count > 0) {
      bullets.push(
        `${importSummary.ignored_sales_point_count} ${pluralize(
          importSummary.ignored_sales_point_count,
          "fila ignorada",
          "filas ignoradas"
        )} por punto de venta no reconocido.`
      );
    }

    return bullets;
  }

  function pluralize(count, singular, plural) {
    return count === 1 ? singular : plural;
  }

  function buildIncidentGroup(title, messages, tone) {
    const group = document.createElement("div");
    group.className = `summary__group summary__group--${tone}`;

    const heading = document.createElement("h4");
    heading.textContent = `${title} (${messages.length})`;
    group.appendChild(heading);

    const list = document.createElement("ul");
    list.className = "message-list";

    messages.forEach((message) => {
      const item = document.createElement("li");
      item.textContent = message;
      list.appendChild(item);
    });

    group.appendChild(list);
    return group;
  }

  // ================= EXPORT PANEL =================

  function initExportPanel() {
    const salesPointSelect = document.getElementById("export-sales-point");
    const yearSelect = document.getElementById("export-year");
    const monthSelect = document.getElementById("export-month");
    const exportButton = document.getElementById("export-button");
    const statusBox = document.getElementById("export-status");

    let periodsByYear = {};

    function populateMonths(year, preferredMonth) {
      monthSelect.innerHTML = "";
      const monthsPresent = periodsByYear[year] || [];

      monthsPresent.forEach((monthName) => {
        const index = MONTHS.indexOf(monthName.toUpperCase());
        if (index === -1) return;

        const option = document.createElement("option");
        option.value = String(index + 1);
        option.textContent = MONTH_LABELS[index];
        monthSelect.appendChild(option);
      });

      if (monthsPresent.length === 0) {
        const option = document.createElement("option");
        option.value = "";
        option.textContent = "Sin datos";
        monthSelect.appendChild(option);
        return;
      }

      if (preferredMonth && monthSelect.querySelector(`option[value="${preferredMonth}"]`)) {
        monthSelect.value = preferredMonth;
      }
    }

    async function load(preserveSelection) {
      const previousSalesPoint = preserveSelection ? salesPointSelect.value : null;
      const previousYear = preserveSelection ? yearSelect.value : null;
      const previousMonth = preserveSelection ? monthSelect.value : null;

      try {
        const [salesPoints, periods] = await Promise.all([
          Api.getSalesPoints(),
          Api.getPeriods(),
        ]);

        periodsByYear = periods;

        salesPointSelect.innerHTML = "";
        salesPoints.forEach((point) => {
          const option = document.createElement("option");
          option.value = point;
          option.textContent = point.replace(/_/g, " ");
          salesPointSelect.appendChild(option);
        });

        if (previousSalesPoint && salesPoints.includes(previousSalesPoint)) {
          salesPointSelect.value = previousSalesPoint;
        }

        const years = Object.keys(periodsByYear).sort();
        yearSelect.innerHTML = "";

        if (years.length === 0) {
          const option = document.createElement("option");
          option.value = "";
          option.textContent = "Sin Excel sincronizados todavía";
          yearSelect.appendChild(option);
          populateMonths("");
          return;
        }

        years.forEach((year) => {
          const option = document.createElement("option");
          option.value = year;
          option.textContent = year;
          yearSelect.appendChild(option);
        });

        const targetYear =
          previousYear && years.includes(previousYear)
            ? previousYear
            : years[years.length - 1];

        yearSelect.value = targetYear;
        populateMonths(targetYear, previousMonth);
      } catch (error) {
        showStatus(statusBox, "error", `No se pudo cargar la información: ${error.message}`);
      }
    }

    yearSelect.addEventListener("change", () => {
      populateMonths(yearSelect.value);
    });

    exportButton.addEventListener("click", async () => {
      const salesPoint = salesPointSelect.value;
      const year = yearSelect.value;
      const month = monthSelect.value;

      if (!salesPoint || !year || !month) {
        showStatus(statusBox, "error", "Selecciona punto de venta, año y mes.");
        return;
      }

      exportButton.disabled = true;
      showStatus(statusBox, "info", "Preparando la descarga…");

      try {
        await Api.downloadExport(salesPoint, year, month);
        showStatus(statusBox, "success", "Descarga completada.");
      } catch (error) {
        showStatus(statusBox, "error", error.message);
      } finally {
        exportButton.disabled = false;
      }
    });

    refreshExportOptions = load;
    load(false);
  }

  document.addEventListener("DOMContentLoaded", () => {
    initTabs();
    initSyncPanel();
    initExportPanel();
  });
})();
