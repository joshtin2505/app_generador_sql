const ui = {
  excelFile: document.getElementById("excelFile"),
  secEjecutable: document.getElementById("secEjecutable"),
  sortBy: document.getElementById("sortBy"),
  btnProcesar: document.getElementById("btnProcesar"),
  btnDescargar: document.getElementById("btnDescargar"),
  btnDescargarTodo: document.getElementById("btnDescargarTodo"),
  btnAbrirTodo: document.getElementById("btnAbrirTodo"),
  status: document.getElementById("status"),
  rows: document.getElementById("rows")
};

let generatedFiles = [];
let generatedPairs = [];
let lastWorkbookName = "";

function setStatus(message) {
  ui.status.textContent = message;
}

function normalizeHeader(value) {
  return String(value || "").trim().toUpperCase();
}

function cleanValue(value) {
  return String(value ?? "").trim();
}

function toSlug(parametro) {
  return cleanValue(parametro).toLowerCase();
}

function toDownloadFileName(path) {
  const normalized = String(path || "").replace(/\\/g, "/");
  const parts = normalized.split("/").filter(Boolean);
  return parts.length ? parts[parts.length - 1] : "archivo.sql";
}

function escapeSqlText(value) {
  return String(value ?? "").replace(/'/g, "''");
}

function maybeNullSql(value) {
  const v = cleanValue(value);
  return v === "" ? "NULL" : `'${escapeSqlText(v)}'`;
}

function toUpperWithoutParentheses(text) {
  return cleanValue(text)
    .replace(/[_-]+/g, " ")
    .replace(/\s*\([^)]*\)\s*/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function stripDiacritics(text) {
  return String(text || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function detectArticulo(nombre) {
  const normalized = stripDiacritics(toUpperWithoutParentheses(nombre));
  const tokens = normalized.split(" ").filter(Boolean);
  const stopWords = new Set(["DE", "DEL", "LA", "EL", "LOS", "LAS", "Y", "E", "O", "U"]);
  const noun = tokens.find((t) => !stopWords.has(t)) || "";

  const feminineWords = new Set([
    "FECHA",
    "CUENTA",
    "LISTA",
    "LINEA",
    "CLASE",
    "SERIE",
    "SUCURSAL",
    "SECCION",
    "VERSION",
    "OPCION"
  ]);

  const masculineWords = new Set([
    "DIA",
    "DOCUMENTO",
    "FONDO",
    "CENTRO",
    "CODIGO",
    "NUMERO",
    "ESTADO",
    "TIPO"
  ]);

  if (feminineWords.has(noun)) return "LA";
  if (masculineWords.has(noun)) return "EL";

  const masculineEndingExceptions = new Set(["PROBLEMA", "SISTEMA", "TEMA", "MAPA", "PROGRAMA"]);
  if (noun.endsWith("A") && !masculineEndingExceptions.has(noun)) return "LA";
  return "EL";
}

function buildParametroDescripcion(nombre) {
  const articulo = detectArticulo(nombre);
  return `PARÁMETRO PARA CONSULTAR POR ${articulo} ${toUpperWithoutParentheses(nombre)}.`;
}

function buildEjecutableDescripcion(nombre) {
  const articulo = detectArticulo(nombre);
  return `ASOCIA PARÁMETRO PARA CONSULTAR POR ${articulo} ${toUpperWithoutParentheses(nombre)}.`;
}

function toSortableNumber(value) {
  const n = Number(cleanValue(value));
  return Number.isFinite(n) ? n : Number.MAX_SAFE_INTEGER;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function detectHeaderRow(rawRows, requiredHeaders) {
  const maxRowsToInspect = Math.min(rawRows.length, 20);
  for (let i = 0; i < maxRowsToInspect; i++) {
    const row = rawRows[i].map(normalizeHeader);
    const matchesAll = requiredHeaders.every((h) => row.includes(h));
    if (matchesAll) {
      return i;
    }
  }
  return -1;
}

function sheetToObjects(sheet, requiredHeaders) {
  const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false, defval: "" });
  const headerRowIndex = detectHeaderRow(rawRows, requiredHeaders);

  if (headerRowIndex === -1) {
    throw new Error(`No se encontraron encabezados requeridos: ${requiredHeaders.join(", ")}`);
  }

  const headers = rawRows[headerRowIndex].map((h) => String(h || "").trim());
  const dataRows = rawRows.slice(headerRowIndex + 1);

  return dataRows
    .map((row) => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index] ?? "";
      });
      return obj;
    })
    .filter((row) => Object.values(row).some((v) => cleanValue(v) !== ""));
}

function findSheet(workbook, expectedName) {
  const wanted = expectedName.toUpperCase();
  const exact = workbook.SheetNames.find((name) => normalizeHeader(name) === wanted);
  if (exact) return workbook.Sheets[exact];

  const loose = workbook.SheetNames.find((name) => normalizeHeader(name).includes(wanted));
  if (loose) return workbook.Sheets[loose];

  return null;
}

function buildParametroSql(paramRow) {
  const secParametro = cleanValue(paramRow.SEC_PARAMETRO);
  const parametro = cleanValue(paramRow.PARAMETRO);
  const nombre = cleanValue(paramRow.NOMBRE);
  const estado = cleanValue(paramRow.ESTADO || "A");
  const tipoDato = cleanValue(paramRow.TIPO_DATO);
  const datos = cleanValue(paramRow.DATOS);
  const descripcion = buildParametroDescripcion(nombre);

  const hasQuery = datos !== "";
  const queryBlock = hasQuery
    ? `    mi_query        VARCHAR2(2000);\n`
    : "";

  const queryAssign = hasQuery
    ? `\n    mi_query := '${escapeSqlText(datos)}';\n`
    : "";

  const queryValue = hasQuery ? "mi_query" : "NULL";

  return `PROMPT Insertando o actualizando en la tabla SRT_PARAMETRO, el objeto ${parametro}...
DECLARE
\tmi_existe_registro\tNUMBER := 0;
${queryBlock}BEGIN

\tSELECT COUNT(*)
    INTO mi_existe_registro
\tFROM srt_parametro
\tWHERE sec_parametro = ${secParametro};
${queryAssign}
\tIF mi_existe_registro = 0 THEN
        INSERT INTO srt_parametro  (
            sec_parametro,
            parametro,
            nombre,
            estado,
            tipo_dato,
            datos,
            descripcion
        ) 
        VALUES ( 
            ${secParametro}, 
            '${escapeSqlText(parametro)}',
            '${escapeSqlText(nombre)}',
            '${escapeSqlText(estado)}',
            '${escapeSqlText(tipoDato)}',
            ${queryValue},
            '${escapeSqlText(descripcion)}'
        );
\tELSE
\t\tUPDATE srt_parametro
\t\tSET 
            parametro       = '${escapeSqlText(parametro)}',
            nombre          = '${escapeSqlText(nombre)}',
            estado          = '${escapeSqlText(estado)}',
            tipo_dato       = '${escapeSqlText(tipoDato)}',
            datos           = ${queryValue},
            descripcion     = '${escapeSqlText(descripcion)}'
\t\tWHERE 
            sec_parametro = ${secParametro};\t
\tEND IF;
\t
\tCOMMIT;
    
EXCEPTION
\tWHEN OTHERS THEN
\t\tROLLBACK;
\t\tpk_excepcion.error_aplicacion;
END;
/`;
}

function buildEjecutableSql(epRow, paramRow) {
  const secEP = cleanValue(epRow.SEC_EJECUTABLE_PARAMETRO);
  const secEjecutable = cleanValue(epRow.SEC_EJECUTABLE);
  const secParametro = cleanValue(epRow.SEC_PARAMETRO);
  const requerido = cleanValue(epRow.REQUERIDO || "N");
  const valorAsumido = cleanValue(epRow.VALOR_ASUMIDO);
  const seleccionMultiple = cleanValue(epRow.SELECCION_MULTIPLE || "N");
  const ordenamiento = cleanValue(epRow.ORDENAMIENTO || "0");
  const agrupador = cleanValue(epRow.AGRUPADOR);
  const parametro = cleanValue(paramRow.PARAMETRO);
  const nombre = cleanValue(paramRow.NOMBRE);
  const descripcion = buildEjecutableDescripcion(nombre);

  return `PROMPT Insertando o actualizando en la tabla SRT_EJECUTABLE_PARAMETRO, el objeto ${parametro}...
DECLARE
\tmi_existe_registro\tNUMBER := 0;
BEGIN

\tSELECT COUNT(*)
    INTO mi_existe_registro
\tFROM srt_ejecutable_parametro
\tWHERE\t
        sec_ejecutable_parametro = ${secEP};
\t\t
\tIF mi_existe_registro = 0 THEN
        INSERT INTO srt_ejecutable_parametro (
            sec_ejecutable_parametro,
            sec_ejecutable,
            sec_parametro,
            requerido,
            valor_asumido,
            seleccion_multiple,
            ordenamiento,
            agrupador,
            descripcion
        ) 
        VALUES (
            ${secEP},
            ${secEjecutable}, 
            ${secParametro},
            '${escapeSqlText(requerido)}',
            ${maybeNullSql(valorAsumido)},
            '${escapeSqlText(seleccionMultiple)}',
            ${ordenamiento},
            ${maybeNullSql(agrupador)},
            '${escapeSqlText(descripcion)}'
        );
\tELSE
\t\tUPDATE srt_ejecutable_parametro
\t\tSET 
            sec_ejecutable      = ${secEjecutable},
            sec_parametro       = ${secParametro},
            requerido           = '${escapeSqlText(requerido)}',
            valor_asumido       = ${maybeNullSql(valorAsumido)},
            seleccion_multiple  = '${escapeSqlText(seleccionMultiple)}',
            ordenamiento        = ${ordenamiento},
            agrupador           = ${maybeNullSql(agrupador)},
            descripcion         = '${escapeSqlText(descripcion)}'
\t\tWHERE 
            sec_ejecutable_parametro = ${secEP};\t
\tEND IF;
\t
\tCOMMIT;
    
EXCEPTION
\tWHEN OTHERS THEN
\t\tROLLBACK;
\t\tpk_excepcion.error_aplicacion;
END;
/`;
}

function drawRows(items) {
  ui.rows.innerHTML = "";
  if (!items.length) return;

  const html = items
    .map((item, index) => {
      const ep = item.ejecutable;
      const p = item.parametro;
      return `<tr>
        <td>${index + 1}</td>
        <td>${cleanValue(ep.SEC_EJECUTABLE_PARAMETRO)}</td>
        <td>${cleanValue(p.SEC_PARAMETRO)}</td>
        <td>${cleanValue(p.PARAMETRO)}</td>
        <td>${cleanValue(p.NOMBRE)}</td>
        <td>${cleanValue(p.TIPO_DATO)}</td>
        <td>${cleanValue(ep.ORDENAMIENTO)}</td>
        <td>${cleanValue(ep.REQUERIDO)}</td>
        <td>
          <button type="button" class="btn btn-mini btn-primary row-download" data-index="${index}">Descargar uno por uno</button>
          <button type="button" class="btn btn-mini row-open" data-index="${index}">Abrir</button>
        </td>
      </tr>`;
    })
    .join("");

  ui.rows.innerHTML = html;

  ui.rows.querySelectorAll(".row-download").forEach((button) => {
    button.addEventListener("click", () => {
      const idx = Number(button.dataset.index);
      downloadPairOneByOne(generatedPairs[idx]);
    });
  });

  ui.rows.querySelectorAll(".row-open").forEach((button) => {
    button.addEventListener("click", () => {
      const idx = Number(button.dataset.index);
      openPairInBrowser(generatedPairs[idx]);
    });
  });
}

function sortPairs(pairs, mode) {
  const cloned = [...pairs];
  if (mode === "sec_ejecutable_parametro") {
    cloned.sort((a, b) => toSortableNumber(a.ejecutable.SEC_EJECUTABLE_PARAMETRO) - toSortableNumber(b.ejecutable.SEC_EJECUTABLE_PARAMETRO));
    return cloned;
  }

  if (mode === "sec_parametro") {
    cloned.sort((a, b) => toSortableNumber(a.parametro.SEC_PARAMETRO) - toSortableNumber(b.parametro.SEC_PARAMETRO));
    return cloned;
  }

  cloned.sort((a, b) => toSortableNumber(a.ejecutable.ORDENAMIENTO) - toSortableNumber(b.ejecutable.ORDENAMIENTO));
  return cloned;
}

function buildGeneratedPairs(filteredPairs) {
  return filteredPairs.map((pair) => {
    const parametro = pair.parametro;
    const ep = pair.ejecutable;
    const folder = toSlug(parametro.PARAMETRO);
    const base = toSlug(parametro.PARAMETRO);

    return {
      ...pair,
      files: [
        {
          path: `${folder}/${base}.sql`,
          content: buildParametroSql(parametro)
        },
        {
          path: `${folder}/${base}_ejecutable.sql`,
          content: buildEjecutableSql(ep, parametro)
        }
      ]
    };
  });
}

function flattenFiles(pairs) {
  return pairs.flatMap((pair) => pair.files);
}

function updateButtonsState(enabled) {
  ui.btnDescargar.disabled = !enabled;
  ui.btnDescargarTodo.disabled = !enabled;
  ui.btnAbrirTodo.disabled = !enabled;
}

function refreshStatus(secEjecutableTarget) {
  const fileList = generatedFiles.map((f) => `- ${f.path}`).join("\n");
  setStatus(
    [
      `Archivo: ${lastWorkbookName}`,
      `SEC_EJECUTABLE: ${secEjecutableTarget}`,
      `Parametros encontrados: ${generatedPairs.length}`,
      `Archivos generados: ${generatedFiles.length}`,
      "",
      fileList
    ].join("\n")
  );
}

function downloadTextFile(path, content) {
  const blob = new Blob([content], { type: "text/plain;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = toDownloadFileName(path);
  a.click();
  URL.revokeObjectURL(a.href);
}

function downloadPairOneByOne(pair) {
  if (!pair) return;
  pair.files.forEach((file) => downloadTextFile(file.path, file.content));
}

function downloadAllUncompressed() {
  if (!generatedFiles.length) return;
  generatedFiles.forEach((file) => downloadTextFile(file.path, file.content));
}

function openPairInBrowser(pair) {
  if (!pair) return;

  const title = `SQL - ${cleanValue(pair.parametro.PARAMETRO)}`;
  const html = `<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>${escapeHtml(title)}</title>
  <style>
    body { font-family: Consolas, monospace; margin: 16px; background: #f7f7f7; }
    h1, h2 { font-family: Arial, sans-serif; }
    pre { background: #fff; padding: 12px; border: 1px solid #ddd; border-radius: 8px; overflow: auto; }
  </style>
</head>
<body>
  <h1>${escapeHtml(title)}</h1>
  <h2>${escapeHtml(pair.files[0].path)}</h2>
  <pre>${escapeHtml(pair.files[0].content)}</pre>
  <h2>${escapeHtml(pair.files[1].path)}</h2>
  <pre>${escapeHtml(pair.files[1].content)}</pre>
</body>
</html>`;

  const win = window.open("", "_blank");
  if (!win) {
    setStatus("El navegador bloqueó la ventana emergente. Habilita popups para esta página.");
    return;
  }
  win.document.open();
  win.document.write(html);
  win.document.close();
}

function openAllInBrowser() {
  if (!generatedFiles.length) return;

  const blocks = generatedFiles
    .map((file) => `<h2>${escapeHtml(file.path)}</h2><pre>${escapeHtml(file.content)}</pre>`)
    .join("\n");

  const html = `<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>SQL Generados</title>
  <style>
    body { font-family: Consolas, monospace; margin: 16px; background: #f7f7f7; }
    h1, h2 { font-family: Arial, sans-serif; }
    pre { background: #fff; padding: 12px; border: 1px solid #ddd; border-radius: 8px; overflow: auto; }
  </style>
</head>
<body>
  <h1>SQL Generados</h1>
  ${blocks}
</body>
</html>`;

  const win = window.open("", "_blank");
  if (!win) {
    setStatus("El navegador bloqueó la ventana emergente. Habilita popups para esta página.");
    return;
  }
  win.document.open();
  win.document.write(html);
  win.document.close();
}

async function processWorkbook() {
  const file = ui.excelFile.files[0];
  const secEjecutableTarget = cleanValue(ui.secEjecutable.value);

  if (!file) {
    setStatus("Debes seleccionar un archivo Excel.");
    return;
  }

  if (secEjecutableTarget === "") {
    setStatus("Debes indicar SEC_EJECUTABLE.");
    return;
  }

  generatedFiles = [];
  generatedPairs = [];
  updateButtonsState(false);
  ui.rows.innerHTML = "";

  try {
    setStatus("Leyendo archivo Excel...");
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    lastWorkbookName = file.name;

    const shEjecutableParametro = findSheet(workbook, "ejecutable_parametro");
    const shParametro = findSheet(workbook, "parametro");

    if (!shEjecutableParametro || !shParametro) {
      throw new Error("No se encontraron las hojas ejecutable_parametro y/o parametro.");
    }

    const ejecutableRows = sheetToObjects(shEjecutableParametro, ["SEC_EJECUTABLE", "SEC_PARAMETRO", "SEC_EJECUTABLE_PARAMETRO"]);
    const parametroRows = sheetToObjects(shParametro, ["SEC_PARAMETRO", "PARAMETRO", "NOMBRE", "TIPO_DATO", "DATOS"]);

    const paramsBySec = new Map(parametroRows.map((row) => [cleanValue(row.SEC_PARAMETRO), row]));

    const filtered = ejecutableRows
      .filter((row) => cleanValue(row.SEC_EJECUTABLE) === secEjecutableTarget)
      .map((row) => ({
        ejecutable: row,
        parametro: paramsBySec.get(cleanValue(row.SEC_PARAMETRO))
      }))
      .filter((pair) => pair.parametro);

    if (!filtered.length) {
      setStatus(`No se encontraron registros para SEC_EJECUTABLE = ${secEjecutableTarget}.`);
      return;
    }

    generatedPairs = buildGeneratedPairs(filtered);
    generatedPairs = sortPairs(generatedPairs, cleanValue(ui.sortBy.value).toLowerCase());
    generatedFiles = flattenFiles(generatedPairs);

    drawRows(generatedPairs);
    updateButtonsState(true);
    refreshStatus(secEjecutableTarget);
  } catch (error) {
    setStatus(`Error: ${error.message}`);
  }
}

async function downloadZip() {
  if (!generatedFiles.length) return;

  const zip = new JSZip();
  generatedFiles.forEach((file) => {
    zip.file(file.path, file.content);
  });

  const secEjecutableTarget = cleanValue(ui.secEjecutable.value);
  const blob = await zip.generateAsync({ type: "blob" });

  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `sql_parametros_${secEjecutableTarget}.zip`;
  a.click();
  URL.revokeObjectURL(a.href);
}

ui.btnProcesar.addEventListener("click", processWorkbook);
ui.btnDescargar.addEventListener("click", downloadZip);
ui.btnDescargarTodo.addEventListener("click", downloadAllUncompressed);
ui.btnAbrirTodo.addEventListener("click", openAllInBrowser);

ui.sortBy.addEventListener("change", () => {
  if (!generatedPairs.length) return;
  generatedPairs = sortPairs(generatedPairs, cleanValue(ui.sortBy.value).toLowerCase());
  generatedFiles = flattenFiles(generatedPairs);
  drawRows(generatedPairs);
  refreshStatus(cleanValue(ui.secEjecutable.value));
});
