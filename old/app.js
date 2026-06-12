const ui = {
  excelFile: document.getElementById("excelFile"),
  secEjecutable: document.getElementById("secEjecutable"),
  rolPrivilegios: document.getElementById("rolPrivilegios"),
  includeMenuPredecessors: document.getElementById("includeMenuPredecessors"),
  sortBy: document.getElementById("sortBy"),
  btnProcesar: document.getElementById("btnProcesar"),
  btnDescargar: document.getElementById("btnDescargar"),
  btnDescargarTodo: document.getElementById("btnDescargarTodo"),
  btnAbrirTodo: document.getElementById("btnAbrirTodo"),
  status: document.getElementById("status"),
  rows: document.getElementById("rows"),
  otherRows: document.getElementById("otherRows")
};

let generatedFiles = [];
let generatedPairs = [];
let generatedContext = null;
let lastWorkbookName = "";

function setStatus(message) {
  ui.status.textContent = message;
}

function normalizeHeader(value) {
  return stripDiacritics(String(value || "").trim()).toUpperCase();
}

function cleanValue(value) {
  return String(value ?? "").trim();
}

function toSlug(parametro) {
  return stripDiacritics(cleanValue(parametro))
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
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

  const headers = rawRows[headerRowIndex].map((h) => normalizeHeader(h));
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

function getSheetRows(sheet) {
  return XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false, defval: "" });
}

function findSheet(workbook, expectedName) {
  const wanted = expectedName.toUpperCase();
  const exact = workbook.SheetNames.find((name) => normalizeHeader(name) === wanted);
  if (exact) return workbook.Sheets[exact];

  const loose = workbook.SheetNames.find((name) => normalizeHeader(name).includes(wanted));
  if (loose) return workbook.Sheets[loose];

  return null;
}

function sqlText(value) {
  const text = cleanValue(value);
  return text === "" ? "NULL" : `'${escapeSqlText(text)}'`;
}

function sqlNumber(value) {
  const text = cleanValue(value);
  return text === "" ? "NULL" : text;
}

function sqlNumberOrDefault(value, defaultValue) {
  const text = cleanValue(value);
  if (text === "") return String(defaultValue);
  return text;
}

function buildComponentLabel(componentRow) {
  const text = cleanValue(componentRow?.nombre);
  if (!text) return "ICEBERG";
  return text.split("-")[0].trim() || text;
}

function parseComponentSheet(sheet) {
  return getSheetRows(sheet)
    .slice(2)
    .map((row) => ({
      id: cleanValue(row[0]),
      componente: cleanValue(row[1]),
      nombre: cleanValue(row[2])
    }))
    .filter((row) => row.id !== "" || row.componente !== "" || row.nombre !== "");
}

function parseMenuSheet(sheet) {
  return getSheetRows(sheet)
    .slice(2)
    .map((row) => ({
      source: {
        menu: cleanValue(row[0]),
        description: cleanValue(row[1]),
        order: cleanValue(row[2]),
        type: cleanValue(row[3]),
        object: cleanValue(row[4]),
        predecessor: cleanValue(row[5]),
        icon: cleanValue(row[6])
      },
      target: {
        menu: cleanValue(row[8]),
        description: cleanValue(row[9]),
        order: cleanValue(row[10]),
        type: cleanValue(row[11]),
        object: cleanValue(row[12]),
        predecessor: cleanValue(row[13])
      }
    }))
    .filter((row) => Object.values(row.source).some((value) => value !== "") || Object.values(row.target).some((value) => value !== ""));
}

function findBestMenuRows(menuRows, executableRow, objectRow) {
  const executableNames = [
    executableRow?.NOMBRE,
    executableRow?.EJECUTABLE,
    executableRow?.DESCRIPCION
  ]
    .map(normalizeHeader)
    .filter(Boolean);

  const objectNames = [
    objectRow?.NOMBRE_OBJETO,
    objectRow?.NOMBRE_COMPLEMENTO,
    objectRow?.DESCRIPCION_OBJETO
  ]
    .map(normalizeHeader)
    .filter(Boolean);

  const ranked = menuRows
    .map((row) => {
      let score = 0;
      const targetObject = normalizeHeader(row.target.object);
      const sourceObject = normalizeHeader(row.source.object);
      const targetDescription = normalizeHeader(row.target.description);

      if (executableNames.includes(targetObject)) score += 10;
      if (objectNames.includes(targetObject)) score += 9;
      if (executableNames.includes(sourceObject)) score += 6;
      if (objectNames.includes(sourceObject)) score += 5;
      if (executableNames.includes(targetDescription)) score += 2;

      return { row, score };
    })
    .filter((item) => item.score > 0)
    .sort((a, b) => b.score - a.score);

  if (!ranked.length) return [];

  const bestScore = ranked[0].score;
  return ranked.filter((item) => item.score === bestScore).map((item) => item.row);
}

function buildMenuLineage(menuRows, seedRows, includePredecessors) {
  if (!includePredecessors) {
    const unique = [];
    const seen = new Set();

    seedRows.forEach((row) => {
      const menuId = cleanValue(row?.target?.menu);
      if (!menuId || seen.has(menuId)) return;
      seen.add(menuId);
      unique.push(row);
    });

    return unique;
  }

  const byTargetMenu = new Map(menuRows.map((row) => [cleanValue(row.target.menu), row]));
  const ordered = [];
  const visited = new Set();

  function visit(row) {
    if (!row) return;
    const menuId = cleanValue(row.target.menu);
    if (!menuId || visited.has(menuId)) return;

    const predecessor = cleanValue(row.target.predecessor);
    if (predecessor) {
      visit(byTargetMenu.get(predecessor));
    }

    visited.add(menuId);
    ordered.push(row);
  }

  seedRows.forEach(visit);
  return ordered;
}

function findObjectRow(objectRows, executableRow, menuRows) {
  const executableNames = [executableRow?.NOMBRE, executableRow?.EJECUTABLE]
    .map(normalizeHeader)
    .filter(Boolean);

  const menuNames = menuRows
    .flatMap((row) => [row.target.object, row.source.object])
    .map(normalizeHeader)
    .filter(Boolean);

  return (
    objectRows.find((row) => {
      const names = [row.NOMBRE_OBJETO, row.NOMBRE_COMPLEMENTO]
        .map(normalizeHeader)
        .filter(Boolean);
      return names.some((name) => executableNames.includes(name) || menuNames.includes(name));
    }) || null
  );
}

function findComponentRow(componentRows, secComponente) {
  const target = cleanValue(secComponente);
  return componentRows.find((row) => cleanValue(row.id) === target) || null;
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
/
`;
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
/
`;
}

function buildAatEjecutableSql(executableRow, componentRow) {
  const secEjecutable = cleanValue(executableRow.SEC_EJECUTABLE);
  const ejecutable = cleanValue(executableRow.EJECUTABLE);
  const secComponente = cleanValue(executableRow.SEC_COMPONENTE || componentRow?.id);
  const nombre = cleanValue(executableRow.NOMBRE);
  const estado = cleanValue(executableRow.ESTADO || "A");
  const tipoEjecutable = cleanValue(executableRow.TIPO);
  const secAtributo = sqlNumberOrDefault(executableRow.ATRIBUTO, 1);
  const descripcion = cleanValue(executableRow.DESCRIPCION || nombre);
  const nivelLog = sqlNumberOrDefault(executableRow.NIVEL, 0);

  return `PROMPT Insertando o actualizando en la tabla AAT_EJECUTABLE, el objeto ${ejecutable}...
DECLARE
    xml_fuente CLOB;
    existe NUMBER:=0;
BEGIN
	xml_fuente := TO_CLOB('');

	SELECT	COUNT(*)
        INTO	existe
	FROM
        aat_ejecutable
	WHERE	
        sec_ejecutable = ${secEjecutable};
		
	IF existe = 0 THEN
    INSERT INTO aat_ejecutable( 
        sec_ejecutable,
        ejecutable,
        sec_componente,
        nombre,
        estado,
        tipo_ejecutable,
        virtualizacion,
        fecha_creacion,
        fecha_ultima_actualizacion,
        fuente,
        sec_atributo,
        sec_icono,
        funcion_previa,
        funcion_posterior,
        descripcion,
        sec_ejecutable_reporte,
        nivel_log)
    VALUES (
    ${secEjecutable},
    '${escapeSqlText(ejecutable)}',
    ${secComponente || "NULL"},
    '${escapeSqlText(nombre)}',
    '${escapeSqlText(estado)}',
    '${escapeSqlText(tipoEjecutable)}',
    'N',
    sysdate,
    sysdate,
    to_clob(xml_fuente),
    ${secAtributo},
    NULL,
    NULL,
    NULL,
    '${escapeSqlText(descripcion)}',
    NULL,
    ${nivelLog});
    ELSE
		UPDATE aat_ejecutable
		SET 
            ejecutable = '${escapeSqlText(ejecutable)}',
            sec_componente = ${secComponente || "NULL"},
            nombre = '${escapeSqlText(nombre)}',
            estado = '${escapeSqlText(estado)}',
            tipo_ejecutable = '${escapeSqlText(tipoEjecutable)}',
            virtualizacion = 'N',
            fecha_ultima_actualizacion = sysdate,
            fuente =to_clob(xml_fuente),
            descripcion = '${escapeSqlText(descripcion)}'
		WHERE 
            sec_ejecutable = ${secEjecutable};	
	END IF;
	
	COMMIT;
EXCEPTION
	WHEN OTHERS THEN
		ROLLBACK;
		pk_excepcion.error_aplicacion;
END;
/
`;
}

function buildMstMenuSql(menuRow) {
  const menu = cleanValue(menuRow.target.menu);
  const descripcion = cleanValue(menuRow.target.description || menuRow.source.description || menuRow.target.object || menu);
  const orden = cleanValue(menuRow.target.order || menuRow.source.order || "0");
  const tipo = cleanValue(menuRow.target.type);
  const objeto = cleanValue(menuRow.target.object);
  const predecesor = cleanValue(menuRow.target.predecessor || menuRow.source.predecessor);
  const icono = cleanValue(menuRow.source.icon || menuRow.target.icon || "");

  return `PROMPT Insertando o actualizando en la tabla MST_MENU, el objeto ${objeto.toUpperCase()}...
DECLARE
	existe			NUMBER := 0;
BEGIN
	SELECT	COUNT(*)
	INTO	existe
	FROM	MST_MENU
	WHERE	
        menu = ${menu || "NULL"};
		
	IF existe = 0 THEN
       INSERT INTO MST_MENU  (MENU,DESCRIPCION,HIJO_ORDEN,TIPO_OBJETO_EJECUTABLE,NOMBRE_OBJETO_EJECUTABLE,MENU_PREDECESOR,ICONO) 
         VALUES ( ${sqlText(menu)}, ${sqlText(descripcion)} ,${sqlText(orden)},${sqlText(tipo)},${sqlText(objeto.toUpperCase())},${sqlText(predecesor)}, ${sqlText(icono || (tipo ? "REPORTE" : "FOLDER"))});    
	ELSE
		UPDATE MST_MENU
		SET 
            descripcion = ${sqlText(descripcion)},
            hijo_orden = ${sqlNumberOrDefault(orden, 0)},
            tipo_objeto_ejecutable = ${sqlText(tipo)},
            nombre_objeto_ejecutable = ${sqlText(objeto.toUpperCase())},
            menu_predecesor = ${sqlNumber(predecesor)},
            icono = ${sqlText(icono || (tipo ? "REPORTE" : "FOLDER"))}
		WHERE 
            menu = ${menu};
	END IF;
	
	COMMIT;
EXCEPTION
	WHEN OTHERS THEN
		ROLLBACK;
		pk_excepcion.error_aplicacion;
END;
/
`;
}

function buildMstObjetoSql(objectRow, executableRow, componentRow) {
  const tipoObjeto = cleanValue((objectRow?.TIPO_OBJETO || executableRow?.TIPO) ? "JR" : null);
  const nombreObjeto = cleanValue(objectRow?.NOMBRE_OBJETO || executableRow?.EJECUTABLE);
  const descripcionObjeto = cleanValue(executableRow?.DESCRIPCION || objectRow?.DESCRIPCION_OBJETO || executableRow?.NOMBRE);
  const nombreComplemento = cleanValue(objectRow?.NOMBRE_COMPLEMENTO || executableRow?.NOMBRE);
  const componente = buildComponentLabel(componentRow);

  return `PROMPT Insertando o actualizando en la tabla MST_OBJETO, el objeto ${nombreObjeto}...
DECLARE
	existe			NUMBER:=0;
BEGIN
	SELECT	COUNT(*)
	INTO	existe
	FROM	mst_objeto
	WHERE	nombre_objeto = ${sqlText(nombreObjeto)}
	AND TIPO_OBJETO = ${sqlText(tipoObjeto)};
		
	IF existe = 0 THEN
         INSERT INTO MST_OBJETO (TIPO_OBJETO,NOMBRE_OBJETO,REVISION,FECHA_INSTALACION,DESCRIPCION_OBJETO,COMPONENTE,NOMBRE_COMPLEMENTO,EJECUTA_BAT) 
                VALUES (${sqlText(tipoObjeto)},${sqlText(nombreObjeto)}, 0, SYSDATE, ${sqlText(descripcionObjeto)},${sqlText(componente)},${sqlText(nombreComplemento)},'N');
	END IF;
	
	COMMIT;
EXCEPTION
	WHEN OTHERS THEN
		ROLLBACK;
		pk_excepcion.error_aplicacion;
END;
/
`;
}

function buildMspPrivilegiosRolSql(objectRow, executableRow, roleName) {
  const rol = cleanValue(roleName || "ICEBERG_ZK");
  const miEjecutable = cleanValue(executableRow?.EJECUTABLE || objectRow?.NOMBRE_OBJETO);

  return `PROMPT Actualizando los permisos de las opciones de menú, el objeto ${miEjecutable}...
DECLARE
    mi_rol VARCHAR2(30)         := ${sqlText(rol)};
    mi_ejecutable VARCHAR2(30)  := ${sqlText(miEjecutable)};
BEGIN

    FOR x in (
        SELECT *
        FROM mst_menu
        WHERE 
            tipo_objeto_ejecutable IS NOT NULL
        AND nombre_objeto_Ejecutable = mi_ejecutable
        AND NOT EXISTS (
            SELECT 'X'
            FROM mst_rol_objeto_menu
            WHERE 
                rol = mi_rol
            AND nombre_objeto_ejecutable = mst_menu.nombre_objeto_ejecutable
            AND tipo_objeto_ejecutable = mst_menu.tipo_objeto_ejecutable
        )
    )
    LOOP
        msp_rol_objeto_menu.crear (
            un_nombre_objeto_ejecutable    => x.nombre_objeto_ejecutable,
            un_tipo_objeto_ejecutable      => x.tipo_objeto_ejecutable,
            un_rol                         => mi_rol,
            un_visible                     => 'S',
            un_leer                        => 'N',
            un_insertar                    => 'N',
            un_actualizar                  => 'N',
            un_borrar                      => 'N',
            un_ejecutar                    => 'N',
            un_por_objeto                  => 'S'
        );

        msp_rol_objeto_menu.insertar_privilegios (
            un_tipo_objeto_ejecutable        => x.tipo_objeto_ejecutable,
            un_nombre_objeto_ejecutable      => x.nombre_objeto_ejecutable,
            un_rol                           => mi_rol
        );
    END LOOP;

    msp_menu.reconstruye_menu;
    msp_menu.reconstruye_menu_rol(mi_rol);

    UPDATE mst_rol
    SET estado ='A'
    WHERE rol = mi_rol;

    COMMIT;

EXCEPTION
	WHEN OTHERS THEN
		ROLLBACK;
		pk_excepcion.error_aplicacion;
END;
/
`;
}

function buildMenuFileName(menuRow) {
  const base = toSlug(menuRow.target.object || menuRow.target.description || menuRow.source.description || menuRow.target.menu);
  return `opcion_menu/mst_menu_${base || cleanValue(menuRow.target.menu)}.sql`;
}

function buildObjectFileName(objectRow, executableRow) {
  const base = toSlug(objectRow?.NOMBRE_COMPLEMENTO || executableRow?.NOMBRE || executableRow?.EJECUTABLE);
  return `opcion_menu/mst_objeto_${base || cleanValue(executableRow?.EJECUTABLE)}.sql`;
}

function buildExecutableFileName(executableRow) {
  const base = toSlug(executableRow?.NOMBRE || executableRow?.EJECUTABLE);
  return `ejecutable/insert_ejecutable_${base || cleanValue(executableRow?.EJECUTABLE)}.sql`;
}

function buildPrivilegesFileName(executableRow) {
  const base = toSlug(executableRow?.NOMBRE || executableRow?.EJECUTABLE);
  return `opcion_menu/msp_privilegios_rol_${base || cleanValue(executableRow?.EJECUTABLE)}.sql`;
}

function buildWorkbookOutputs(context, roleName) {
  const parameterPairs = context.parameterPairs;
  const menuFiles = context.menuChain.map((menuRow) => ({
    path: buildMenuFileName(menuRow),
    content: buildMstMenuSql(menuRow)
  }));

  const executableFile = {
    path: buildExecutableFileName(context.executableRow),
    content: buildAatEjecutableSql(context.executableRow, context.componentRow)
  };

  const objectFile = {
    path: buildObjectFileName(context.objectRow, context.executableRow),
    content: buildMstObjetoSql(context.objectRow, context.executableRow, context.componentRow)
  };

  const privilegesFile = {
    path: buildPrivilegesFileName(context.executableRow),
    content: buildMspPrivilegiosRolSql(context.objectRow, context.executableRow, roleName)
  };

  return {
    parameterPairs,
    files: [...flattenFiles(parameterPairs), executableFile, ...menuFiles, objectFile, privilegesFile]
  };
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

function getNonParameterFiles(allFiles, parameterPairs) {
  const parameterPaths = new Set(parameterPairs.flatMap((pair) => pair.files.map((file) => file.path)));
  return allFiles.filter((file) => !parameterPaths.has(file.path));
}

function openFileInBrowser(file) {
  if (!file) return;

  const html = `<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>${escapeHtml(file.path)}</title>
  <style>
    body { font-family: Consolas, monospace; margin: 16px; background: #f7f7f7; }
    h1, h2 { font-family: Arial, sans-serif; }
    pre { background: #fff; padding: 12px; border: 1px solid #ddd; border-radius: 8px; overflow: auto; }
  </style>
</head>
<body>
  <h1>SQL Generado</h1>
  <h2>${escapeHtml(file.path)}</h2>
  <pre>${escapeHtml(file.content)}</pre>
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

function drawOtherFiles(items) {
  ui.otherRows.innerHTML = "";
  if (!items.length) return;

  const html = items
    .map((file, index) => `<tr>
      <td>${index + 1}</td>
      <td>${escapeHtml(file.path)}</td>
      <td>
        <button type="button" class="btn btn-mini btn-primary other-download" data-index="${index}">Descargar</button>
        <button type="button" class="btn btn-mini other-open" data-index="${index}">Abrir</button>
      </td>
    </tr>`)
    .join("");

  ui.otherRows.innerHTML = html;

  ui.otherRows.querySelectorAll(".other-download").forEach((button) => {
    button.addEventListener("click", () => {
      const idx = Number(button.dataset.index);
      const file = items[idx];
      downloadTextFile(file.path, file.content);
    });
  });

  ui.otherRows.querySelectorAll(".other-open").forEach((button) => {
    button.addEventListener("click", () => {
      const idx = Number(button.dataset.index);
      openFileInBrowser(items[idx]);
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
  const otherFiles = getNonParameterFiles(generatedFiles, generatedPairs);
  const fileList = generatedFiles.map((f,i) => `${i + 1}. ${f.path}`).join("\n");
  setStatus(
    [
      `Archivo: ${lastWorkbookName}`,
      `SEC_EJECUTABLE: ${secEjecutableTarget}`,
      `Rol privilegios: ${cleanValue(ui.rolPrivilegios.value) || "ICEBERG_ZK"}`,
      `Incluir menús predecesores: ${ui.includeMenuPredecessors?.checked ? "SI" : "NO"}`,
      `Ejecutable detectado: ${generatedContext?.executableRow?.EJECUTABLE || "N/A"}`,
      `Menu detectado: ${generatedContext?.menuChain?.length || 0}`,
      `Parametros encontrados: ${generatedPairs.length}`,
      `Otros archivos detectados: ${otherFiles.length}`,
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
  const roleName = cleanValue(ui.rolPrivilegios.value) || "ICEBERG_ZK";
  const includeMenuPredecessors = Boolean(ui.includeMenuPredecessors?.checked);

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
  generatedContext = null;
  updateButtonsState(false);
  ui.rows.innerHTML = "";
  ui.otherRows.innerHTML = "";

  try {
    setStatus("Leyendo archivo Excel...");
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    lastWorkbookName = file.name;

    const shEjecutable = findSheet(workbook, "Ejecutable");
    const shEjecutableParametro = findSheet(workbook, "Ejecutable_Parametro");
    const shParametro = findSheet(workbook, "Parametro");
    const shMenu = findSheet(workbook, "Menu");
    const shMstObjeto = findSheet(workbook, "MST_OBJETO");
    const shComponentes = findSheet(workbook, "Control IDs");

    if (!shEjecutable || !shEjecutableParametro || !shParametro || !shMenu || !shMstObjeto || !shComponentes) {
      throw new Error("No se encontraron las hojas requeridas (Ejecutable, Ejecutable_Parametro, Parametro, Menu, MST_OBJETO o Control IDs).");
    }

    const executableRows = sheetToObjects(shEjecutable, ["SEC_EJECUTABLE", "EJECUTABLE", "SEC_COMPONENTE", "NOMBRE"]);
    const ejecutableRows = sheetToObjects(shEjecutableParametro, ["SEC_EJECUTABLE", "SEC_PARAMETRO", "SEC_EJECUTABLE_PARAMETRO"]);
    const parametroRows = sheetToObjects(shParametro, ["SEC_PARAMETRO", "PARAMETRO", "NOMBRE", "TIPO_DATO", "DATOS"]);
    const menuRows = parseMenuSheet(shMenu);
    const objectRows = sheetToObjects(shMstObjeto, ["TIPO_OBJETO", "NOMBRE_OBJETO", "DESCRIPCION_OBJETO", "NOMBRE_COMPLEMENTO"]);
    const componentRows = parseComponentSheet(shComponentes);

    const executableRow = executableRows.find((row) => cleanValue(row.SEC_EJECUTABLE) === secEjecutableTarget);

    if (!executableRow) {
      throw new Error(`No se encontró el ejecutable con SEC_EJECUTABLE = ${secEjecutableTarget}.`);
    }

    const relatedExecutableRows = ejecutableRows.filter((row) => cleanValue(row.SEC_EJECUTABLE) === secEjecutableTarget);

    if (!relatedExecutableRows.length) {
      setStatus(`No se encontraron filas en Ejecutable_Parametro para SEC_EJECUTABLE = ${secEjecutableTarget}.`);
    }

    const paramsBySec = new Map(parametroRows.map((row) => [cleanValue(row.SEC_PARAMETRO), row]));
    const componentRow = findComponentRow(componentRows, executableRow.SEC_COMPONENTE);
    const preliminaryObjectRow = findObjectRow(objectRows, executableRow, []);
    const menuSeedRows = findBestMenuRows(menuRows, executableRow, preliminaryObjectRow);
    const menuChain = buildMenuLineage(menuRows, menuSeedRows, includeMenuPredecessors);
    const objectRow = findObjectRow(objectRows, executableRow, menuChain) || preliminaryObjectRow;

    const filtered = relatedExecutableRows
      .map((row) => ({
        ejecutable: row,
        parametro: paramsBySec.get(cleanValue(row.SEC_PARAMETRO))
      }))
      .filter((pair) => pair.parametro);

    if (!filtered.length) {
      setStatus(`No hay parámetros para SEC_EJECUTABLE = ${secEjecutableTarget}; se generarán los demás scripts.`);
    }

    const sortedParameterPairs = sortPairs(buildGeneratedPairs(filtered), cleanValue(ui.sortBy.value).toLowerCase());
    generatedContext = {
      executableRow,
      componentRow,
      objectRow,
      menuChain,
      includeMenuPredecessors,
      parameterPairs: sortedParameterPairs
    };

    const outputs = buildWorkbookOutputs(generatedContext, roleName);
    generatedPairs = outputs.parameterPairs;
    generatedFiles = outputs.files;
    const otherFiles = getNonParameterFiles(generatedFiles, generatedPairs);

    drawRows(generatedPairs);
    drawOtherFiles(otherFiles);
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
  if (!generatedContext) return;
  generatedContext.parameterPairs = sortPairs(generatedContext.parameterPairs, cleanValue(ui.sortBy.value).toLowerCase());
  const outputs = buildWorkbookOutputs(generatedContext, cleanValue(ui.rolPrivilegios.value) || "ICEBERG_ZK");
  generatedPairs = outputs.parameterPairs;
  generatedFiles = outputs.files;
  const otherFiles = getNonParameterFiles(generatedFiles, generatedPairs);
  drawRows(generatedPairs);
  drawOtherFiles(otherFiles);
  refreshStatus(cleanValue(ui.secEjecutable.value));
});
