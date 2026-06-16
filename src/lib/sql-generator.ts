import * as XLSX from "xlsx";

export type SortMode = "ordenamiento" | "sec_ejecutable_parametro" | "sec_parametro";

export type GeneratedFile = {
  path: string;
  content: string;
};

export type GeneratedPair = {
  ejecutable: Record<string, string>;
  parametro: Record<string, string>;
  files: [GeneratedFile, GeneratedFile];
};

export type MenuRow = {
  source: {
    menu: string;
    description: string;
    order: string;
    type: string;
    object: string;
    predecessor: string;
    icon: string;
  };
  target: {
    menu: string;
    description: string;
    order: string;
    type: string;
    object: string;
    predecessor: string;
    icon?: string;
  };
};

export type GenerationContext = {
  workbookName: string;
  secEjecutable: string;
  roleName: string;
  includeMenuPredecessors: boolean;
  executableRow: Record<string, string>;
  componentRow: Record<string, string> | null;
  objectRow: Record<string, string> | null;
  menuChain: MenuRow[];
  parameterPairs: GeneratedPair[];
};

export type WorkbookArtifacts = {
  context: GenerationContext;
  parameterPairs: GeneratedPair[];
  files: GeneratedFile[];
  otherFiles: GeneratedFile[];
};

export type GeneratedContext = GenerationContext;
export type SqlFile = GeneratedFile;

type SheetRow = Record<string, string>;

function cleanValue(value: unknown) {
  return String(value ?? "").trim();
}

function stripDiacritics(text: string) {
  return String(text || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function normalizeHeader(value: unknown) {
  return stripDiacritics(String(value || "").trim()).toUpperCase();
}

function toSlug(value: unknown) {
  return stripDiacritics(cleanValue(value))
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function escapeSqlText(value: unknown) {
  return String(value ?? "").replace(/'/g, "''");
}

function maybeNullSql(value: unknown) {
  const text = cleanValue(value);
  return text === "" ? "NULL" : `'${escapeSqlText(text)}'`;
}

function toUpperWithoutParentheses(text: unknown) {
  return cleanValue(text)
    .replace(/[_-]+/g, " ")
    .replace(/\s*\([^)]*\)\s*/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function detectArticulo(nombre: unknown) {
  const normalized = stripDiacritics(toUpperWithoutParentheses(nombre));
  const tokens = normalized.split(" ").filter(Boolean);
  const stopWords = new Set(["DE", "DEL", "LA", "EL", "LOS", "LAS", "Y", "E", "O", "U"]);
  const noun = tokens.find((token) => !stopWords.has(token)) || "";

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

  const masculineWords = new Set(["DIA", "DOCUMENTO", "FONDO", "CENTRO", "CODIGO", "NUMERO", "ESTADO", "TIPO"]);

  if (feminineWords.has(noun)) return "LA";
  if (masculineWords.has(noun)) return "EL";

  const masculineEndingExceptions = new Set(["PROBLEMA", "SISTEMA", "TEMA", "MAPA", "PROGRAMA"]);
  if (noun.endsWith("A") && !masculineEndingExceptions.has(noun)) return "LA";
  return "EL";
}

function buildParametroDescripcion(nombre: unknown) {
  const articulo = detectArticulo(nombre);
  return `PARÁMETRO PARA CONSULTAR POR ${articulo} ${toUpperWithoutParentheses(nombre)}.`;
}

function buildEjecutableDescripcion(nombre: unknown) {
  const articulo = detectArticulo(nombre);
  return `ASOCIA PARÁMETRO PARA CONSULTAR POR ${articulo} ${toUpperWithoutParentheses(nombre)}.`;
}

function toSortableNumber(value: unknown) {
  const n = Number(cleanValue(value));
  return Number.isFinite(n) ? n : Number.MAX_SAFE_INTEGER;
}

function sqlText(value: unknown) {
  const text = cleanValue(value);
  return text === "" ? "NULL" : `'${escapeSqlText(text)}'`;
}

function sqlNumber(value: unknown) {
  const text = cleanValue(value);
  return text === "" ? "NULL" : text;
}

function sqlNumberOrDefault(value: unknown, defaultValue: number) {
  const text = cleanValue(value);
  return text === "" ? String(defaultValue) : text;
}

function escapeHtml(value: unknown) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function detectHeaderRow(rawRows: unknown[][], requiredHeaders: string[]) {
  const maxRowsToInspect = Math.min(rawRows.length, 20);
  for (let i = 0; i < maxRowsToInspect; i++) {
    const row = rawRows[i].map(normalizeHeader);
    const matchesAll = requiredHeaders.every((header) => row.includes(header));
    if (matchesAll) return i;
  }
  return -1;
}

function sheetToObjects(sheet: XLSX.WorkSheet, requiredHeaders: string[]) {
  const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false, defval: "" }) as unknown[][];
  const headerRowIndex = detectHeaderRow(rawRows, requiredHeaders);

  if (headerRowIndex === -1) {
    throw new Error(`No se encontraron encabezados requeridos: ${requiredHeaders.join(", ")}`);
  }

  const headers = rawRows[headerRowIndex].map((header) => normalizeHeader(header));
  const dataRows = rawRows.slice(headerRowIndex + 1);

  return dataRows
    .map((row) => {
      const obj: SheetRow = {};
      headers.forEach((header, index) => {
        obj[header] = cleanValue(row[index]);
      });
      return obj;
    })
    .filter((row) => Object.values(row).some((value) => cleanValue(value) !== ""));
}

function getSheetRows(sheet: XLSX.WorkSheet) {
  return XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false, defval: "" }) as unknown[][];
}

function findSheet(workbook: XLSX.WorkBook, expectedName: string) {
  const wanted = expectedName.toUpperCase();
  const exact = workbook.SheetNames.find((name) => normalizeHeader(name) === wanted);
  if (exact) return workbook.Sheets[exact];

  const loose = workbook.SheetNames.find((name) => normalizeHeader(name).includes(wanted));
  if (loose) return workbook.Sheets[loose];

  return null;
}

function parseComponentSheet(sheet: XLSX.WorkSheet) {
  return getSheetRows(sheet)
    .slice(2)
    .map((row) => ({
      id: cleanValue(row[0]),
      componente: cleanValue(row[1]),
      nombre: cleanValue(row[2])
    }))
    .filter((row) => row.id !== "" || row.componente !== "" || row.nombre !== "");
}

function parseMenuSheet(sheet: XLSX.WorkSheet) {
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
    .filter(
      (row) =>
        Object.values(row.source).some((value) => value !== "") ||
        Object.values(row.target).some((value) => value !== "")
    );
}

function findBestMenuRows(menuRows: MenuRow[], executableRow: SheetRow, objectRow: SheetRow | null) {
  const executableNames = [executableRow?.NOMBRE, executableRow?.EJECUTABLE, executableRow?.DESCRIPCION]
    .map(normalizeHeader)
    .filter(Boolean);

  const objectNames = [objectRow?.NOMBRE_OBJETO, objectRow?.NOMBRE_COMPLEMENTO, objectRow?.DESCRIPCION_OBJETO]
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

function buildMenuLineage(menuRows: MenuRow[], seedRows: MenuRow[], includePredecessors: boolean) {
  if (!includePredecessors) {
    const unique: MenuRow[] = [];
    const seen = new Set<string>();

    seedRows.forEach((row) => {
      const menuId = cleanValue(row?.target?.menu);
      if (!menuId || seen.has(menuId)) return;
      seen.add(menuId);
      unique.push(row);
    });

    return unique;
  }

  const byTargetMenu = new Map(menuRows.map((row) => [cleanValue(row.target.menu), row]));
  const ordered: MenuRow[] = [];
  const visited = new Set<string>();

  function visit(row: MenuRow | undefined) {
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

function findObjectRow(objectRows: SheetRow[], executableRow: SheetRow, menuRows: MenuRow[]) {
  const executableNames = [executableRow?.NOMBRE, executableRow?.EJECUTABLE]
    .map(normalizeHeader)
    .filter(Boolean);

  const menuNames = menuRows
    .flatMap((row) => [row.target.object, row.source.object])
    .map(normalizeHeader)
    .filter(Boolean);

  return (
    objectRows.find((row) => {
      const names = [row.NOMBRE_OBJETO, row.NOMBRE_COMPLEMENTO].map(normalizeHeader).filter(Boolean);
      return names.some((name) => executableNames.includes(name) || menuNames.includes(name));
    }) || null
  );
}

function findComponentRow(componentRows: Array<{ id: string }>, secComponente: unknown) {
  const target = cleanValue(secComponente);
  return componentRows.find((row) => cleanValue(row.id) === target) || null;
}

function buildParametroSql(paramRow: SheetRow) {
  const secParametro = cleanValue(paramRow.SEC_PARAMETRO);
  const parametro = cleanValue(paramRow.PARAMETRO);
  const nombre = cleanValue(paramRow.NOMBRE);
  const estado = cleanValue(paramRow.ESTADO || "A");
  const tipoDato = cleanValue(paramRow.TIPO_DATO);
  const datos = cleanValue(paramRow.DATOS);
  const descripcion = buildParametroDescripcion(nombre);

  const hasQuery = datos !== "";
  const queryBlock = hasQuery ? `    mi_query        VARCHAR2(2000);\n` : "";
  const queryAssign = hasQuery ? `\n    mi_query := '${escapeSqlText(datos)}';\n` : "";
  const queryValue = hasQuery ? "mi_query" : "NULL";

  return `PROMPT Insertando o actualizando en la tabla SRT_PARAMETRO, el objeto ${parametro}...\nDECLARE\n\tmi_existe_registro\tNUMBER := 0;\n${queryBlock}BEGIN\n\n\tSELECT COUNT(*)\n    INTO mi_existe_registro\n\tFROM srt_parametro\n\tWHERE sec_parametro = ${secParametro};\n${queryAssign}\n\tIF mi_existe_registro = 0 THEN\n        INSERT INTO srt_parametro  (\n            sec_parametro,\n            parametro,\n            nombre,\n            estado,\n            tipo_dato,\n            datos,\n            descripcion\n        ) \n        VALUES ( \n            ${secParametro}, \n            '${escapeSqlText(parametro)}',\n            '${escapeSqlText(nombre)}',\n            '${escapeSqlText(estado)}',\n            '${escapeSqlText(tipoDato)}',\n            ${queryValue},\n            '${escapeSqlText(descripcion)}'\n        );\n\tELSE\n\t\tUPDATE srt_parametro\n\t\tSET \n            parametro       = '${escapeSqlText(parametro)}',\n            nombre          = '${escapeSqlText(nombre)}',\n            estado          = '${escapeSqlText(estado)}',\n            tipo_dato       = '${escapeSqlText(tipoDato)}',\n            datos           = ${queryValue},\n            descripcion     = '${escapeSqlText(descripcion)}'\n\t\tWHERE \n            sec_parametro = ${secParametro};\t\n\tEND IF;\n\t\n\tCOMMIT;\n    \nEXCEPTION\n\tWHEN OTHERS THEN\n\t\tROLLBACK;\n\t\tpk_excepcion.error_aplicacion;\nEND;\n/\n`;
}

function buildEjecutableSql(epRow: SheetRow, paramRow: SheetRow) {
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

  return `PROMPT Insertando o actualizando en la tabla SRT_EJECUTABLE_PARAMETRO, el objeto ${parametro}...\nDECLARE\n\tmi_existe_registro\tNUMBER := 0;\nBEGIN\n\n\tSELECT COUNT(*)\n    INTO mi_existe_registro\n\tFROM srt_ejecutable_parametro\n\tWHERE\t\n        sec_ejecutable_parametro = ${secEP};\n\t\t\n\tIF mi_existe_registro = 0 THEN\n        INSERT INTO srt_ejecutable_parametro (\n            sec_ejecutable_parametro,\n            sec_ejecutable,\n            sec_parametro,\n            requerido,\n            valor_asumido,\n            seleccion_multiple,\n            ordenamiento,\n            agrupador,\n            descripcion\n        ) \n        VALUES (\n            ${secEP},\n            ${secEjecutable}, \n            ${secParametro},\n            '${escapeSqlText(requerido)}',\n            ${maybeNullSql(valorAsumido)},\n            '${escapeSqlText(seleccionMultiple)}',\n            ${ordenamiento},\n            ${maybeNullSql(agrupador)},\n            '${escapeSqlText(descripcion)}'\n        );\n\tELSE\n\t\tUPDATE srt_ejecutable_parametro\n\t\tSET \n            sec_ejecutable      = ${secEjecutable},\n            sec_parametro       = ${secParametro},\n            requerido           = '${escapeSqlText(requerido)}',\n            valor_asumido       = ${maybeNullSql(valorAsumido)},\n            seleccion_multiple  = '${escapeSqlText(seleccionMultiple)}',\n            ordenamiento        = ${ordenamiento},\n            agrupador           = ${maybeNullSql(agrupador)},\n            descripcion         = '${escapeSqlText(descripcion)}'\n\t\tWHERE \n            sec_ejecutable_parametro = ${secEP};\t\n\tEND IF;\n\t\n\tCOMMIT;\n    \nEXCEPTION\n\tWHEN OTHERS THEN\n\t\tROLLBACK;\n\t\tpk_excepcion.error_aplicacion;\nEND;\n/\n`;
}

function buildAatEjecutableSql(executableRow: SheetRow, componentRow: SheetRow | null) {
  const secEjecutable = cleanValue(executableRow.SEC_EJECUTABLE);
  const ejecutable = cleanValue(executableRow.EJECUTABLE);
  const secComponente = cleanValue(executableRow.SEC_COMPONENTE || componentRow?.id);
  const nombre = cleanValue(executableRow.NOMBRE);
  const estado = cleanValue(executableRow.ESTADO || "A");
  const tipoEjecutable = cleanValue(executableRow.TIPO);
  const secAtributo = sqlNumberOrDefault(executableRow.ATRIBUTO, 1);
  const descripcion = cleanValue(executableRow.DESCRIPCION || nombre);
  const nivelLog = sqlNumberOrDefault(executableRow.NIVEL, 0);

  return `PROMPT Insertando o actualizando en la tabla AAT_EJECUTABLE, el objeto ${ejecutable}...\nDECLARE\n    xml_fuente CLOB;\n    existe NUMBER:=0;\nBEGIN\n\txml_fuente := TO_CLOB('');\n\n\tSELECT\tCOUNT(*)\n        INTO\texiste\n\tFROM\n        aat_ejecutable\n\tWHERE\t\n        sec_ejecutable = ${secEjecutable};\n\t\t\n\tIF existe = 0 THEN\n    INSERT INTO aat_ejecutable( \n        sec_ejecutable,\n        ejecutable,\n        sec_componente,\n        nombre,\n        estado,\n        tipo_ejecutable,\n        virtualizacion,\n        fecha_creacion,\n        fecha_ultima_actualizacion,\n        fuente,\n        sec_atributo,\n        sec_icono,\n        funcion_previa,\n        funcion_posterior,\n        descripcion,\n        sec_ejecutable_reporte,\n        nivel_log)\n    VALUES (\n    ${secEjecutable},\n    '${escapeSqlText(ejecutable)}',\n    ${secComponente || "NULL"},\n    '${escapeSqlText(nombre)}',\n    '${escapeSqlText(estado)}',\n    '${escapeSqlText(tipoEjecutable)}',\n    'N',\n    sysdate,\n    sysdate,\n    to_clob(xml_fuente),\n    ${secAtributo},\n    NULL,\n    NULL,\n    NULL,\n    '${escapeSqlText(descripcion)}',\n    NULL,\n    ${nivelLog});\n    ELSE\n\t\tUPDATE aat_ejecutable\n\t\tSET \n            ejecutable = '${escapeSqlText(ejecutable)}',\n            sec_componente = ${secComponente || "NULL"},\n            nombre = '${escapeSqlText(nombre)}',\n            estado = '${escapeSqlText(estado)}',\n            tipo_ejecutable = '${escapeSqlText(tipoEjecutable)}',\n            virtualizacion = 'N',\n            fecha_ultima_actualizacion = sysdate,\n            fuente =to_clob(xml_fuente),\n            descripcion = '${escapeSqlText(descripcion)}'\n\t\tWHERE \n            sec_ejecutable = ${secEjecutable};\t\n\tEND IF;\n\t\n\tCOMMIT;\nEXCEPTION\n\tWHEN OTHERS THEN\n\t\tROLLBACK;\n\t\tpk_excepcion.error_aplicacion;\nEND;\n/\n`;
}

function buildMstMenuSql(menuRow: MenuRow) {
  const menu = cleanValue(menuRow.target.menu);
  const descripcion = cleanValue(menuRow.target.description || menuRow.source.description || menuRow.target.object || menu);
  const orden = cleanValue(menuRow.target.order || menuRow.source.order || "0");
  const tipo = cleanValue(menuRow.target.type);
  const objeto = cleanValue(menuRow.target.object);
  const predecesor = cleanValue(menuRow.target.predecessor || menuRow.source.predecessor);
  const icono = cleanValue(menuRow.source.icon || menuRow.target.icon || "");

  return `PROMPT Insertando o actualizando en la tabla MST_MENU, el objeto ${objeto.toUpperCase()}...\nDECLARE\n\texiste\t\t\tNUMBER := 0;\nBEGIN\n\tSELECT\tCOUNT(*)\n\tINTO\texiste\n\tFROM\tMST_MENU\n\tWHERE\t\n        menu = ${menu || "NULL"};\n\t\t\n\tIF existe = 0 THEN\n       INSERT INTO MST_MENU  (MENU,DESCRIPCION,HIJO_ORDEN,TIPO_OBJETO_EJECUTABLE,NOMBRE_OBJETO_EJECUTABLE,MENU_PREDECESOR,ICONO) \n         VALUES ( ${sqlText(menu)}, ${sqlText(descripcion)} ,${sqlText(orden)},${sqlText(tipo)},${sqlText(objeto.toUpperCase())},${sqlText(predecesor)}, ${sqlText(icono || (tipo ? "REPORTE" : "FOLDER"))});    \n\tELSE\n\t\tUPDATE MST_MENU\n\t\tSET \n            descripcion = ${sqlText(descripcion)},\n            hijo_orden = ${sqlNumberOrDefault(orden, 0)},\n            tipo_objeto_ejecutable = ${sqlText(tipo)},\n            nombre_objeto_ejecutable = ${sqlText(objeto.toUpperCase())},\n            menu_predecesor = ${sqlNumber(predecesor)},\n            icono = ${sqlText(icono || (tipo ? "REPORTE" : "FOLDER"))}\n\t\tWHERE \n            menu = ${menu};\n\tEND IF;\n\t\n\tCOMMIT;\nEXCEPTION\n\tWHEN OTHERS THEN\n\t\tROLLBACK;\n\t\tpk_excepcion.error_aplicacion;\nEND;\n/\n`;
}

function buildMstObjetoSql(objectRow: SheetRow | null, executableRow: SheetRow, componentRow: SheetRow | null) {
  const tipoObjeto = cleanValue((objectRow?.TIPO_OBJETO || executableRow?.TIPO) ? "JR" : null);
  const nombreObjeto = cleanValue(objectRow?.NOMBRE_OBJETO || executableRow?.EJECUTABLE);
  const descripcionObjeto = cleanValue(executableRow?.DESCRIPCION || objectRow?.DESCRIPCION_OBJETO || executableRow?.NOMBRE);
  const nombreComplemento = cleanValue(objectRow?.NOMBRE_COMPLEMENTO || executableRow?.NOMBRE);
  const componente = cleanValue(componentRow?.nombre || componentRow?.componente || componentRow?.id || "ICEBERG");

  return `PROMPT Insertando o actualizando en la tabla MST_OBJETO, el objeto ${nombreObjeto}...\nDECLARE\n\texiste\t\t\tNUMBER:=0;\nBEGIN\n\tSELECT\tCOUNT(*)\n\tINTO\texiste\n\tFROM\tmst_objeto\n\tWHERE\tnombre_objeto = ${sqlText(nombreObjeto)}\n\tAND TIPO_OBJETO = ${sqlText(tipoObjeto)};\n\t\t\n\tIF existe = 0 THEN\n         INSERT INTO MST_OBJETO (TIPO_OBJETO,NOMBRE_OBJETO,REVISION,FECHA_INSTALACION,DESCRIPCION_OBJETO,COMPONENTE,NOMBRE_COMPLEMENTO,EJECUTA_BAT) \n                VALUES (${sqlText(tipoObjeto)},${sqlText(nombreObjeto)}, 0, SYSDATE, ${sqlText(descripcionObjeto)},${sqlText(componente)},${sqlText(nombreComplemento)},'N');\n\tEND IF;\n\t\n\tCOMMIT;\nEXCEPTION\n\tWHEN OTHERS THEN\n\t\tROLLBACK;\n\t\tpk_excepcion.error_aplicacion;\nEND;\n/\n`;
}

function buildMspPrivilegiosRolSql(objectRow: SheetRow | null, executableRow: SheetRow, roleName: string) {
  const rol = cleanValue(roleName || "ICEBERG_ZK");
  const miEjecutable = cleanValue(executableRow?.EJECUTABLE || objectRow?.NOMBRE_OBJETO);

  return `PROMPT Actualizando los permisos de las opciones de menú, el objeto ${miEjecutable}...\nDECLARE\n    mi_rol VARCHAR2(30)         := ${sqlText(rol)};\n    mi_ejecutable VARCHAR2(30)  := ${sqlText(miEjecutable)};\nBEGIN\n\n    FOR x in (\n        SELECT *\n        FROM mst_menu\n        WHERE \n            tipo_objeto_ejecutable IS NOT NULL\n        AND nombre_objeto_Ejecutable = mi_ejecutable\n        AND NOT EXISTS (\n            SELECT 'X'\n            FROM mst_rol_objeto_menu\n            WHERE \n                rol = mi_rol\n            AND nombre_objeto_ejecutable = mst_menu.nombre_objeto_ejecutable\n            AND tipo_objeto_ejecutable = mst_menu.tipo_objeto_ejecutable\n        )\n    )\n    LOOP\n        msp_rol_objeto_menu.crear (\n            un_nombre_objeto_ejecutable    => x.nombre_objeto_ejecutable,\n            un_tipo_objeto_ejecutable      => x.tipo_objeto_ejecutable,\n            un_rol                         => mi_rol,\n            un_visible                     => 'S',\n            un_leer                        => 'N',\n            un_insertar                    => 'N',\n            un_actualizar                  => 'N',\n            un_borrar                      => 'N',\n            un_ejecutar                    => 'N',\n            un_por_objeto                  => 'S'\n        );\n\n        msp_rol_objeto_menu.insertar_privilegios (\n            un_tipo_objeto_ejecutable        => x.tipo_objeto_ejecutable,\n            un_nombre_objeto_ejecutable      => x.nombre_objeto_ejecutable,\n            un_rol                           => mi_rol\n        );\n    END LOOP;\n\n    msp_menu.reconstruye_menu;\n    msp_menu.reconstruye_menu_rol(mi_rol);\n\n    UPDATE mst_rol\n    SET estado ='A'\n    WHERE rol = mi_rol;\n\n    COMMIT;\n\nEXCEPTION\n\tWHEN OTHERS THEN\n\t\tROLLBACK;\n\t\tpk_excepcion.error_aplicacion;\nEND;\n/\n`;
}

function buildMenuFileName(menuRow: MenuRow) {
  const base = toSlug(menuRow.target.object || menuRow.target.description || menuRow.source.description || menuRow.target.menu);
  return `opcion_menu/mst_menu_${base || cleanValue(menuRow.target.menu)}.sql`;
}

function buildObjectFileName(objectRow: SheetRow | null, executableRow: SheetRow) {
  const base = toSlug(objectRow?.NOMBRE_COMPLEMENTO || executableRow?.NOMBRE || executableRow?.EJECUTABLE);
  return `opcion_menu/mst_objeto_${base || cleanValue(executableRow?.EJECUTABLE)}.sql`;
}

function buildExecutableFileName(executableRow: SheetRow) {
  const base = toSlug(executableRow?.NOMBRE || executableRow?.EJECUTABLE);
  return `ejecutable/insert_ejecutable_${base || cleanValue(executableRow?.EJECUTABLE)}.sql`;
}

function buildPrivilegesFileName(executableRow: SheetRow) {
  const base = toSlug(executableRow?.NOMBRE || executableRow?.EJECUTABLE);
  return `opcion_menu/msp_privilegios_rol_${base || cleanValue(executableRow?.EJECUTABLE)}.sql`;
}

function buildParametroFileName(parametro: unknown) {
  const base = toSlug(parametro);
  return `parametro/${base || cleanValue(parametro)}.sql`;
}

function buildGeneratedPairs(filteredPairs: Array<{ ejecutable: SheetRow; parametro: SheetRow }>) {
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
      ] as [GeneratedFile, GeneratedFile]
    };
  });
}

function flattenFiles(pairs: GeneratedPair[]) {
  return pairs.flatMap((pair) => pair.files);
}

export function sortPairs(pairs: GeneratedPair[], mode: SortMode) {
  const cloned = [...pairs];

  if (mode === "sec_ejecutable_parametro") {
    cloned.sort(
      (a, b) => toSortableNumber(a.ejecutable.SEC_EJECUTABLE_PARAMETRO) - toSortableNumber(b.ejecutable.SEC_EJECUTABLE_PARAMETRO)
    );
    return cloned;
  }

  if (mode === "sec_parametro") {
    cloned.sort((a, b) => toSortableNumber(a.parametro.SEC_PARAMETRO) - toSortableNumber(b.parametro.SEC_PARAMETRO));
    return cloned;
  }

  cloned.sort((a, b) => toSortableNumber(a.ejecutable.ORDENAMIENTO) - toSortableNumber(b.ejecutable.ORDENAMIENTO));
  return cloned;
}

export function getNonParameterFiles(allFiles: GeneratedFile[], parameterPairs: GeneratedPair[]) {
  const parameterPaths = new Set(parameterPairs.flatMap((pair) => pair.files.map((file) => file.path)));
  return allFiles.filter((file) => !parameterPaths.has(file.path));
}

function buildWorkbookOutputs(context: GenerationContext, roleName: string, sortBy: SortMode) {
  const parameterPairs = sortPairs(context.parameterPairs, sortBy);
  const menuFiles = context.menuChain.map((menuRow) => ({
    path: buildMenuFileName(menuRow),
    content: buildMstMenuSql(menuRow)
  }));

  const executableFile: GeneratedFile = {
    path: buildExecutableFileName(context.executableRow),
    content: buildAatEjecutableSql(context.executableRow, context.componentRow)
  };

  const objectFile: GeneratedFile = {
    path: buildObjectFileName(context.objectRow, context.executableRow),
    content: buildMstObjetoSql(context.objectRow, context.executableRow, context.componentRow)
  };

  const privilegesFile: GeneratedFile = {
    path: buildPrivilegesFileName(context.executableRow),
    content: buildMspPrivilegiosRolSql(context.objectRow, context.executableRow, roleName)
  };

  return {
    parameterPairs,
    files: [...flattenFiles(parameterPairs), executableFile, ...menuFiles, objectFile, privilegesFile]
  };
}

export function parseWorkbook(
  file: File,
  secEjecutable: string,
  roleName: string,
  includeMenuPredecessors: boolean,
  sortBy: SortMode
) {
  return file.arrayBuffer().then((buffer) => {
    const workbook = XLSX.read(buffer, { type: "array" });
    const workbookName = file.name;

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

    const executableRow = executableRows.find((row) => cleanValue(row.SEC_EJECUTABLE) === secEjecutable);

    if (!executableRow) {
      throw new Error(`No se encontró el ejecutable con SEC_EJECUTABLE = ${secEjecutable}.`);
    }

    const relatedExecutableRows = ejecutableRows.filter((row) => cleanValue(row.SEC_EJECUTABLE) === secEjecutable);
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
      .filter((pair): pair is { ejecutable: SheetRow; parametro: SheetRow } => Boolean(pair.parametro));

    const sortedParameterPairs = sortPairs(buildGeneratedPairs(filtered), sortBy);

    const context: GenerationContext = {
      workbookName,
      secEjecutable,
      roleName,
      includeMenuPredecessors,
      executableRow,
      componentRow,
      objectRow,
      menuChain,
      parameterPairs: sortedParameterPairs
    };

    const outputs = buildWorkbookOutputs(context, roleName, sortBy);
    const otherFiles = getNonParameterFiles(outputs.files, outputs.parameterPairs);

    return {
      workbookName,
      secEjecutable,
      roleName,
      generatedFiles: outputs.files,
      generatedPairs: outputs.parameterPairs,
      generatedContext: context,
      otherFiles
    };
  });
}

export function getDownloadName(path: string) {
  return toDownloadFileName(path);
}

export function getNonParameterOutputFiles(allFiles: GeneratedFile[], parameterPairs: GeneratedPair[]) {
  return getNonParameterFiles(allFiles, parameterPairs);
}

export async function buildWorkbookArtifacts(file: File, options: {
  secEjecutable: string;
  roleName: string;
  includeMenuPredecessors: boolean;
  sortBy: SortMode;
}) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const workbookName = file.name;

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
  const executableParamRows = sheetToObjects(shEjecutableParametro, ["SEC_EJECUTABLE", "SEC_PARAMETRO", "SEC_EJECUTABLE_PARAMETRO"]);
  const parametroRows = sheetToObjects(shParametro, ["SEC_PARAMETRO", "PARAMETRO", "NOMBRE", "TIPO_DATO", "DATOS"]);
  const menuRows = parseMenuSheet(shMenu);
  const objectRows = sheetToObjects(shMstObjeto, ["TIPO_OBJETO", "NOMBRE_OBJETO", "DESCRIPCION_OBJETO", "NOMBRE_COMPLEMENTO"]);
  const componentRows = parseComponentSheet(shComponentes);

  const executableRow = executableRows.find((row) => cleanValue(row.SEC_EJECUTABLE) === options.secEjecutable);
  if (!executableRow) {
    throw new Error(`No se encontró el ejecutable con SEC_EJECUTABLE = ${options.secEjecutable}.`);
  }

  const relatedExecutableRows = executableParamRows.filter((row) => cleanValue(row.SEC_EJECUTABLE) === options.secEjecutable);
  const paramsBySec = new Map(parametroRows.map((row) => [cleanValue(row.SEC_PARAMETRO), row]));
  const componentRow = findComponentRow(componentRows, executableRow.SEC_COMPONENTE) as SheetRow | null;
  const preliminaryObjectRow = findObjectRow(objectRows, executableRow, []);
  const menuSeedRows = findBestMenuRows(menuRows, executableRow, preliminaryObjectRow);
  const menuChain = buildMenuLineage(menuRows, menuSeedRows, options.includeMenuPredecessors);
  const objectRow = findObjectRow(objectRows, executableRow, menuChain) || preliminaryObjectRow;

  const filtered = relatedExecutableRows
    .map((row) => ({
      ejecutable: row,
      parametro: paramsBySec.get(cleanValue(row.SEC_PARAMETRO))
    }))
    .filter((pair): pair is { ejecutable: SheetRow; parametro: SheetRow } => Boolean(pair.parametro));

  const parameterPairs = buildGeneratedPairs(filtered);
  const context: GenerationContext = {
    workbookName,
    secEjecutable: options.secEjecutable,
    roleName: options.roleName,
    includeMenuPredecessors: options.includeMenuPredecessors,
    executableRow,
    componentRow,
    objectRow,
    menuChain,
    parameterPairs
  };

  const outputs = buildWorkbookOutputs(context, options.roleName, options.sortBy);
  const otherFiles = getNonParameterFiles(outputs.files, outputs.parameterPairs);

  return {
    context,
    parameterPairs: outputs.parameterPairs,
    files: outputs.files,
    otherFiles
  } as WorkbookArtifacts;
}

export function getWorkbookStatusLines(args: {
  workbookName: string;
  secEjecutable: string;
  roleName: string;
  includeMenuPredecessors: boolean;
  executableRow: Record<string, string> | null;
  menuChain: MenuRow[];
  parameterPairs: GeneratedPair[];
  files: GeneratedFile[];
  otherFiles: GeneratedFile[];
}) {
  const fileList = args.files.map((file, index) => `${index + 1}. ${file.path}`).join("\n");

  return [
    `Archivo: ${args.workbookName}`,
    `SEC_EJECUTABLE: ${args.secEjecutable}`,
    `Rol privilegios: ${args.roleName}`,
    `Incluir menús predecesores: ${args.includeMenuPredecessors ? "SI" : "NO"}`,
    `Ejecutable detectado: ${args.executableRow?.EJECUTABLE || "N/A"}`,
    `Menu detectado: ${args.menuChain.length}`,
    `Parametros encontrados: ${args.parameterPairs.length}`,
    `Otros archivos detectados: ${args.otherFiles.length}`,
    `Archivos generados: ${args.files.length}`,
    "",
    fileList
  ].join("\n");
}

export function renderFilePreview(file: GeneratedFile) {
  return `<!doctype html>
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
}

export function renderPairPreview(pair: GeneratedPair) {
  const title = `SQL - ${cleanValue(pair.parametro.PARAMETRO)}`;
  return `<!doctype html>
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
}

export function renderAllFilesPreview(files: GeneratedFile[]) {
  const blocks = files
    .map((file) => `<h2>${escapeHtml(file.path)}</h2><pre>${escapeHtml(file.content)}</pre>`)
    .join("\n");

  return `<!doctype html>
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
}

export function toDownloadFileName(path: string) {
  const normalized = String(path || "").replace(/\\/g, "/");
  const parts = normalized.split("/").filter(Boolean);
  return parts.length ? parts[parts.length - 1] : "archivo.sql";
}

export function buildZipFileName(secEjecutable: string) {
  return `sql_parametros_${secEjecutable}.zip`;
}

export function buildParameterGroupFileName(parametro: unknown) {
  return buildParametroFileName(parametro);
}
