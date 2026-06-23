import type { GeneratedContext, GeneratedFile, GeneratedPair, MenuRow, SortMode } from "@/lib/sql-generator";
import type { MenuData, ParamData, PreviewReportRecord, ReportData } from "@/types";

function cleanValue(value: unknown) {
  return String(value ?? "").trim();
}

function toNumber(value: unknown) {
  const parsed = Number(cleanValue(value));
  return Number.isFinite(parsed) ? parsed : 0;
}

function toBooleanFlag(value: unknown) {
  const normalized = cleanValue(value).toUpperCase();
  return normalized === "S" || normalized === "SI" || normalized === "Y" || normalized === "1" || normalized === "TRUE";
}

function normalizeParamType(value: unknown): ParamData["type"] {
  const normalized = cleanValue(value).toUpperCase();
  if (normalized === "A" || normalized === "C" || normalized === "L" || normalized === "F") {
    return normalized;
  }

  return "A";
}

function slugify(value: unknown) {
  return cleanValue(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function buildMenuNode(menuRow: MenuRow): MenuData {
  return {
    menu: toNumber(menuRow.target.menu),
    predecessor: toNumber(menuRow.target.predecessor || menuRow.source.predecessor),
    order: toNumber(menuRow.target.order || menuRow.source.order),
    icon: cleanValue(menuRow.target.icon || menuRow.source.icon),
    description: cleanValue(menuRow.target.description || menuRow.source.description || menuRow.target.object || menuRow.target.menu),
    type: cleanValue(menuRow.target.type || menuRow.source.type),
    object: cleanValue(menuRow.target.object || menuRow.source.object),
  };
}

function buildMenuTree(menuChain: MenuRow[]) {
  return menuChain.reduceRight<MenuData | null>((accumulator, menuRow) => {
    const node = buildMenuNode(menuRow);

    if (accumulator) {
      node.predecessors = [accumulator];
    }

    return node;
  }, null) ?? {
    menu: 0,
    predecessor: 0,
    order: 0,
    icon: "",
    description: "",
    type: "",
    object: "",
  };
}

function buildParams(generatedPairs: GeneratedPair[]): ParamData[] {
  return generatedPairs.map((pair) => ({
    secEjecutableParam: toNumber(pair.ejecutable.SEC_EJECUTABLE_PARAMETRO),
    secParam: toNumber(pair.parametro.SEC_PARAMETRO),
    required: toBooleanFlag(pair.ejecutable.REQUERIDO),
    assumedValue: cleanValue(pair.ejecutable.VALOR_ASUMIDO),
    multipleSelection: toBooleanFlag(pair.ejecutable.SELECCION_MULTIPLE),
    order: toNumber(pair.ejecutable.ORDENAMIENTO),
    name: cleanValue(pair.parametro.NOMBRE),
    param: cleanValue(pair.parametro.PARAMETRO),
    state: cleanValue(pair.parametro.ESTADO),
    type: normalizeParamType(pair.parametro.TIPO_DATO),
    data: cleanValue(pair.parametro.DATOS),
  }));
}

export function getTemplateKeyFromReport(reportData: ReportData) {
  const raw = cleanValue(reportData.nombre || reportData.ejecutable || reportData.component || `sec-${reportData.secEjecutable}`);
  return slugify(raw) || `sec-${reportData.secEjecutable}`;
}

export function buildReportData(args: {
  generatedContext: GeneratedContext;
  generatedPairs: GeneratedPair[];
  roleName: string;
}): ReportData {
  const executableRow = args.generatedContext.executableRow;
  const componentRow = args.generatedContext.componentRow;

  return {
    component: cleanValue(componentRow?.nombre || componentRow?.componente || executableRow.SEC_COMPONENTE),
    secEjecutable: toNumber(args.generatedContext.secEjecutable),
    ejecutable: cleanValue(executableRow.EJECUTABLE),
    secComponent: toNumber(executableRow.SEC_COMPONENTE || componentRow?.id),
    nombre: cleanValue(executableRow.NOMBRE || executableRow.EJECUTABLE),
    atribute: toNumber(executableRow.ATRIBUTO),
    description: cleanValue(executableRow.DESCRIPCION || executableRow.NOMBRE || executableRow.EJECUTABLE),
    privilegesRole: cleanValue(args.roleName),
    menu: buildMenuTree(args.generatedContext.menuChain),
    params: buildParams(args.generatedPairs),
  };
}

export function buildPreviewReportRecord(args: {
  generatedContext: GeneratedContext;
  generatedPairs: GeneratedPair[];
  generatedFiles: GeneratedFile[];
  otherFiles: GeneratedFile[];
  workbookName: string;
  roleName: string;
  includeMenuPredecessors: boolean;
  sortBy: SortMode;
}): PreviewReportRecord {
  const reportData = buildReportData({
    generatedContext: args.generatedContext,
    generatedPairs: args.generatedPairs,
    roleName: args.roleName,
  });

  const templateKey = getTemplateKeyFromReport(reportData);

  return {
    templateKey,
    templateName: reportData.ejecutable || reportData.nombre || args.workbookName,
    workbookName: args.workbookName,
    generatedAt: new Date().toISOString(),
    includeMenuPredecessors: args.includeMenuPredecessors,
    sortBy: args.sortBy,
    reportData,
    generatedContext: args.generatedContext,
    generatedPairs: args.generatedPairs,
    generatedFiles: args.generatedFiles,
    otherFiles: args.otherFiles,
  };
}

export function findFileByPath(files: GeneratedFile[], filePath: string) {
  return files.find((file) => file.path === filePath) || null;
}
