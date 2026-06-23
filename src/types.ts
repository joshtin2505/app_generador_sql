export interface ReportData {
  component: string;
  secEjecutable: number;
  ejecutable: string;
  secComponent: number;
  nombre: string;
  atribute: number;
  description: string;
  privilegesRole: string;
  menu: MenuData;
  params: ParamData[];
}

export type MenuData = {
  menu: number;
  predecessor: number;
  order: number;
  icon: string;
  description: string;
  type: string;
  object: string;
  predecessors?: MenuData[];
};

export type ParamData = {
  secEjecutableParam: number;
  secParam: number;
  required: boolean;
  assumedValue: string;
  multipleSelection: boolean;
  order: number;
  name: string;
  param: string;
  state: string;
  type: "A" | "C" | "L" | "F";
  data: string;
};

export type PreviewFileData = {
  path: string;
  content: string;
};

export interface PreviewReportRecord {
  templateKey: string;
  templateName: string;
  workbookName: string;
  generatedAt: string;
  includeMenuPredecessors: boolean;
  sortBy: import("@/lib/sql-generator").SortMode;
  reportData: ReportData;
  generatedContext: import("@/lib/sql-generator").GeneratedContext | null;
  generatedPairs: import("@/lib/sql-generator").GeneratedPair[];
  generatedFiles: import("@/lib/sql-generator").SqlFile[];
  otherFiles: import("@/lib/sql-generator").SqlFile[];
}
