"use client";

import * as React from "react";

import {
  getDownloadName,
  getNonParameterOutputFiles,
  parseWorkbook,
  type GeneratedContext,
  type GeneratedPair,
  type SortMode,
  type SqlFile,
} from "@/lib/sql-generator";
import { openHtmlPreview, renderFilesPreview, renderPairPreview } from "./file-preview";

const defaultSecEjecutable = "30198";
const defaultRoleName = "ICEBERG_ZK";

function downloadTextFile(path: string, content: string) {
  const blob = new Blob([content], { type: "text/plain;charset=utf-8" });
  const href = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = href;
  link.download = getDownloadName(path);
  link.click();
  URL.revokeObjectURL(href);
}

function buildZipBlob(files: SqlFile[]) {
  return import("jszip").then(({ default: JSZip }) => {
    const zip = new JSZip();
    files.forEach((file) => {
      zip.file(file.path, file.content);
    });
    return zip.generateAsync({ type: "blob" });
  });
}

function buildStatus(args: {
  workbookName: string;
  secEjecutable: string;
  roleName: string;
  includeMenuPredecessors: boolean;
  generatedContext: GeneratedContext | null;
  generatedPairs: GeneratedPair[];
  generatedFiles: SqlFile[];
}) {
  const otherFiles = getNonParameterOutputFiles(args.generatedFiles, args.generatedPairs);
  const fileList = args.generatedFiles.map((file, index) => `${index + 1}. ${file.path}`).join("\n");

  return [
    `Archivo: ${args.workbookName || "N/A"}`,
    `SEC_EJECUTABLE: ${args.secEjecutable}`,
    `Rol privilegios: ${args.roleName}`,
    `Incluir menús predecesores: ${args.includeMenuPredecessors ? "SI" : "NO"}`,
    `Ejecutable detectado: ${args.generatedContext?.executableRow?.EJECUTABLE || "N/A"}`,
    `Menu detectado: ${args.generatedContext?.menuChain?.length || 0}`,
    `Parametros encontrados: ${args.generatedPairs.length}`,
    `Otros archivos detectados: ${otherFiles.length}`,
    `Archivos generados: ${args.generatedFiles.length}`,
    "",
    fileList,
  ].join("\n");
}

type GeneratorState = {
  file: File | null;
  secEjecutable: string;
  roleName: string;
  includeMenuPredecessors: boolean;
  sortBy: SortMode;
  status: string;
  isProcessing: boolean;
  workbookName: string;
  generatedContext: GeneratedContext | null;
  generatedPairs: GeneratedPair[];
  generatedFiles: SqlFile[];
  otherFiles: SqlFile[];
};

type GeneratorAction =
  | { type: "file/changed"; file: File | null }
  | { type: "sec/changed"; value: string }
  | { type: "role/changed"; value: string }
  | { type: "menuPredecessors/changed"; value: boolean }
  | { type: "sort/changed"; value: SortMode }
  | { type: "process/start"; status: string }
  | {
      type: "process/success";
      workbookName: string;
      status: string;
      generatedContext: GeneratedContext;
      generatedPairs: GeneratedPair[];
      generatedFiles: SqlFile[];
      otherFiles: SqlFile[];
    }
  | { type: "process/error"; status: string }
  | { type: "reset/outputs" };

function reducer(state: GeneratorState, action: GeneratorAction): GeneratorState {
  switch (action.type) {
    case "file/changed":
      return { ...state, file: action.file };
    case "sec/changed":
      return { ...state, secEjecutable: action.value };
    case "role/changed":
      return { ...state, roleName: action.value };
    case "menuPredecessors/changed":
      return { ...state, includeMenuPredecessors: action.value };
    case "sort/changed":
      return { ...state, sortBy: action.value };
    case "process/start":
      return { ...state, isProcessing: true, status: action.status };
    case "process/success":
      return {
        ...state,
        isProcessing: false,
        workbookName: action.workbookName,
        generatedContext: action.generatedContext,
        generatedPairs: action.generatedPairs,
        generatedFiles: action.generatedFiles,
        otherFiles: action.otherFiles,
        status: action.status,
      };
    case "process/error":
      return {
        ...state,
        isProcessing: false,
        generatedContext: null,
        generatedPairs: [],
        generatedFiles: [],
        otherFiles: [],
        status: action.status,
      };
    case "reset/outputs":
      return {
        ...state,
        generatedContext: null,
        generatedPairs: [],
        generatedFiles: [],
        otherFiles: [],
      };
    default:
      return state;
  }
}

const initialState: GeneratorState = {
  file: null,
  secEjecutable: defaultSecEjecutable,
  roleName: defaultRoleName,
  includeMenuPredecessors: false,
  sortBy: "ordenamiento",
  status: "Esperando archivo...",
  isProcessing: false,
  workbookName: "",
  generatedContext: null,
  generatedPairs: [],
  generatedFiles: [],
  otherFiles: [],
};

export function useSqlGenerator() {
  const [state, dispatch] = React.useReducer(reducer, initialState);

  const canExport = state.generatedFiles.length > 0;

  const processFile = React.useCallback(async () => {
    if (!state.file) {
      dispatch({ type: "process/error", status: "Debes seleccionar un archivo Excel." });
      return;
    }

    if (!state.secEjecutable.trim()) {
      dispatch({ type: "process/error", status: "Debes indicar SEC_EJECUTABLE." });
      return;
    }

    dispatch({ type: "process/start", status: "Leyendo archivo Excel..." });

    try {
      const result = await parseWorkbook(
        state.file,
        state.secEjecutable.trim(),
        state.roleName.trim() || defaultRoleName,
        state.includeMenuPredecessors,
        state.sortBy,
      );

      const status = buildStatus({
        workbookName: result.workbookName,
        secEjecutable: result.secEjecutable,
        roleName: result.roleName,
        includeMenuPredecessors: state.includeMenuPredecessors,
        generatedContext: result.generatedContext,
        generatedPairs: result.generatedPairs,
        generatedFiles: result.generatedFiles,
      });

      dispatch({
        type: "process/success",
        workbookName: result.workbookName,
        status,
        generatedContext: result.generatedContext,
        generatedPairs: result.generatedPairs,
        generatedFiles: result.generatedFiles,
        otherFiles: result.otherFiles,
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : "Error inesperado al procesar el archivo.";
      dispatch({ type: "process/error", status: `Error: ${message}` });
    }
  }, [state.file, state.includeMenuPredecessors, state.roleName, state.secEjecutable, state.sortBy]);

  const handleSortChange = React.useCallback(
    async (value: SortMode) => {
      dispatch({ type: "sort/changed", value });

      if (!state.file) {
        return;
      }

      try {
        const result = await parseWorkbook(
          state.file,
          state.secEjecutable.trim(),
          state.roleName.trim() || defaultRoleName,
          state.includeMenuPredecessors,
          value,
        );

        const status = buildStatus({
          workbookName: result.workbookName,
          secEjecutable: result.secEjecutable,
          roleName: result.roleName,
          includeMenuPredecessors: state.includeMenuPredecessors,
          generatedContext: result.generatedContext,
          generatedPairs: result.generatedPairs,
          generatedFiles: result.generatedFiles,
        });

        dispatch({
          type: "process/success",
          workbookName: result.workbookName,
          status,
          generatedContext: result.generatedContext,
          generatedPairs: result.generatedPairs,
          generatedFiles: result.generatedFiles,
          otherFiles: result.otherFiles,
        });
      } catch (error) {
        const message = error instanceof Error ? error.message : "Error inesperado al reprocesar el orden.";
        dispatch({ type: "process/error", status: `Error: ${message}` });
      }
    },
    [state.file, state.includeMenuPredecessors, state.roleName, state.secEjecutable],
  );

  const downloadZip = React.useCallback(() => {
    void buildZipBlob(state.generatedFiles).then((blob) => {
      const href = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = href;
      link.download = `sql_parametros_${state.secEjecutable.trim()}.zip`;
      link.click();
      URL.revokeObjectURL(href);
    });
  }, [state.generatedFiles, state.secEjecutable]);

  const downloadAllUncompressed = React.useCallback(() => {
    state.generatedFiles.forEach((output) => {
      downloadTextFile(output.path, output.content);
    });
  }, [state.generatedFiles]);

  const openAllInBrowser = React.useCallback(() => {
    if (!openHtmlPreview("SQL Generados", renderFilesPreview(state.generatedFiles))) {
      dispatch({ type: "process/error", status: "El navegador bloqueó la ventana emergente. Habilita popups para esta página." });
    }
  }, [state.generatedFiles]);

  const openPairInBrowser = React.useCallback((pair: GeneratedPair) => {
    if (!openHtmlPreview(`SQL - ${pair.parametro.PARAMETRO}`, renderPairPreview(pair))) {
      dispatch({ type: "process/error", status: "El navegador bloqueó la ventana emergente. Habilita popups para esta página." });
    }
  }, []);

  const downloadFile = React.useCallback((file: SqlFile) => {
    downloadTextFile(file.path, file.content);
  }, []);

  const downloadPair = React.useCallback((pair: GeneratedPair) => {
    pair.files.forEach((file) => {
      downloadTextFile(file.path, file.content);
    });
  }, []);

  const openFile = React.useCallback((file: SqlFile) => {
    if (!openHtmlPreview(file.path, renderFilesPreview([file]))) {
      dispatch({ type: "process/error", status: "El navegador bloqueó la ventana emergente. Habilita popups para esta página." });
    }
  }, []);

  return {
    file: state.file,
    secEjecutable: state.secEjecutable,
    roleName: state.roleName,
    includeMenuPredecessors: state.includeMenuPredecessors,
    sortBy: state.sortBy,
    status: state.status,
    isProcessing: state.isProcessing,
    workbookName: state.workbookName,
    generatedContext: state.generatedContext,
    generatedPairs: state.generatedPairs,
    generatedFiles: state.generatedFiles,
    otherFiles: state.otherFiles,
    canExport,
    setFile: (file: File | null) => dispatch({ type: "file/changed", file }),
    setSecEjecutable: (value: string) => dispatch({ type: "sec/changed", value }),
    setRoleName: (value: string) => dispatch({ type: "role/changed", value }),
    setIncludeMenuPredecessors: (value: boolean) => dispatch({ type: "menuPredecessors/changed", value }),
    processFile,
    handleSortChange,
    downloadZip,
    downloadAllUncompressed,
    downloadFile,
    downloadPair,
    openAllInBrowser,
    openFile,
    openPairInBrowser,
  };
}
