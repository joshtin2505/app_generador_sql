"use client";

import * as React from "react";

import { buildPreviewFileHref, buildPreviewParamsHref, buildPreviewTemplateHref, openPreviewRoute } from "./file-preview";
import { buildPreviewReportRecord } from "@/lib/preview-report";
import { getDownloadName, getNonParameterOutputFiles, parseWorkbook, type GeneratedContext, type GeneratedPair, type SortMode, type SqlFile } from "@/lib/sql-generator";
import { useSqlGeneratorStore } from "./sql-generator-store";

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

export function useSqlGenerator() {
  const file = useSqlGeneratorStore((state) => state.file);
  const secEjecutable = useSqlGeneratorStore((state) => state.secEjecutable);
  const roleName = useSqlGeneratorStore((state) => state.roleName);
  const includeMenuPredecessors = useSqlGeneratorStore((state) => state.includeMenuPredecessors);
  const sortBy = useSqlGeneratorStore((state) => state.sortBy);
  const status = useSqlGeneratorStore((state) => state.status);
  const isProcessing = useSqlGeneratorStore((state) => state.isProcessing);
  const workbookName = useSqlGeneratorStore((state) => state.workbookName);
  const generatedContext = useSqlGeneratorStore((state) => state.generatedContext);
  const generatedPairs = useSqlGeneratorStore((state) => state.generatedPairs);
  const generatedFiles = useSqlGeneratorStore((state) => state.generatedFiles);
  const otherFiles = useSqlGeneratorStore((state) => state.otherFiles);
  const activeTemplateKey = useSqlGeneratorStore((state) => state.activeTemplateKey);
  const setFile = useSqlGeneratorStore((state) => state.setFile);
  const setSecEjecutable = useSqlGeneratorStore((state) => state.setSecEjecutable);
  const setRoleName = useSqlGeneratorStore((state) => state.setRoleName);
  const setIncludeMenuPredecessors = useSqlGeneratorStore((state) => state.setIncludeMenuPredecessors);
  const setSortBy = useSqlGeneratorStore((state) => state.setSortBy);
  const startProcessing = useSqlGeneratorStore((state) => state.startProcessing);
  const setProcessSuccess = useSqlGeneratorStore((state) => state.setProcessSuccess);
  const setProcessError = useSqlGeneratorStore((state) => state.setProcessError);
  const savePreviewReport = useSqlGeneratorStore((state) => state.savePreviewReport);
  const selectPreviewReport = useSqlGeneratorStore((state) => state.selectPreviewReport);

  const canExport = generatedFiles.length > 0;

  const persistResult = React.useCallback(
    (result: Awaited<ReturnType<typeof parseWorkbook>>, sortMode: SortMode) => {
      const previewReport = buildPreviewReportRecord({
        generatedContext: result.generatedContext,
        generatedPairs: result.generatedPairs,
        generatedFiles: result.generatedFiles,
        otherFiles: result.otherFiles,
        workbookName: result.workbookName,
        roleName: result.roleName,
        includeMenuPredecessors,
        sortBy: sortMode,
      });

      savePreviewReport(previewReport);
      selectPreviewReport(previewReport.templateKey);

      setProcessSuccess({
        workbookName: result.workbookName,
        status: buildStatus({
          workbookName: result.workbookName,
          secEjecutable: result.secEjecutable,
          roleName: result.roleName,
          includeMenuPredecessors,
          generatedContext: result.generatedContext,
          generatedPairs: result.generatedPairs,
          generatedFiles: result.generatedFiles,
        }),
        generatedContext: result.generatedContext,
        generatedPairs: result.generatedPairs,
        generatedFiles: result.generatedFiles,
        otherFiles: result.otherFiles,
      });
    },
    [includeMenuPredecessors, savePreviewReport, selectPreviewReport, setProcessSuccess],
  );

  const processFile = React.useCallback(async () => {
    if (!file) {
      setProcessError("Debes seleccionar un archivo Excel.");
      return;
    }

    if (!secEjecutable.trim()) {
      setProcessError("Debes indicar SEC_EJECUTABLE.");
      return;
    }

    startProcessing("Leyendo archivo Excel...");

    try {
      const result = await parseWorkbook(
        file,
        secEjecutable.trim(),
        roleName.trim() || defaultRoleName,
        includeMenuPredecessors,
        sortBy,
      );

      persistResult(result, sortBy);
    } catch (error) {
      const message = error instanceof Error ? error.message : "Error inesperado al procesar el archivo.";
      setProcessError(`Error: ${message}`);
    }
  }, [file, secEjecutable, roleName, includeMenuPredecessors, sortBy, startProcessing, setProcessError, persistResult]);

  const handleSortChange = React.useCallback(
    async (value: SortMode) => {
      setSortBy(value);

      if (!file) {
        return;
      }

      try {
        const result = await parseWorkbook(
          file,
          secEjecutable.trim(),
          roleName.trim() || defaultRoleName,
          includeMenuPredecessors,
          value,
        );

        persistResult(result, value);
      } catch (error) {
        const message = error instanceof Error ? error.message : "Error inesperado al reprocesar el orden.";
        setProcessError(`Error: ${message}`);
      }
    },
    [file, secEjecutable, roleName, includeMenuPredecessors, setProcessError, setSortBy, persistResult],
  );

  const downloadZip = React.useCallback(() => {
    void buildZipBlob(generatedFiles).then((blob) => {
      const href = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = href;
      link.download = `sql_parametros_${secEjecutable.trim()}.zip`;
      link.click();
      URL.revokeObjectURL(href);
    });
  }, [generatedFiles, secEjecutable]);

  const downloadAllUncompressed = React.useCallback(() => {
    generatedFiles.forEach((output) => {
      downloadTextFile(output.path, output.content);
    });
  }, [generatedFiles]);

  const openAllInBrowser = React.useCallback(() => {
    if (!activeTemplateKey) {
      setProcessError("Todavía no hay un preview guardado.");
      return;
    }

    if (!openPreviewRoute(buildPreviewTemplateHref(activeTemplateKey))) {
      setProcessError("El navegador bloqueó la ventana emergente. Habilita popups para esta página.");
    }
  }, [activeTemplateKey, setProcessError]);

  const openPairInBrowser = React.useCallback(
    (pair: GeneratedPair) => {
      if (!activeTemplateKey) {
        setProcessError("Todavía no hay un preview guardado.");
        return;
      }

      const pairPath = pair.files[0]?.path || pair.files[1]?.path || "";
      const href = pairPath ? buildPreviewFileHref(activeTemplateKey, pairPath) : buildPreviewParamsHref(activeTemplateKey);

      if (!openPreviewRoute(href)) {
        setProcessError("El navegador bloqueó la ventana emergente. Habilita popups para esta página.");
      }
    },
    [activeTemplateKey, setProcessError],
  );

  const downloadFile = React.useCallback((file: SqlFile) => {
    downloadTextFile(file.path, file.content);
  }, []);

  const downloadPair = React.useCallback((pair: GeneratedPair) => {
    pair.files.forEach((file) => {
      downloadTextFile(file.path, file.content);
    });
  }, []);

  const openFile = React.useCallback(
    (file: SqlFile) => {
      if (!activeTemplateKey) {
        setProcessError("Todavía no hay un preview guardado.");
        return;
      }

      if (!openPreviewRoute(buildPreviewFileHref(activeTemplateKey, file.path))) {
        setProcessError("El navegador bloqueó la ventana emergente. Habilita popups para esta página.");
      }
    },
    [activeTemplateKey, setProcessError],
  );

  return {
    file,
    secEjecutable,
    roleName,
    includeMenuPredecessors,
    sortBy,
    status,
    isProcessing,
    workbookName,
    generatedContext,
    generatedPairs,
    generatedFiles,
    otherFiles,
    canExport,
    setFile,
    setSecEjecutable,
    setRoleName,
    setIncludeMenuPredecessors,
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
