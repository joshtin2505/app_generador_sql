"use client";

import { create } from "zustand";
import { createJSONStorage, persist } from "zustand/middleware";

import type { GeneratedContext, GeneratedPair, SortMode, SqlFile } from "@/lib/sql-generator";
import type { PreviewReportRecord } from "@/types";

type SqlGeneratorStore = {
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
  activeTemplateKey: string | null;
  previewReports: Record<string, PreviewReportRecord>;
  setFile: (file: File | null) => void;
  setSecEjecutable: (value: string) => void;
  setRoleName: (value: string) => void;
  setIncludeMenuPredecessors: (value: boolean) => void;
  setSortBy: (value: SortMode) => void;
  startProcessing: (status: string) => void;
  setProcessSuccess: (args: {
    workbookName: string;
    status: string;
    generatedContext: GeneratedContext;
    generatedPairs: GeneratedPair[];
    generatedFiles: SqlFile[];
    otherFiles: SqlFile[];
  }) => void;
  setProcessError: (status: string) => void;
  savePreviewReport: (report: PreviewReportRecord) => void;
  selectPreviewReport: (templateKey: string | null) => void;
  clearPreviewReports: () => void;
};

const initialState = {
  file: null,
  secEjecutable: "30198",
  roleName: "ICEBERG_ZK",
  includeMenuPredecessors: false,
  sortBy: "ordenamiento" as SortMode,
  status: "Esperando archivo...",
  isProcessing: false,
  workbookName: "",
  generatedContext: null,
  generatedPairs: [],
  generatedFiles: [],
  otherFiles: [],
  activeTemplateKey: null,
  previewReports: {},
};

export const useSqlGeneratorStore = create<SqlGeneratorStore>()(
  persist(
    (set) => ({
      ...initialState,
      setFile: (file) => set({ file }),
      setSecEjecutable: (value) => set({ secEjecutable: value }),
      setRoleName: (value) => set({ roleName: value }),
      setIncludeMenuPredecessors: (value) => set({ includeMenuPredecessors: value }),
      setSortBy: (value) => set({ sortBy: value }),
      startProcessing: (status) => set({ isProcessing: true, status }),
      setProcessSuccess: (args) =>
        set({
          isProcessing: false,
          workbookName: args.workbookName,
          generatedContext: args.generatedContext,
          generatedPairs: args.generatedPairs,
          generatedFiles: args.generatedFiles,
          otherFiles: args.otherFiles,
          status: args.status,
        }),
      setProcessError: (status) =>
        set({
          isProcessing: false,
          generatedContext: null,
          generatedPairs: [],
          generatedFiles: [],
          otherFiles: [],
          status,
        }),
      savePreviewReport: (report) =>
        set((state) => ({
          activeTemplateKey: report.templateKey,
          previewReports: {
            ...state.previewReports,
            [report.templateKey]: report,
          },
        })),
      selectPreviewReport: (templateKey) => set({ activeTemplateKey: templateKey }),
      clearPreviewReports: () => set({ previewReports: {}, activeTemplateKey: null }),
    }),
    {
      name: "app_generador_sql.preview",
      storage: createJSONStorage(() => localStorage),
      partialize: (state) => ({
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
        activeTemplateKey: state.activeTemplateKey,
        previewReports: state.previewReports,
      }),
    },
  ),
);
