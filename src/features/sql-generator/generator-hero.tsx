"use client";

import { Download, ExternalLink, FileSpreadsheet, Loader2, Upload } from "lucide-react";

import { Button } from "@/components/ui/button";

export function GeneratorHero(props: {
  isProcessing: boolean;
  canExport: boolean;
  onProcess: () => void;
  onDownloadZip: () => void;
  onDownloadAll: () => void;
  onOpenAll: () => void;
}) {
  return (
    <section className="overflow-hidden rounded-3xl border border-slate-900/10 bg-white/75 p-6 shadow-[0_10px_35px_rgba(15,57,62,0.12)] backdrop-blur-sm sm:p-8">
      <div className="flex flex-col gap-6 lg:flex-row lg:items-end lg:justify-between">
        <div className="max-w-4xl space-y-4">
          <p className="text-xs font-bold uppercase tracking-[0.18em] text-teal-700">SQL Tool</p>
          <div className="space-y-2">
            <h1 className="text-3xl font-semibold tracking-tight text-slate-900 sm:text-5xl">
              Generador de Scripts de Migración desde Excel
            </h1>
            <p className="max-w-4xl text-base leading-7 text-slate-700 sm:text-lg">
              Carga el archivo, filtra por <strong>SEC_EJECUTABLE</strong>, cruza hojas y descarga un ZIP con scripts para <strong>AAT_EJECUTABLE</strong>, <strong>SRT_PARAMETRO</strong>, <strong>SRT_EJECUTABLE_PARAMETRO</strong>, <strong>MST_MENU</strong>, <strong>MST_OBJETO</strong> y <strong>MSP_PRIVILEGIOS_ROL</strong>.
            </p>
          </div>
        </div>

        <div className="flex flex-wrap gap-2">
          <Button onClick={props.onProcess} disabled={props.isProcessing} className="bg-teal-700 text-white hover:bg-teal-800">
            {props.isProcessing ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Upload className="mr-2 h-4 w-4" />}
            Procesar
          </Button>
          <Button onClick={props.onDownloadZip} disabled={!props.canExport} className="bg-amber-600 text-white hover:bg-amber-700">
            <Download className="mr-2 h-4 w-4" />
            Descargar ZIP
          </Button>
          <Button variant="secondary" onClick={props.onDownloadAll} disabled={!props.canExport}>
            <FileSpreadsheet className="mr-2 h-4 w-4" />
            Descargar todo sin comprimir
          </Button>
          <Button variant="outline" type="button" onClick={props.onOpenAll} disabled={!props.canExport}>
            <ExternalLink className="mr-2 h-4 w-4" />
            Abrir todo en navegador
          </Button>
        </div>
      </div>
    </section>
  );
}
