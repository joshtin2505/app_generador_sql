"use client";

import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import type { GeneratedContext, GeneratedPair, SqlFile } from "@/lib/sql-generator";

export function GeneratorStatus(props: {
  workbookName: string;
  status: string;
  generatedContext: GeneratedContext | null;
  generatedPairs: GeneratedPair[];
  generatedFiles: SqlFile[];
  otherFiles: SqlFile[];
}) {
  return (
    <section className="grid gap-4 lg:grid-cols-[1.1fr_0.9fr]">
      <Card className="border-slate-900/10 bg-white/80 shadow-[0_10px_35px_rgba(15,57,62,0.12)] backdrop-blur-sm">
        <CardHeader className="border-b border-slate-200/80 pb-4">
          <CardTitle className="text-lg">Estado</CardTitle>
          <CardDescription>Resumen de la última corrida del generador.</CardDescription>
        </CardHeader>
        <CardContent className="pt-4">
          <pre className="min-h-45 whitespace-pre-wrap rounded-2xl border border-slate-200 bg-slate-50 p-4 font-mono text-sm leading-6 text-slate-700">
            {props.status}
          </pre>
        </CardContent>
      </Card>

      <Card className="border-slate-900/10 bg-white/80 shadow-[0_10px_35px_rgba(15,57,62,0.12)] backdrop-blur-sm">
        <CardHeader className="border-b border-slate-200/80 pb-4">
          <CardTitle className="text-lg">Resumen</CardTitle>
          <CardDescription>Datos detectados después del procesamiento.</CardDescription>
        </CardHeader>
        <CardContent className="space-y-3 pt-4 text-sm text-slate-700">
          <div className="flex items-center justify-between rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3">
            <span>Archivo</span>
            <span className="font-medium text-slate-900">{props.workbookName || "N/A"}</span>
          </div>
          <div className="flex items-center justify-between rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3">
            <span>Parametros</span>
            <span className="font-medium text-slate-900">{props.generatedPairs.length}</span>
          </div>
          <div className="flex items-center justify-between rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3">
            <span>Otros archivos</span>
            <span className="font-medium text-slate-900">{props.otherFiles.length}</span>
          </div>
          <div className="flex items-center justify-between rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3">
            <span>Estado de exportación</span>
            <span className="font-medium text-slate-900">{props.generatedFiles.length > 0 ? "Listo" : "Pendiente"}</span>
          </div>
          <div className="flex items-center justify-between rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3">
            <span>Menú detectado</span>
            <span className="font-medium text-slate-900">{props.generatedContext?.menuChain.length ?? 0}</span>
          </div>
          <div className="flex items-center justify-between rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3">
            <span>Ejecutable detectado</span>
            <span className="max-w-[18rem] truncate font-medium text-slate-900">{props.generatedContext?.executableRow?.EJECUTABLE || "N/A"}</span>
          </div>
        </CardContent>
      </Card>
    </section>
  );
}
