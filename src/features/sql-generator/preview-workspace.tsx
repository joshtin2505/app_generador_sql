"use client";

import { useMemo, useState } from "react";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { findFileByPath } from "@/lib/preview-report";
import type { SqlFile } from "@/lib/sql-generator";
import { useSqlGeneratorStore } from "./sql-generator-store";

function formatDate(value: string) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return "N/A";
  }

  return new Intl.DateTimeFormat("es-ES", {
    dateStyle: "medium",
    timeStyle: "short",
  }).format(date);
}

async function copyText(value: string) {
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(value);
    return;
  }

  const textarea = document.createElement("textarea");
  textarea.value = value;
  textarea.style.position = "fixed";
  textarea.style.opacity = "0";
  document.body.appendChild(textarea);
  textarea.focus();
  textarea.select();
  document.execCommand("copy");
  document.body.removeChild(textarea);
}

function CodeBlock(props: { title: string; content: string }) {
  const [copied, setCopied] = useState(false);

  const handleCopy = async () => {
    await copyText(props.content);
    setCopied(true);
    window.setTimeout(() => setCopied(false), 1400);
  };

  return (
    <section className="overflow-hidden rounded-3xl border border-slate-200 bg-slate-950 shadow-[0_12px_40px_rgba(15,23,42,0.16)]">
      <header className="flex items-center justify-between gap-3 border-b border-white/10 bg-slate-900 px-4 py-3">
        <div className="min-w-0">
          <h3 className="truncate text-sm font-semibold text-slate-100">{props.title}</h3>
        </div>
        <Button size="sm" variant="secondary" onClick={handleCopy}>
          {copied ? "Copiado" : "Copiar"}
        </Button>
      </header>
      <pre className="max-h-[78vh] overflow-auto p-4 font-mono text-xs leading-6 text-slate-100 whitespace-pre-wrap">
        {props.content}
      </pre>
    </section>
  );
}

export function PreviewWorkspace() {
  const previewReports = useSqlGeneratorStore((state) => state.previewReports);
  const reports = useMemo(() => Object.values(previewReports).sort((left, right) => right.generatedAt.localeCompare(left.generatedAt)), [previewReports]);
  const template = reports[0] || null;
  const [selectedFilePath, setSelectedFilePath] = useState<string | null>(null);
  const [showAll, setShowAll] = useState(false);

  if (!template) {
    return (
      <main className="min-h-screen bg-[radial-gradient(circle_at_top_left,rgba(218,255,239,0.85),transparent_30%),radial-gradient(circle_at_top_right,rgba(255,231,204,0.8),transparent_28%),linear-gradient(160deg,#f7f5ef,#e8efe6)] px-4 py-8 text-slate-900 sm:px-6 lg:px-8">
        <div className="mx-auto flex w-full max-w-6xl flex-col gap-6">
          <Card className="border-slate-900/10 bg-white/85 shadow-[0_10px_35px_rgba(15,57,62,0.12)] backdrop-blur-sm">
            <CardHeader>
              <CardTitle className="text-2xl">Preview persistente</CardTitle>
              <CardDescription>No hay reportes guardados todavía en el localStorage.</CardDescription>
            </CardHeader>
          </Card>
        </div>
      </main>
    );
  }

  const selectedFile = selectedFilePath ? findFileByPath(template.generatedFiles, selectedFilePath) : null;
  const fileBlocks = template.generatedFiles.map((file: SqlFile) => <CodeBlock key={file.path} title={file.path} content={file.content} />);

  return (
    <main className="min-h-screen bg-[radial-gradient(circle_at_top_left,rgba(218,255,239,0.85),transparent_30%),radial-gradient(circle_at_top_right,rgba(255,231,204,0.8),transparent_28%),linear-gradient(160deg,#f7f5ef,#e8efe6)] px-4 py-6 text-slate-900 sm:px-6 lg:px-8">
      <div className="mx-auto flex w-full max-w-7xl flex-col gap-6">
        <Card className="border-slate-900/10 bg-white/85 shadow-[0_10px_35px_rgba(15,57,62,0.12)] backdrop-blur-sm">
          <CardHeader className="border-b border-slate-200/80 pb-4">
            <CardTitle className="text-2xl">{template.templateName || template.reportData.nombre || template.templateKey}</CardTitle>
            <CardDescription>
              {template.workbookName || "Reporte persistido"} · {template.generatedFiles.length} archivos · {template.generatedPairs.length} parámetros
            </CardDescription>
          </CardHeader>
          <CardContent className="grid gap-3 pt-4 text-sm text-slate-700 md:grid-cols-2 xl:grid-cols-4">
            <div className="rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3">
              <div className="text-xs uppercase tracking-[0.18em] text-slate-500">Ejecutable</div>
              <div className="mt-1 font-medium text-slate-900">{template.reportData.ejecutable || "N/A"}</div>
            </div>
            <div className="rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3">
              <div className="text-xs uppercase tracking-[0.18em] text-slate-500">Rol</div>
              <div className="mt-1 font-medium text-slate-900">{template.reportData.privilegesRole}</div>
            </div>
            <div className="rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3">
              <div className="text-xs uppercase tracking-[0.18em] text-slate-500">Actualizado</div>
              <div className="mt-1 font-medium text-slate-900">{formatDate(template.generatedAt)}</div>
            </div>
            <div className="rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3">
              <div className="text-xs uppercase tracking-[0.18em] text-slate-500">Vista</div>
               <div className="mt-1 font-medium text-slate-900">{showAll ? "Todos" : selectedFile ? "Archivo seleccionado" : "Lista de archivos"}</div>
            </div>
          </CardContent>
        </Card>

        <div className="grid gap-6">
          <Card className="border-slate-900/10 bg-white/85 shadow-[0_10px_35px_rgba(15,57,62,0.12)] backdrop-blur-sm">
            <CardHeader className="border-b border-slate-200/80 pb-4">
              <CardTitle className="text-lg">Contenido</CardTitle>
              <CardDescription>El código se muestra tal cual fue generado y puede copiarse bloque por bloque.</CardDescription>
            </CardHeader>
            <CardContent className="space-y-4 pt-4 text-sm text-slate-700">
              <div className="flex flex-wrap gap-2">
                <Button size="sm" variant={showAll ? "default" : "outline"} onClick={() => { setShowAll(true); setSelectedFilePath(null); }} className={showAll ? "bg-teal-700 text-white hover:bg-teal-800" : ""}>
                  Ver todos
                </Button>
                <Button size="sm" variant={!showAll && !selectedFilePath ? "default" : "outline"} onClick={() => { setShowAll(false); setSelectedFilePath(null); }} className={!showAll && !selectedFilePath ? "bg-teal-700 text-white hover:bg-teal-800" : ""}>
                  Seleccionar archivo
                </Button>
              </div>

              {!showAll && !selectedFilePath && (
              <div className="space-y-2">
                <div className="text-xs font-semibold uppercase tracking-[0.18em] text-slate-500">Archivos disponibles</div>
                <div className="flex flex-wrap gap-2">
                  {template.generatedFiles.map((file: SqlFile) => (
                    <Button key={file.path} variant="outline" size="sm" onClick={() => setSelectedFilePath(file.path)} className="justify-start overflow-hidden text-left">
                      {file.path}
                    </Button>
                  ))}
                </div>
              </div>
              )}
            </CardContent>
          </Card>

          {showAll ? (
            <div className="grid gap-4">
              {fileBlocks.length ? fileBlocks : (
                <Card className="border-slate-900/10 bg-white/85 shadow-[0_10px_35px_rgba(15,57,62,0.12)] backdrop-blur-sm">
                  <CardContent className="p-6 text-center text-sm text-slate-500">No hay archivos guardados para esta plantilla.</CardContent>
                </Card>
              )}
            </div>
          ) : selectedFile ? (
            <CodeBlock title={selectedFile.path} content={selectedFile.content} />
          ) : null}
        </div>
      </div>
    </main>
  );
}
