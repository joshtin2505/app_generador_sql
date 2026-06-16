"use client";

import { Checkbox } from "@/components/ui/checkbox";
import { Input } from "@/components/ui/input";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import type { SortMode } from "@/lib/sql-generator";

export function GeneratorControls(props: {
  file: File | null;
  secEjecutable: string;
  roleName: string;
  includeMenuPredecessors: boolean;
  sortBy: SortMode;
  onFileChange: (file: File | null) => void;
  onSecEjecutableChange: (value: string) => void;
  onRoleNameChange: (value: string) => void;
  onIncludeMenuPredecessorsChange: (value: boolean) => void;
  onSortChange: (value: SortMode) => void;
}) {
  return (
    <section className="grid gap-4 rounded-3xl border border-slate-900/10 bg-white/80 p-5 shadow-[0_10px_35px_rgba(15,57,62,0.12)] backdrop-blur-sm xl:grid-cols-[1.3fr_170px_200px_220px_220px_auto] xl:items-end">
      <div className="grid gap-2">
        <span className="text-sm font-semibold text-slate-700">Archivo Excel (.xlsx)</span>
        <Input type="file" accept=".xlsx,.xlsm,.xls" onChange={(event) => props.onFileChange(event.target.files?.[0] ?? null)} className="bg-white" />
        <span className="text-xs text-slate-500">{props.file ? props.file.name : "Ningún archivo seleccionado"}</span>
      </div>

      <div className="grid gap-2">
        <span className="text-sm font-semibold text-slate-700">SEC_EJECUTABLE</span>
        <Input type="number" value={props.secEjecutable} onChange={(event) => props.onSecEjecutableChange(event.target.value)} className="bg-white" />
      </div>

      <div className="grid gap-2">
        <span className="text-sm font-semibold text-slate-700">Rol privilegios</span>
        <Input type="text" value={props.roleName} onChange={(event) => props.onRoleNameChange(event.target.value)} className="bg-white" />
      </div>

      <div className="grid gap-2 rounded-2xl border border-slate-200 bg-white px-3 py-2">
        <span className="text-sm font-semibold text-slate-700">Opciones de menú</span>
        <span className="flex min-h-10 items-center gap-2 text-sm text-slate-700">
          <Checkbox checked={props.includeMenuPredecessors} onCheckedChange={(checked) => props.onIncludeMenuPredecessorsChange(Boolean(checked))} />
          Incluir menús predecesores
        </span>
      </div>

      <div className="grid gap-2">
        <span className="text-sm font-semibold text-slate-700">Ordenar por</span>
        <Select value={props.sortBy} onValueChange={(value) => props.onSortChange(value as SortMode)}>
          <SelectTrigger className="bg-white">
            <SelectValue placeholder="Selecciona" />
          </SelectTrigger>
          <SelectContent>
            <SelectItem value="ordenamiento">ORDENAMIENTO</SelectItem>
            <SelectItem value="sec_ejecutable_parametro">SEC_EJECUTABLE_PARAMETRO</SelectItem>
            <SelectItem value="sec_parametro">SEC_PARAMETRO</SelectItem>
          </SelectContent>
        </Select>
      </div>

      <p className="text-sm leading-6 text-slate-600 xl:col-span-6">
        El libro debe incluir las hojas <strong>Ejecutable</strong>, <strong>Ejecutable_Parametro</strong>, <strong>Parametro</strong>, <strong>Menu</strong>, <strong>MST_OBJETO</strong> y <strong>Control IDs</strong>.
      </p>
    </section>
  );
}
