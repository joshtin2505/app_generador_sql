"use client";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import type { GeneratedFile } from "@/lib/sql-generator";

export function GeneratorFilesTable(props: {
  files: GeneratedFile[];
  onDownload: (file: GeneratedFile) => void;
  onOpen: (file: GeneratedFile) => void;
}) {
  return (
    <Card className="border-slate-900/10 bg-white/80 shadow-[0_10px_35px_rgba(15,57,62,0.12)] backdrop-blur-sm">
      <CardHeader className="border-b border-slate-200/80 pb-4">
        <CardTitle className="text-lg">Insert, MSP y MST generados</CardTitle>
        <CardDescription>Archivos que no pertenecen a parámetros.</CardDescription>
      </CardHeader>
      <CardContent className="pt-4">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>#</TableHead>
              <TableHead>ARCHIVO</TableHead>
              <TableHead>ACCIONES</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            {props.files.length ? (
              props.files.map((item, index) => (
                <TableRow key={item.path}>
                  <TableCell>{index + 1}</TableCell>
                  <TableCell className="max-w-[50ch] truncate">{item.path}</TableCell>
                  <TableCell>
                    <div className="flex flex-wrap gap-2">
                      <Button size="sm" className="bg-teal-700 text-white hover:bg-teal-800" onClick={() => props.onDownload(item)}>
                        Descargar
                      </Button>
                      <Button size="sm" variant="outline" onClick={() => props.onOpen(item)}>
                        Abrir
                      </Button>
                    </div>
                  </TableCell>
                </TableRow>
              ))
            ) : (
              <TableRow>
                <TableCell colSpan={3} className="py-8 text-center text-slate-500">
                  Sin resultados todavía.
                </TableCell>
              </TableRow>
            )}
          </TableBody>
        </Table>
      </CardContent>
    </Card>
  );
}
