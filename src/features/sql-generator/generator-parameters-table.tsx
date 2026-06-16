"use client";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import type { GeneratedPair } from "@/lib/sql-generator";

export function GeneratorParametersTable(props: {
  pairs: GeneratedPair[];
  onDownloadPair: (pair: GeneratedPair) => void;
  onOpenPair: (pair: GeneratedPair) => void;
}) {
  return (
    <Card className="border-slate-900/10 bg-white/80 shadow-[0_10px_35px_rgba(15,57,62,0.12)] backdrop-blur-sm">
      <CardHeader className="border-b border-slate-200/80 pb-4">
        <CardTitle className="text-lg">Parametros detectados</CardTitle>
        <CardDescription>Filas asociadas a cada parámetro del ejecutable.</CardDescription>
      </CardHeader>
      <CardContent className="pt-4">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>#</TableHead>
              <TableHead>SEC_EJECUTABLE_PARAMETRO</TableHead>
              <TableHead>SEC_PARAMETRO</TableHead>
              <TableHead>PARAMETRO</TableHead>
              <TableHead>NOMBRE</TableHead>
              <TableHead>TIPO_DATO</TableHead>
              <TableHead>ORDEN</TableHead>
              <TableHead>REQUERIDO</TableHead>
              <TableHead>ACCIONES</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            {props.pairs.length ? (
              props.pairs.map((item, index) => (
                <TableRow key={`${item.ejecutable.SEC_EJECUTABLE_PARAMETRO}-${item.parametro.SEC_PARAMETRO}`}>
                  <TableCell>{index + 1}</TableCell>
                  <TableCell>{item.ejecutable.SEC_EJECUTABLE_PARAMETRO}</TableCell>
                  <TableCell>{item.parametro.SEC_PARAMETRO}</TableCell>
                  <TableCell>{item.parametro.PARAMETRO}</TableCell>
                  <TableCell>{item.parametro.NOMBRE}</TableCell>
                  <TableCell>{item.parametro.TIPO_DATO}</TableCell>
                  <TableCell>{item.ejecutable.ORDENAMIENTO}</TableCell>
                  <TableCell>{item.ejecutable.REQUERIDO}</TableCell>
                  <TableCell>
                    <div className="flex flex-wrap gap-2">
                      <Button size="sm" className="bg-teal-700 text-white hover:bg-teal-800" onClick={() => props.onDownloadPair(item)}>
                        Descargar uno por uno
                      </Button>
                      <Button size="sm" variant="outline" onClick={() => props.onOpenPair(item)}>
                        Abrir
                      </Button>
                    </div>
                  </TableCell>
                </TableRow>
              ))
            ) : (
              <TableRow>
                <TableCell colSpan={9} className="py-8 text-center text-slate-500">
                  Sin parámetros procesados todavía.
                </TableCell>
              </TableRow>
            )}
          </TableBody>
        </Table>
      </CardContent>
    </Card>
  );
}
