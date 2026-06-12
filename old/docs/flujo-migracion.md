# Flujo de Migración

Este proyecto toma un libro de Excel de migración y arma los SQL asociados a un reporte o proceso específico.

## Entrada principal

El disparador del proceso es el `SEC_EJECUTABLE`.

A partir de ese valor la app busca:

- La fila correspondiente en `Ejecutable`.
- Las filas relacionadas en `Ejecutable_Parametro`.
- Los parámetros en `Parametro`.
- La ruta de menú en `Menu`.
- El objeto en `MST_OBJETO`.
- El componente en `Control IDs`.

## Salida generada

La salida se organiza por familia de script:

- `parametro/`: scripts de `SRT_PARAMETRO` y `SRT_EJECUTABLE_PARAMETRO`.
- `ejecutable/`: script del `AAT_EJECUTABLE`.
- `opcion_menu/`: scripts de `MST_MENU`, `MST_OBJETO` y `MSP_PRIVILEGIOS_ROL`.

## Criterios de cruce

- `Ejecutable_Parametro.SEC_EJECUTABLE` debe coincidir con el valor ingresado.
- `Menu` se cruza por el nombre del ejecutable y por el nombre complementario.
- `MST_OBJETO.NOMBRE_COMPLEMENTO` se usa como referencia para el objeto final.
- `Control IDs.ID` se usa para obtener el componente del ejecutable.

## Dependencias del workbook

La hoja `Menu` contiene dos lados:

- Izquierda: estructura de origen.
- Derecha: estructura de migración.

La app usa la parte de migración para producir el SQL final y conserva la parte de origen como contexto.
