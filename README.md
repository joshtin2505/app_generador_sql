# Generador SQL de Migraciones

Aplicación web local para leer un libro de Excel de migración y generar scripts SQL para reportes, parámetros y opciones de menú.

## Qué genera

- `AAT_EJECUTABLE`
- `SRT_PARAMETRO`
- `SRT_EJECUTABLE_PARAMETRO`
- `MST_MENU`
- `MST_OBJETO`
- `MSP_PRIVILEGIOS_ROL`

## Flujo de uso

1. Abre `index.html` en el navegador.
2. Carga el archivo Excel de migración.
3. Indica el `SEC_EJECUTABLE` que quieres migrar.
4. Ajusta el rol de privilegios si aplica.
5. Pulsa `Procesar`.
6. Descarga el ZIP o revisa los scripts uno a uno.

## Estructura de salida

- `parametro/` para los scripts de parámetros.
- `ejecutable/` para el insert o update del ejecutable.
- `opcion_menu/` para `mst_menu`, `mst_objeto` y privilegios.

## Hojas esperadas en el Excel

- `Ejecutable`
- `Ejecutable_Parametro`
- `Parametro`
- `Menu`
- `MST_OBJETO`
- `Control IDs`

## Notas

- La hoja `Menu` se interpreta por bloques: la parte izquierda representa el origen y la derecha la migración.
- El cruce principal se hace por `SEC_EJECUTABLE` y por el nombre complementario del ejecutable.
- Si una hoja o columna esperada falta, la app detiene el proceso con un mensaje claro.
