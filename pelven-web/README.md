# PELVEN · Generador de consolidado (Node + React)

Genera el consolidado **"Pelven aplica AAAA"** a partir de los reportes semanales/diarios
de taquilla que los exhibidores envían al CNAC (formato `ESPECIFICACIONES.docx`).

## Hallazgos verificados (2025 → 2026)

- **Input**: el `.xls` de Cines Unidos es en realidad un **TSV `latin-1` con decimales con coma**
  (551.094 filas, 20 complejos, 164 salas). El parser detecta automáticamente TSV vs `.xlsx` real.
- **Validación**: `Total Recaudación = Total Precio × Cantidad` (estricto).
  `Total Precio = Boleto + Servicio + Otros×(1+IVA_otros)` — Boleto/Servicio ya incluyen IVA,
  Otros Cargos suma su IVA (verificado: 78,49+78,49+11,22×1,16 = 169,99). Con el fix: **0 errores**
  en el archivo real, **0 filas inválidas**. Las filas con error se excluyen del cálculo y bloquean la exportación.
- **Fórmula base** (el `.xls` original trae solo valores; deducida por regresión):
  `base(sala) = (espectadores en funciones Vespertina+Intermedia+Noche / días con alguna de esas funciones) × 4,2`.
  Matinee, Media Noche y `Sin Función` no cuentan. Réplica sobre las 164 salas CU: error medio 0,26 %,
  máximo 1,54 %. El Excel generado coincide celda a celda con `Pelven_generado_2026.xlsx`.
  No requiere excepciones por complejo (Metrocenter incluido).
- **Columnas 20→1 exactas**: `base × n/21`.
- El Pelven nacional tiene 323 salas (todos los circuitos); el maestro
  (`maestro-seed.json`, extraído del Pelven 2026) aporta Circuito/Administradora/Estado/Ciudad/Aforo.
  Se mantiene editando `maestro-seed.json` y reiniciando el servidor; la pestaña Maestro de la UI es de solo lectura.

## Uso

```cmd
cd pelven-web
npm.cmd install
npm.cmd run server   :: API en http://localhost:3001
npm.cmd run dev      :: UI React en http://localhost:5173 (proxy /api)
```

1. Abre la UI, sube uno o varios reportes (`*_aaaammdd-aaaammdd.xls`), ajusta año y factor.
2. Revisa el resumen por archivo, los avisos y las salas sin maestro. Si hay errores, la exportación
   queda bloqueada y las filas inválidas se excluyen del cálculo: corrige el archivo fuente y vuelve a subirlo.
3. Descarga `Pelven_aplica_AAAA.xlsx` (mismo layout: cabecera Providencia 002, 21→1 + %).

## Pruebas

```cmd
npm.cmd test          :: regresiones del motor (parser, números, validación, cálculo, export)
npm.cmd run verify:real   :: verifica el reporte real de Cines Unidos (requiere el .xls de 78 MB en la raíz)
```

`npm test` no depende del archivo real. `verify:real` valida 551.094 filas, exige 0 errores y compara
la base calculada contra `base2026` del maestro.

## Notas

- Si ves `EADDRINUSE :::3001`, ya hay un servidor corriendo: cierra la otra ventana
  o libera el puerto antes de iniciar otro (`netstat -ano | findstr :3001` y `Stop-Process -Id <PID>`).
- Los temporales de carga se borran solos tras procesar (incluso si el parseo falla); no se acumulan en `uploads/`.
- La exportación responde `409` si se envía `bloqueado: true`; la UI ya deshabilita la descarga con errores.

## API

- `GET /api/health`, `GET /api/maestro`, `POST /api/maestro`, `PUT /api/maestro/:idx`
- `POST /api/upload` (multipart `files`, campos `factorDiario`, `factoresPorComplejo` JSON, `semanasPorAnio`)
  → responde `{ archivos, filasValidas, filasInvalidas, bloqueado, errores, avisos, _full, ... }`
- `POST /api/export` (`{ filas, anio, bloqueado }` → xlsx; `409` si `bloqueado`)
