# PELVEN · Generador de consolidado (Node + React)

Genera el consolidado **"Pelven aplica AAAA"** a partir de los reportes semanales/diarios
de taquilla que los exhibidores envían al CNAC (formato `ESPECIFICACIONES.docx`).

## Hallazgos verificados (2025 → 2026)

- **Input**: el `.xls` de Cines Unidos es en realidad un **TSV `latin-1` con decimales con coma**
  (551.094 filas, 20 complejos, 164 salas). El parser detecta automáticamente TSV vs `.xlsx` real.
- **Validación**: `Total Recaudación = Total Precio × Cantidad` (estricto).
  `Total Precio = Boleto + Servicio + Otros×(1+IVA_otros)` — Boleto/Servicio ya incluyen IVA,
  Otros Cargos suma su IVA (verificado: 78,49+78,49+11,22×1,16 = 169,99). Con el fix: **0 errores**.
- **Fórmula base** (el `.xls` original trae solo valores; deducida por regresión):
  `base(sala) = (espectadores en funciones Vespertina+Intermedia+Noche / días con alguna de esas funciones) × 4,2`.
  Matinee, Media Noche y `Sin Función` no cuentan. Réplica sobre las 164 salas CU: error mediano 0,19 %,
  máximo 1,5 %; 12 salas exactas. No requiere excepciones por complejo (Metrocenter incluido).
- **Columnas 20→1 exactas**: `base × n/21`.
- El Pelven nacional tiene 323 salas (todos los circuitos); el maestro
  (`maestro-seed.json`, extraído del Pelven 2026) aporta Circuito/Administradora/Estado/Ciudad/Aforo
  y se mantiene manualmente en la UI.

## Uso

```cmd
cd pelven-web
npm.cmd install
npm.cmd run server   :: API en http://localhost:3001
npm.cmd run dev      :: UI React en http://localhost:5173 (proxy /api)
```

1. Abre la UI, sube uno o varios reportes (`*_aaaammdd-aaaammdd.xls`), ajusta año y factor.
2. Revisa preview, errores y salas sin maestro; completa el maestro si falta alguna.
3. Descarga `Pelven_aplica_AAAA.xlsx` (mismo layout: cabecera Providencia 002, 21→1 + %).

## Notas

- Si ves `EADDRINUSE :::3001`, ya hay un servidor corriendo: cierra la otra ventana
  o libera el puerto antes de iniciar otro (`Stop-Process -Name node`).
- Los temporales de carga se borran solos tras procesar; no se acumulan en `uploads/`.

## API

- `GET /api/health`, `GET /api/maestro`, `POST /api/maestro`, `PUT /api/maestro/:idx`
- `POST /api/upload` (multipart `files`, campos `factorDiario`, `factoresPorComplejo` JSON, `semanasPorAnio`)
- `POST /api/export` (`{ filas, anio }` → xlsx)
