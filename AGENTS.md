# AGENTS.md

## Repo layout

- `pelven-web/` — the actual app: Node/Express API (`server.js`) + React/Vite UI (`src/`). All dev happens here.
- `pelven-web/pelven-engine.js` — the core domain logic (parsing, validation, aggregation, PELVEN formula, xlsx export). Server imports it directly; the UI is a thin shell over the API.
- `scratch/` — throwaway exploratory scripts from reverse-engineering the PELVEN formula. Not part of the app; ignore unless reproducing the original analysis.
- Root `package.json` + root `node_modules/` — only hold `xlsx` for the scratch scripts. Do not add app dependencies at the root.
- `*.xls`, `*.xlsx`, `ESPECIFICACIONES.docx` at the root — reference input/output data (CNAC taquilla reports and target Pelven consolidados). Keep them in place; they are the ground truth for format questions.

## Commands (run inside `pelven-web/` on Windows; use `npm.cmd`)

```cmd
cd pelven-web
npm.cmd run server      :: API on http://localhost:3001
npm.cmd run dev         :: Vite UI on http://localhost:5173, proxies /api -> 3001
npm.cmd run build       :: production bundle into pelven-web/dist/
npm.cmd test            :: regresiones del motor (node --test)
npm.cmd run verify:real :: verifica el .xls real de 78 MB en la raíz (0 errores, 551.094 filas)
```

- There is **no lint or typecheck script**. `npm test` covers parser, números, validación, cálculo y export; verifica el flujo completo levantando el servidor y subiendo un archivo.
- `server.js` serves `dist/` if it exists (stale builds shadow the source). After UI changes, either use `npm.cmd run dev` or rebuild with `npm.cmd run build`.
- `EADDRINUSE :::3001` means a server is already running. Find the PID (`netstat -ano | findstr :3001`) and stop only that process; do not kill every `node` process.

## Domain facts (verified; do not "simplify" these away)

Full derivation is in `pelven-web/README.md`; the high-signal constraints:

- Input `.xls` files from CNAC are often **TSV encoded latin-1 with comma decimals**, not real xls. `parseTaquillaFile` sniffs the format (zip magic bytes vs delimited text) and auto-detects the delimiter. Numbers use `.` as thousands separator and `,` as decimal separator.
- In delimited text, only the double quote (`"`) starts a quoted field, and only at the beginning of the field. Apostrophes (`'`) are literal text; treating them as quotes corrupts movie titles like `HOPE ON THE STAGE' in JAPAN`. `parseNumber` accepts both `1.234,56` (es-VE) and `1,234.56` (en-US), and native Excel numbers.
- Base formula (reverse-engineered, not documented upstream): `base(sala) = espectadores(Vespertina+Intermedia+Noche) / días con esas funciones × 4.2` (4.2 = 21/5, configurable as `factorDiario`). Matinee, Media Noche, Continuada and Sin Función **do not count**.
- Output columns 20..1 are exactly `base × n/21`.
- Validation rule: `Total Recaudación = Total Precio × Cantidad`, where `Total Precio = Boleto + Servicio + Otros×(1+IVA_otros)` (Boleto/Servicio already include IVA).
- Text matching across files is accent- and case-insensitive (`norm()` helper, used in both engine and server for maestro keys `complejo|sala`).

## Version control

- Remote: `https://github.com/Cristhofer1312/Pelven.git`, branch `main`. Baseline tag: `baseline-2026`.
- The 78 MB `COMPAÑIA ANONIMA EMPRESA CINES UNIDOS 20250102-20251231.xls` is intentionally **not tracked**. Keep it external (or as a release asset); never commit it.
- Ignored: `node_modules/`, `dist/`, `uploads/`, `scratch/`, `*.pkl`, root `*.xls`/`*.xlsx` except the two small reference files.
- Commit prefixes: `fix:`, `feat:`, `chore:`, `docs:`. Work on `fix/...` or `feat/...` branches; tag releases as `vX.Y.Z`.

## Operational gotchas

- `maestro-seed.json` is loaded **once at server startup into memory**; UI edits via `/api/maestro` are not persisted back to disk. To make a maestro change permanent, edit the JSON file and restart the server. The Maestro tab is currently read-only.
- Uploads go to `pelven-web/uploads/` via multer but are deleted immediately after parsing (in a `finally`, so parse failures are cleaned too). The folder itself is recreated on boot.
- Validation is blocking: rows with errors are excluded from `aggregateBySala`, the response sets `bloqueado: true`, and the UI/export endpoint refuse to generate the xlsx (`409`). Do not re-enable export with errors without an explicit override decision.
- API responses embed the full result set as `_full` (the UI re-sends it to `/api/export`); don't "optimize" it away or the export breaks. `_full` contains only valid rows.
- multer's `originalname` arrives latin1-decoded; `fixFilename()` in `server.js` converts it so accented filenames display correctly. Keep using it for per-file labels and error context.
- Code, UI strings, and data are in **Spanish**; domain terms (complejo, sala, taquilla, recaudación, maestro) map to the CNAC/Pelven domain and should stay consistent.
