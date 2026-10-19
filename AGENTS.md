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
npm.cmd run server   :: API on http://localhost:3001
npm.cmd run dev      :: Vite UI on http://localhost:5173, proxies /api -> 3001
npm.cmd run build    :: production bundle into pelven-web/dist/
```

- There is **no test, lint, or typecheck script**. Verify changes by running the server and exercising the upload/export flow, or by comparing output against the reference xlsx files at the repo root.
- `server.js` serves `dist/` if it exists (stale builds shadow the source). After UI changes, either use `npm.cmd run dev` or rebuild with `npm.cmd run build`.
- `EADDRINUSE :::3001` means a server is already running; kill node processes before restarting.

## Domain facts (verified; do not "simplify" these away)

Full derivation is in `pelven-web/README.md`; the high-signal constraints:

- Input `.xls` files from CNAC are often **TSV encoded latin-1 with comma decimals**, not real xls. `parseTaquillaFile` sniffs the format (zip magic bytes vs delimited text) and auto-detects the delimiter. Numbers use `.` as thousands separator and `,` as decimal separator.
- Base formula (reverse-engineered, not documented upstream): `base(sala) = espectadores(Vespertina+Intermedia+Noche) / días con esas funciones × 4.2` (4.2 = 21/5, configurable as `factorDiario`). Matinee, Media Noche, Continuada and Sin Función **do not count**.
- Output columns 20..1 are exactly `base × n/21`.
- Validation rule: `Total Recaudación = Total Precio × Cantidad`, where `Total Precio = Boleto + Servicio + Otros×(1+IVA_otros)` (Boleto/Servicio already include IVA).
- Text matching across files is accent- and case-insensitive (`norm()` helper, used in both engine and server for maestro keys `complejo|sala`).

## Operational gotchas

- `maestro-seed.json` is loaded **once at server startup into memory**; UI edits via `/api/maestro` are not persisted back to disk. To make a maestro change permanent, edit the JSON file and restart the server.
- Uploads go to `pelven-web/uploads/` via multer but are deleted immediately after parsing. The folder itself is recreated on boot.
- API responses embed the full result set as `_full` (the UI re-sends it to `/api/export`); don't "optimize" it away or the export breaks.
- Code, UI strings, and data are in **Spanish**; domain terms (complejo, sala, taquilla, recaudación, maestro) map to the CNAC/Pelven domain and should stay consistent.
