import express from 'express';
import multer from 'multer';
import cors from 'cors';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import { parseTaquillaFile, validateRows, aggregateBySala, calcPelven, buildWorkbook } from './pelven-engine.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app = express();
const PORT = process.env.PORT || 3001;
app.use(cors());
app.use(express.json({ limit: '50mb' }));

const upload = multer({ dest: path.join(__dirname, 'uploads'), limits: { fileSize: 200 * 1024 * 1024 } });
if (!fs.existsSync(path.join(__dirname, 'uploads'))) fs.mkdirSync(path.join(__dirname, 'uploads'));

// Maestro en memoria (seed + ediciones). Clave: complejoNormalizado|sala
let maestro = JSON.parse(fs.readFileSync(path.join(__dirname, 'maestro-seed.json'), 'utf-8'));
const norm = (s) => String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();

// multer entrega el nombre en latin1; se corrige para nombres con tildes/ñ.
const fixFilename = (name) => {
  const raw = String(name ?? '');
  const fixed = Buffer.from(raw, 'latin1').toString('utf8');
  return fixed.includes('\uFFFD') ? raw : fixed;
};

function maestroMap() {
  const m = new Map();
  for (const r of maestro) m.set(norm(r.complejo) + '|' + String(r.sala), r);
  return m;
}

// Salud
app.get('/api/health', (req, res) => res.json({ ok: true }));

// Maestro CRUD
app.get('/api/maestro', (req, res) => res.json(maestro));
app.post('/api/maestro', (req, res) => {
  const r = req.body;
  if (!r.complejo || !r.sala) return res.status(400).json({ error: 'complejo y sala requeridos' });
  const key = norm(r.complejo) + '|' + String(r.sala);
  const map = maestroMap();
  if (map.has(key)) return res.status(409).json({ error: 'ya existe, usa PUT' });
  maestro.push({ nro: maestro.length + 1, circuito: '', administradora: '', estado: '', ciudad: '', aforo: 0, ...r });
  res.json({ ok: true });
});
app.put('/api/maestro/:idx', (req, res) => {
  const i = Number(req.params.idx);
  if (!maestro[i]) return res.status(404).json({ error: 'no existe' });
  maestro[i] = { ...maestro[i], ...req.body };
  res.json(maestro[i]);
});

// Subir uno o varios reportes CNAC (.xls, .xlsx, .ods, .csv, .tsv)
app.post('/api/upload', upload.any(), async (req, res) => {
  try {
    if (!req.files?.length) return res.status(400).json({ error: 'No se recibió ningún archivo' });

    const semanasPorAnio = Number(req.body.semanasPorAnio || 52);
    const factorDiario = Number(req.body.factorDiario || 4.2); // base = promedio diario (Vesp+Inter+Noche) * factorDiario
    let factoresPorComplejo = {};
    if (req.body.factoresPorComplejo) {
      try {
        factoresPorComplejo = JSON.parse(req.body.factoresPorComplejo);
      } catch {
        return res.status(400).json({ error: 'factoresPorComplejo no es JSON válido' });
      }
    }
    if (!Number.isFinite(factorDiario) || factorDiario <= 0) {
      return res.status(400).json({ error: 'factorDiario debe ser un número mayor que cero' });
    }

    const allRows = [];
    const invalidRows = new Set();
    const allErrors = [];
    const allWarnings = [];
    const perFile = [];

    for (const f of req.files) {
      const fileName = fixFilename(f.originalname);
      let parsed;
      try {
        parsed = await parseTaquillaFile(f.path, fileName);
      } catch (e) {
        throw new Error(`${fileName}: ${e.message || e}`);
      } finally {
        fs.unlink(f.path, () => {}); // borrar temporal de multer incluso si el parseo falla
      }

      const rows = parsed.rows.map((r) => ({ ...r, _file: fileName }));
      const { errors, warnings, invalidRows: fileInvalid } = validateRows(rows);
      for (const r of fileInvalid) invalidRows.add(r);
      for (const r of rows) allRows.push(r);
      for (const e of errors) allErrors.push(e);
      for (const w of warnings) allWarnings.push(w);
      perFile.push({
        file: fileName,
        filas: rows.length,
        filasValidas: rows.length - fileInvalid.length,
        filasInvalidas: fileInvalid.length,
        errores: errors.length,
        avisos: warnings.length,
        semanas: parsed.semanas.size,
        rango: parsed.rango,
      });
    }

    const incluirTodosCircuitos = req.body.incluirTodosCircuitos === 'true' || req.body.incluirTodosCircuitos === true;
    const validRows = allRows.filter((r) => !invalidRows.has(r));
    const agg = aggregateBySala(validRows);
    const pelven = calcPelven(agg, maestroMap(), { semanasPorAnio, factorDiario, factoresPorComplejo, incluirTodosCircuitos });
    const circuitosDetectados = Array.from(new Set(pelven.map((r) => r.circuito).filter(Boolean)));

    res.json({
      archivos: perFile,
      totalFilas: allRows.length,
      filasValidas: validRows.length,
      filasInvalidas: invalidRows.size,
      salasDetectadas: Object.keys(agg).length,
      circuitosDetectados,
      bloqueado: allErrors.length > 0,
      errores: allErrors.slice(0, 200),
      totalErrores: allErrors.length,
      avisos: allWarnings.slice(0, 200),
      totalAvisos: allWarnings.length,
      preview: pelven.slice(0, 20),
      _full: pelven
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: String(e.message || e) });
  } finally {
    // Limpieza defensiva: cubre errores tempranos y archivos que no llegaron a parsearse.
    for (const f of req.files || []) fs.unlink(f.path, () => {});
  }
});

// Generar y descargar consolidado
app.post('/api/export', async (req, res) => {
  try {
    const { filas, anio = 2026, bloqueado } = req.body || {};
    if (bloqueado) return res.status(409).json({ error: 'Exportación bloqueada: hay errores de validación sin resolver' });
    const year = Number(anio);
    if (!Number.isInteger(year) || year < 2000 || year > 2100) return res.status(400).json({ error: 'Año inválido' });
    if (!Array.isArray(filas) || !filas.length) return res.status(400).json({ error: 'filas vacías' });
    const buf = await buildWorkbook(filas, year);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=Pelven_aplica_${year}.xlsx`);
    res.send(Buffer.from(buf));
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: String(e.message || e) });
  }
});

// Frontend build (vite) o dev estático
const dist = path.join(__dirname, 'dist');
if (fs.existsSync(dist)) app.use(express.static(dist));
else app.use(express.static(path.join(__dirname, 'public')));

app.listen(PORT, () => console.log(`PELVEN web en http://localhost:${PORT}`));
