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
app.use(express.json({ limit: '10mb' }));

const upload = multer({ dest: path.join(__dirname, 'uploads'), limits: { fileSize: 200 * 1024 * 1024 } });
if (!fs.existsSync(path.join(__dirname, 'uploads'))) fs.mkdirSync(path.join(__dirname, 'uploads'));

// Maestro en memoria (seed + ediciones). Clave: complejoNormalizado|sala
let maestro = JSON.parse(fs.readFileSync(path.join(__dirname, 'maestro-seed.json'), 'utf-8'));
const norm = (s) => String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();

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
    const semanasPorAnio = Number(req.body.semanasPorAnio || 52);
    const factorDiario = Number(req.body.factorDiario || 4.2); // base = promedio diario (Vesp+Inter+Noche) * factorDiario
    const factoresPorComplejo = req.body.factoresPorComplejo ? JSON.parse(req.body.factoresPorComplejo) : {};
    const allRows = [];
    const perFile = [];
    for (const f of req.files) {
      const parsed = await parseTaquillaFile(f.path, f.originalname);
      fs.unlink(f.path, () => {}); // borrar temporal de multer para no acumular 78MB por carga
      const { errors, warnings } = validateRows(parsed.rows);
      perFile.push({ file: f.originalname, filas: parsed.rows.length, errores: errors.length, avisos: warnings.length, semanas: parsed.semanas.size, rango: parsed.rango });
      for (const r of parsed.rows) allRows.push({ ...r, _file: f.originalname });
    }
    const incluirTodosCircuitos = req.body.incluirTodosCircuitos === 'true' || req.body.incluirTodosCircuitos === true;
    const { errors, warnings } = validateRows(allRows);
    const agg = aggregateBySala(allRows);
    const pelven = calcPelven(agg, maestroMap(), { semanasPorAnio, factorDiario, factoresPorComplejo, incluirTodosCircuitos });
    const circuitosDetectados = Array.from(new Set(pelven.map((r) => r.circuito).filter(Boolean)));
    res.json({
      archivos: perFile,
      totalFilas: allRows.length,
      salasDetectadas: Object.keys(agg).length,
      circuitosDetectados,
      errores: errors.slice(0, 200),
      totalErrores: errors.length,
      avisos: warnings.slice(0, 200),
      totalAvisos: warnings.length,
      preview: pelven.slice(0, 20),
      _full: pelven
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: String(e.message || e) });
  }
});

// Generar y descargar consolidado
app.post('/api/export', express.json({ limit: '50mb' }), async (req, res) => {
  try {
    const { filas, anio = 2026 } = req.body;
    if (!Array.isArray(filas) || !filas.length) return res.status(400).json({ error: 'filas vacías' });
    const buf = await buildWorkbook(filas, anio);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=Pelven_aplica_${anio}.xlsx`);
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
