import express from 'express';
import multer from 'multer';
import cors from 'cors';
import path from 'path';
import fs from 'fs';
import crypto from 'crypto';
import { fileURLToPath } from 'url';
import {
  parseTaquillaFile,
  validateRows,
  aggregateBySala,
  calcPelven,
  buildWorkbook,
  findDuplicateRows,
} from './pelven-engine.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app = express();
const PORT = process.env.PORT || 3001;
app.use(cors());
app.use(express.json({ limit: '50mb' }));

const uploadsDir = path.join(__dirname, 'uploads');
const runsDir = path.join(__dirname, 'runs');
const maestroFile = path.join(__dirname, 'maestro.json');
const seedFile = path.join(__dirname, 'maestro-seed.json');
fs.mkdirSync(uploadsDir, { recursive: true });
fs.mkdirSync(runsDir, { recursive: true });

let APP_VERSION = '1.0.0';
try {
  APP_VERSION = JSON.parse(fs.readFileSync(path.join(__dirname, 'package.json'), 'utf8')).version || APP_VERSION;
} catch {
  // El manifiesto no debe impedir que arranque la aplicación.
}

const upload = multer({ dest: uploadsDir, limits: { fileSize: 200 * 1024 * 1024 } });

const norm = (s) => String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
const claveMaestro = (r) => norm(r.complejo) + '|' + String(r.sala ?? '').trim();

// multer entrega el nombre en latin1; se corrige para nombres con tildes/ñ.
const fixFilename = (name) => {
  const raw = String(name ?? '');
  const fixed = Buffer.from(raw, 'latin1').toString('utf8');
  return fixed.includes('\uFFFD') ? raw : fixed;
};

function escribirJsonAtomico(file, data) {
  const tmp = `${file}.${process.pid}.${Date.now()}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(data, null, 2), 'utf8');
  try {
    fs.renameSync(tmp, file);
  } catch {
    // En Windows renombrar sobre un archivo existente puede fallar; se conserva
    // la escritura completa en el destino y se elimina el temporal.
    fs.copyFileSync(tmp, file);
    fs.unlinkSync(tmp);
  }
}

// ---------------------------------------------------------------------------
// Maestro CNAC persistente
// ---------------------------------------------------------------------------
// maestro-seed.json es la semilla versionada. En el primer arranque se copia a
// maestro.json (ignorado por Git), que pasa a ser la fuente de verdad editable.
function cargarMaestro() {
  const seed = JSON.parse(fs.readFileSync(seedFile, 'utf8'));
  if (!fs.existsSync(maestroFile)) {
    escribirJsonAtomico(maestroFile, seed);
    return seed;
  }
  try {
    const data = JSON.parse(fs.readFileSync(maestroFile, 'utf8'));
    if (!Array.isArray(data)) throw new Error('el maestro no es un arreglo');
    return data;
  } catch (e) {
    console.error(`No se pudo leer maestro.json (${e.message}); se usa maestro-seed.json`);
    return seed;
  }
}

let maestro = cargarMaestro();

function guardarMaestro(data = maestro) {
  escribirJsonAtomico(maestroFile, data);
}

function maestroMap() {
  const m = new Map();
  for (const r of maestro) m.set(claveMaestro(r), r);
  return m;
}

function sanitizeMaestro(input = {}, base = {}) {
  const texto = (v, fallback = '') => String(v ?? fallback ?? '').trim();
  const nroRaw = input.nro !== undefined && input.nro !== '' ? input.nro : base.nro;
  const nro = Number(nroRaw);
  const aforo = Number(input.aforo !== undefined && input.aforo !== '' ? input.aforo : base.aforo);
  const baseRaw = input.base2026 !== undefined && input.base2026 !== '' ? input.base2026 : base.base2026;
  const base2026 = baseRaw == null || baseRaw === '' ? null : Number(baseRaw);

  return {
    ...base,
    nro: Number.isFinite(nro) ? Math.trunc(nro) : 0,
    circuito: texto(input.circuito, base.circuito),
    administradora: texto(input.administradora, base.administradora),
    complejo: texto(input.complejo, base.complejo),
    estado: texto(input.estado, base.estado),
    ciudad: texto(input.ciudad, base.ciudad),
    sala: texto(input.sala, base.sala),
    aforo: Number.isFinite(aforo) ? Math.max(0, Math.trunc(aforo)) : 0,
    base2026: Number.isFinite(base2026) ? base2026 : null,
  };
}

function siguienteNro() {
  return Math.max(0, ...maestro.map((m) => Number(m.nro) || 0)) + 1;
}

// ---------------------------------------------------------------------------
// Cobertura declarada en el nombre del archivo (ej. ...20250102-20251231.xls)
// ---------------------------------------------------------------------------
function parseFechaDMY(value) {
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(String(value || '').trim());
  if (!m) return null;
  const d = Number(m[1]);
  const mo = Number(m[2]);
  const y = Number(m[3]);
  const dt = new Date(y, mo - 1, d);
  if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== d) return null;
  return dt;
}

function formatearDMY(dt) {
  return `${String(dt.getDate()).padStart(2, '0')}/${String(dt.getMonth() + 1).padStart(2, '0')}/${dt.getFullYear()}`;
}

function sumarDias(dt, dias) {
  return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate() + dias);
}

function rangoDeNombre(name) {
  const m = /(\d{8})-(\d{8})/.exec(String(name || ''));
  if (!m) return null;
  const toDate = (s) => new Date(Number(s.slice(0, 4)), Number(s.slice(4, 6)) - 1, Number(s.slice(6, 8)));
  const parse = (s) => {
    const dt = toDate(s);
    if (Number.isNaN(dt.getTime())) return null;
    return formatearDMY(dt) === `${s.slice(6, 8)}/${s.slice(4, 6)}/${s.slice(0, 4)}` ? dt : null;
  };
  const desde = parse(m[1]);
  const hasta = parse(m[2]);
  if (!desde || !hasta || hasta < desde) return null;
  // Un reporte anual no puede cubrir más de ~400 días; evita rangos absurdos.
  if ((hasta - desde) / 86400000 > 400) return null;
  return { desde: formatearDMY(desde), hasta: formatearDMY(hasta) };
}

function calcularCobertura(archivos, rows) {
  const declarados = archivos
    .map((a) => a.rangoDeclarado)
    .filter(Boolean)
    .map((r) => ({ desde: parseFechaDMY(r.desde), hasta: parseFechaDMY(r.hasta) }))
    .filter((r) => r.desde && r.hasta);

  const fechas = rows.map((r) => parseFechaDMY(r.Fecha)).filter(Boolean).sort((a, b) => a - b);
  const detectado = fechas.length
    ? { desde: formatearDMY(fechas[0]), hasta: formatearDMY(fechas[fechas.length - 1]) }
    : null;

  if (!declarados.length || !detectado) {
    return { esperado: null, detectado, faltantes: [], totalFaltantes: 0, completa: null };
  }

  const esperadoDesde = declarados.reduce((min, r) => (r.desde < min ? r.desde : min), declarados[0].desde);
  const esperadoHasta = declarados.reduce((max, r) => (r.hasta > max ? r.hasta : max), declarados[0].hasta);
  const realDesde = fechas[0];
  const realHasta = fechas[fechas.length - 1];
  const faltantes = [];

  // Solo se reportan huecos en los extremos: el reporte puede no operar todos
  // los días (feriados, cierres) y eso no es un error de cobertura.
  for (let d = esperadoDesde; d < realDesde; d = sumarDias(d, 1)) faltantes.push(formatearDMY(d));
  for (let d = sumarDias(realHasta, 1); d <= esperadoHasta; d = sumarDias(d, 1)) faltantes.push(formatearDMY(d));

  return {
    esperado: { desde: formatearDMY(esperadoDesde), hasta: formatearDMY(esperadoHasta) },
    detectado,
    faltantes: faltantes.slice(0, 31),
    totalFaltantes: faltantes.length,
    completa: faltantes.length === 0,
  };
}

function nuevoIdCorrida() {
  return `${new Date().toISOString().replace(/[:.]/g, '-')}-${crypto.randomBytes(3).toString('hex')}`;
}

// ---------------------------------------------------------------------------
// API
// ---------------------------------------------------------------------------
app.get('/api/health', (req, res) => res.json({ ok: true, version: APP_VERSION }));

// Maestro CRUD persistente
app.get('/api/maestro', (req, res) => res.json(maestro));

app.post('/api/maestro', (req, res) => {
  const item = sanitizeMaestro(req.body || {}, { nro: siguienteNro() });
  if (!item.complejo || !item.sala) {
    return res.status(400).json({ error: 'complejo y sala son obligatorios' });
  }
  if (maestroMap().has(claveMaestro(item))) {
    return res.status(409).json({ error: 'Esa sala ya existe en el maestro' });
  }

  maestro.push(item);
  try {
    guardarMaestro();
  } catch (e) {
    maestro.pop();
    return res.status(500).json({ error: `No se pudo guardar el maestro: ${e.message}` });
  }
  res.status(201).json({ ok: true, item, index: maestro.length - 1 });
});

app.put('/api/maestro/:idx', (req, res) => {
  const i = Number(req.params.idx);
  if (!Number.isInteger(i) || i < 0 || i >= maestro.length) {
    return res.status(404).json({ error: 'Sala no encontrada' });
  }

  const anterior = maestro[i];
  const item = sanitizeMaestro(req.body || {}, anterior);
  if (!item.complejo || !item.sala) {
    return res.status(400).json({ error: 'complejo y sala son obligatorios' });
  }
  const key = claveMaestro(item);
  if (maestro.some((m, j) => j !== i && claveMaestro(m) === key)) {
    return res.status(409).json({ error: 'Otra sala ya usa ese complejo y número de sala' });
  }

  maestro[i] = item;
  try {
    guardarMaestro();
  } catch (e) {
    maestro[i] = anterior;
    return res.status(500).json({ error: `No se pudo guardar el maestro: ${e.message}` });
  }
  res.json({ ok: true, item, index: i });
});

app.delete('/api/maestro/:idx', (req, res) => {
  const i = Number(req.params.idx);
  if (!Number.isInteger(i) || i < 0 || i >= maestro.length) {
    return res.status(404).json({ error: 'Sala no encontrada' });
  }
  const [eliminado] = maestro.splice(i, 1);
  try {
    guardarMaestro();
  } catch (e) {
    maestro.splice(i, 0, eliminado);
    return res.status(500).json({ error: `No se pudo guardar el maestro: ${e.message}` });
  }
  res.json({ ok: true, eliminado, index: i });
});

// Subir uno o varios reportes CNAC (.xls, .xlsx, .ods, .csv, .tsv)
app.post('/api/upload', upload.any(), async (req, res) => {
  try {
    if (!req.files?.length) return res.status(400).json({ error: 'No se recibió ningún archivo' });

    const anio = Number(req.body.anio || 2026);
    if (!Number.isInteger(anio) || anio < 2000 || anio > 2100) {
      return res.status(400).json({ error: 'Año inválido' });
    }
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
      // El hash se calcula antes de borrar el temporal y queda en el manifiesto
      // como evidencia de qué archivo exacto produjo la corrida.
      const sha256 = crypto.createHash('sha256').update(fs.readFileSync(f.path)).digest('hex');
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
        sha256,
        filas: rows.length,
        filasValidas: rows.length - fileInvalid.length,
        filasInvalidas: fileInvalid.length,
        errores: errors.length,
        avisos: warnings.length,
        semanas: parsed.semanas.size,
        rango: parsed.rango,
        rangoDeclarado: rangoDeNombre(fileName),
      });
    }

    const incluirTodosCircuitos = req.body.incluirTodosCircuitos === 'true' || req.body.incluirTodosCircuitos === true;
    const validRows = allRows.filter((r) => !invalidRows.has(r));
    const agg = aggregateBySala(validRows);
    const pelven = calcPelven(agg, maestroMap(), { factorDiario, factoresPorComplejo, incluirTodosCircuitos });
    const circuitosDetectados = Array.from(new Set(pelven.map((r) => r.circuito).filter(Boolean)));
    const duplicados = findDuplicateRows(validRows, { limit: 100 });
    const cobertura = calcularCobertura(perFile, validRows);

    if (cobertura.totalFaltantes > 0) {
      const muestra = cobertura.faltantes.slice(0, 10).join(', ');
      allWarnings.push({
        archivo: '',
        fila: '',
        campo: 'Cobertura',
        valor: `Faltan ${cobertura.totalFaltantes} día(s) del rango declarado en el nombre: ${muestra}`,
      });
    }

    const bloqueado = allErrors.length > 0;
    const manifiesto = {
      id: nuevoIdCorrida(),
      fecha: new Date().toISOString(),
      version: APP_VERSION,
      parametros: { anio, factorDiario, factoresPorComplejo, incluirTodosCircuitos },
      archivos: perFile,
      totales: {
        filas: allRows.length,
        filasValidas: validRows.length,
        filasInvalidas: invalidRows.size,
        errores: allErrors.length,
        avisos: allWarnings.length,
        duplicadosGrupos: duplicados.grupos,
        duplicadosFilasExtra: duplicados.filasExtra,
      },
      salidas: {
        salas: pelven.length,
        circuitos: circuitosDetectados.length,
        bloqueado,
      },
      cobertura,
      duplicados: { grupos: duplicados.grupos, filasExtra: duplicados.filasExtra },
      guardado: true,
    };

    try {
      escribirJsonAtomico(path.join(runsDir, `${manifiesto.id}.json`), {
        ...manifiesto,
        errores: allErrors,
        avisos: allWarnings,
      });
    } catch (e) {
      console.error(`No se pudo guardar el manifiesto: ${e.message}`);
      manifiesto.guardado = false;
      allWarnings.push({
        archivo: '',
        fila: '',
        campo: 'Manifiesto',
        valor: `No se pudo guardar el manifiesto de esta corrida: ${e.message}`,
      });
      manifiesto.totales.avisos = allWarnings.length;
    }

    res.json({
      archivos: perFile,
      totalFilas: allRows.length,
      filasValidas: validRows.length,
      filasInvalidas: invalidRows.size,
      salasDetectadas: Object.keys(agg).length,
      circuitosDetectados,
      bloqueado,
      errores: allErrors.slice(0, 200),
      totalErrores: allErrors.length,
      avisos: allWarnings.slice(0, 200),
      totalAvisos: allWarnings.length,
      duplicados,
      cobertura,
      manifiesto,
      preview: pelven.slice(0, 20),
      _full: pelven,
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: String(e.message || e) });
  } finally {
    // Limpieza defensiva: cubre errores tempranos y archivos que no llegaron a parsearse.
    for (const f of req.files || []) fs.unlink(f.path, () => {});
  }
});

// Historial de corridas (manifiestos persistidos)
app.get('/api/runs', (req, res) => {
  try {
    const resumenes = fs.readdirSync(runsDir)
      .filter((f) => f.endsWith('.json'))
      .map((file) => {
        try {
          const m = JSON.parse(fs.readFileSync(path.join(runsDir, file), 'utf8'));
          return {
            id: m.id,
            fecha: m.fecha,
            version: m.version,
            archivos: (m.archivos || []).map((a) => ({ file: a.file, filas: a.filas, sha256: a.sha256 })),
            totales: m.totales,
            salidas: m.salidas,
            bloqueado: m.salidas?.bloqueado ?? false,
            guardado: m.guardado !== false,
          };
        } catch {
          return null;
        }
      })
      .filter(Boolean)
      .sort((a, b) => String(b.fecha || '').localeCompare(String(a.fecha || '')))
      .slice(0, 30);
    res.json(resumenes);
  } catch (e) {
    res.status(500).json({ error: String(e.message || e) });
  }
});

app.get('/api/runs/:id', (req, res) => {
  const id = String(req.params.id || '');
  if (!/^[A-Za-z0-9._-]+$/.test(id)) return res.status(400).json({ error: 'ID inválido' });
  const file = path.join(runsDir, `${id}.json`);
  if (!fs.existsSync(file)) return res.status(404).json({ error: 'Corrida no encontrada' });
  try {
    res.json(JSON.parse(fs.readFileSync(file, 'utf8')));
  } catch (e) {
    res.status(500).json({ error: String(e.message || e) });
  }
});

// Generar y descargar consolidado
app.post('/api/export', async (req, res) => {
  try {
    const { filas, anio = 2026, bloqueado, sufijo = '' } = req.body || {};
    if (bloqueado) return res.status(409).json({ error: 'Exportación bloqueada: hay errores de validación sin resolver' });
    const year = Number(anio);
    if (!Number.isInteger(year) || year < 2000 || year > 2100) return res.status(400).json({ error: 'Año inválido' });
    if (!Array.isArray(filas) || !filas.length) return res.status(400).json({ error: 'filas vacías' });
    const sufijoLimpio = String(sufijo || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 30);
    const buf = await buildWorkbook(filas, year);
    const nombre = `Pelven_aplica_${year}${sufijoLimpio ? '_' + sufijoLimpio : ''}.xlsx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=${nombre}`);
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
