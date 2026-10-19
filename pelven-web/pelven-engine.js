// Motor PELVEN: parseo, validación, agregación y cálculo.
// Fórmula base deducida por regresión contra "Pelven aplica 2026" (164 salas de Cines Unidos):
//   base(sala) = (espectadores en funciones Vespertina+Intermedia+Noche / días con alguna de esas funciones) * 4.2
// Matinee, Media Noche, Continuada y Sin Función NO cuentan. 4.2 = 21/5. Columnas 20..1 = base * n/21 (exacto).
// Verificación: error mediano 0,19 %, máximo 1,5 % en las 164 salas (incluye Metrocenter sin factor especial).
import fs from 'fs';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';

const FUNCIONES = ['Matinee', 'Vespertina', 'Intermedia', 'Noche', 'Media Noche', 'Continuada', 'Sin Función'];
const TIPOS = ['Completo', 'Popular', 'Niños', 'Tercera Edad', 'Promoción', 'Cortesía', 'Sin Función'];
const normCache = new Map();
const norm = (s) => {
  const key = String(s ?? '');
  const cached = normCache.get(key);
  if (cached !== undefined) return cached;
  const value = key.normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase(); // sin tildes: 'C.U. EL MARQUÉS' == 'C.U. EL MARQUES'
  normCache.set(key, value);
  return value;
};

// Acepta números nativos de Excel y textos en formato es-VE (1.234,56) o en-US (1,234.56).
// Un único punto con exactamente 3 decimales se interpreta como separador de miles (1.234 -> 1234).
export function parseNumber(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : NaN;
  if (value == null) return NaN;

  let s = String(value).trim();
  if (!s) return NaN;

  s = s.replace(/\u00a0/g, ' ').replace(/\s+/g, '');
  let negative = /^\(.*\)$/.test(s);
  if (negative) s = s.slice(1, -1);
  if (s.startsWith('-')) negative = true;
  s = s.replace(/[^\d.,]/g, '');
  if (!s || !/\d/.test(s)) return NaN;

  const hasDot = s.includes('.');
  const hasComma = s.includes(',');
  if (hasDot && hasComma) {
    // El último separador es el decimal; el otro es de miles.
    if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g, '').replace(/,/g, '.');
    else s = s.replace(/,/g, '');
  } else if (hasComma) {
    const parts = s.split(',');
    if (parts.length > 2 || (parts.length === 2 && parts[1].length === 3 && parts[0].length > 3)) s = s.replace(/,/g, '');
    else s = s.replace(',', '.');
  } else if (hasDot) {
    const parts = s.split('.');
    if (parts.length > 2) s = s.replace(/\./g, '');
    else if (parts[1].length === 3 && parts[0].length <= 3 && parts[0] !== '0') s = s.replace('.', '');
  }

  const n = Number(s);
  if (!Number.isFinite(n)) return NaN;
  return negative ? -n : n;
}

// Parser para texto delimitado (TSV/CSV con tab, ;, , o |).
// Solo la comilla doble abre/cierra un campo, y solo al inicio del campo. El apóstrofo es texto normal.
export function parseDelimitedText(text) {
  const lines = String(text ?? '').replace(/^\uFEFF/, '').split(/\r?\n/);
  const headerIndex = lines.findIndex((l) => l.trim().length > 0);
  if (headerIndex === -1) return [];

  // Detección automática del separador sobre la primera línea no vacía.
  const firstLine = lines[headerIndex];
  const candidates = ['\t', ';', ',', '|'];
  let bestDelim = '\t';
  let maxCount = -1;
  for (const d of candidates) {
    const count = (firstLine.match(new RegExp('\\' + d, 'g')) || []).length;
    if (count > maxCount) {
      maxCount = count;
      bestDelim = d;
    }
  }

  function splitLine(line, delim) {
    // Camino rápido: sin comillas no hace falta recorrer carácter por carácter.
    if (!line.includes('"')) return { values: line.split(delim).map((v) => v.trim()), warning: '' };

    const result = [];
    let current = '';
    let inQuotes = false;
    let atFieldStart = true;
    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      if (inQuotes) {
        if (char === '"') {
          if (line[i + 1] === '"') {
            current += '"';
            i++; // comilla escapada ("" -> ")
          } else {
            inQuotes = false;
          }
        } else {
          current += char;
        }
      } else if (char === '"' && atFieldStart) {
        inQuotes = true;
        atFieldStart = false;
      } else if (char === delim) {
        result.push(current.trim());
        current = '';
        atFieldStart = true;
      } else {
        current += char;
        if (!/\s/.test(char)) atFieldStart = false;
      }
    }

    if (inQuotes) {
      // Una comilla desbalanceada no debe fusionar el resto de columnas.
      return { values: line.split(delim).map((v) => v.trim().replace(/^"|"$/g, '')), warning: 'comillas desbalanceadas' };
    }

    result.push(current.trim());
    return { values: result, warning: '' };
  }

  const headers = splitLine(firstLine, bestDelim).values.map((h) => h.trim());
  const rows = [];
  for (let i = headerIndex + 1; i < lines.length; i++) {
    const line = lines[i];
    if (!line.trim()) continue; // se conserva el número de línea original
    const { values, warning } = splitLine(line, bestDelim);
    if (values.every((v) => v === '')) continue;
    const obj = {};
    headers.forEach((h, idx) => {
      obj[h] = values[idx] ?? '';
    });
    obj.__line = i + 1;
    if (warning) obj.__parseWarning = warning;
    rows.push(obj);
  }
  return rows;
}

export async function parseTaquillaFile(filePath, originalName = '') {
  const buf = fs.readFileSync(filePath);
  const ext = (originalName.split('.').pop() || '').toLowerCase();
  const head = buf.slice(0, 4);
  const isZip = head[0] === 0x50 && head[1] === 0x4b; // PK (zip: .xlsx, .ods)
  const isOle = head[0] === 0xd0 && head[1] === 0xcf; // OLE2 (.xls binario antiguo)

  let rows = [];

  // CASO 1: Archivo OpenDocument Spreadsheet (.ods)
  if (ext === 'ods' || (isZip && buf.includes(Buffer.from('opendocument.spreadsheet')))) {
    const wb = XLSX.read(buf, { type: 'buffer' });
    const sheetName = wb.SheetNames[0];
    const sheetData = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '', raw: false });
    rows = sheetData.map((o, i) => normalizeRow({ ...o, __line: i + 2 }));
  }
  // CASO 2: Archivo Excel OpenXML (.xlsx)
  else if (ext === 'xlsx' || isZip) {
    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buf);
      const ws = wb.worksheets[0];
      const header = ws.getRow(1).values.slice(1).map((h) => String(h ?? '').trim());
      for (let r = 2; r <= ws.rowCount; r++) {
        const vals = ws.getRow(r).values.slice(1);
        const o = { __line: r };
        header.forEach((h, i) => (o[h] = vals[i] ?? ''));
        rows.push(normalizeRow(o));
      }
    } catch {
      const wb = XLSX.read(buf, { type: 'buffer' });
      const sheetName = wb.SheetNames[0];
      const sheetData = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '', raw: false });
      rows = sheetData.map((o, i) => normalizeRow({ ...o, __line: i + 2 }));
    }
  }
  // CASO 3: Archivo XLS binario antiguo (OLE2 BIFF8)
  else if (isOle) {
    const wb = XLSX.read(buf, { type: 'buffer' });
    const sheetName = wb.SheetNames[0];
    const sheetData = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '', raw: false });
    rows = sheetData.map((o, i) => normalizeRow({ ...o, __line: i + 2 }));
  }
  // CASO 4: Archivos de texto delimitado (.csv, .tsv, .txt, o .xls TSV de Cines Unidos)
  else {
    let text = buf.toString('utf-8').replace(/^\uFEFF/, '');
    if (text.includes('\uFFFD') || !text.includes('\n')) {
      text = buf.toString('latin1').replace(/^\uFEFF/, '');
    }
    const rawRows = parseDelimitedText(text);
    rows = rawRows.map(normalizeRow);
  }

  const fechas = rows.map((r) => r.Fecha).filter(Boolean);
  const toKey = (f) => f.slice(6, 10) + f.slice(3, 5) + f.slice(0, 2);
  fechas.sort((a, b) => (toKey(a) < toKey(b) ? -1 : toKey(a) > toKey(b) ? 1 : 0));
  const semanas = new Set(rows.map((r) => semanaCine(r.Fecha)).filter(Boolean));
  return { rows, semanas, rango: fechas.length ? [fechas[0], fechas[fechas.length - 1]] : [] };
}

const FIELD_ALIASES = {
  Fecha: ['Fecha'],
  Complejo: ['Complejo', 'Complejo Cinematográfico'],
  Sala: ['Salas', 'N° de Sala', 'Sala'],
  Registro: ['N° Registro de la Obra', 'N Registro'],
  Obra: ['Obra Cinematografica', 'Título de la Obra Cinematográfica'],
  Formato: ['Formato de Proyeccion', 'Formato de Proyección'],
  Hora: ['Hora', 'Hora de Exhibición'],
  Funcion: ['Función', 'Funcion'],
  TipoBoleto: ['Tipo de Boleto'],
  PrecioBoleto: ['Precio Boleto', 'Precio de Boleto Bs.'],
  PrecioServicio: ['Precio Servicio', 'Precio Servicio Bs.'],
  OtrosCargos: ['Otros Cargos', 'Otros Cargos Bs.'],
  OtrosIva: ['Otros Cargos IVA', 'Otros Cargos Iva Alícuota (%)'],
  TotalPrecio: ['Total Precio', 'Total Precio de Boleto Bs.'],
  Cantidad: ['Cantidad', 'N° de Espectadores'],
  TotalRecaudacion: ['Total Recaudación', 'Total Recaudacion Bs.'],
};
for (const field of Object.keys(FIELD_ALIASES)) FIELD_ALIASES[field] = FIELD_ALIASES[field].map(norm);

export function normalizeRow(o = {}) {
  const lookup = new Map();
  for (const [key, value] of Object.entries(o)) {
    const nk = norm(key);
    if (!lookup.has(nk)) lookup.set(nk, value);
  }
  const get = (field) => {
    for (const alias of FIELD_ALIASES[field]) {
      if (lookup.has(alias)) return lookup.get(alias);
    }
    return '';
  };

  return {
    Fecha: String(get('Fecha') ?? '').trim(),
    Complejo: String(get('Complejo') ?? '').trim(),
    Sala: String(get('Sala') ?? '').trim(),
    Registro: String(get('Registro') ?? '').trim(),
    Obra: String(get('Obra') ?? '').trim(),
    Formato: String(get('Formato') ?? '').trim(),
    Hora: String(get('Hora') ?? '').trim(),
    Funcion: String(get('Funcion') ?? '').trim(),
    TipoBoleto: String(get('TipoBoleto') ?? '').trim(),
    PrecioBoleto: get('PrecioBoleto'),
    PrecioServicio: get('PrecioServicio'),
    OtrosCargos: get('OtrosCargos'),
    OtrosIva: get('OtrosIva'),
    TotalPrecio: get('TotalPrecio'),
    Cantidad: get('Cantidad'),
    TotalRecaudacion: get('TotalRecaudacion'),
    Linea: Number(o?.__line) || 0,
    ParseWarning: String(o?.__parseWarning ?? '').trim(),
  };
}

const FUNCIONES_NORM = new Set(FUNCIONES.map(norm));
const TIPOS_NORM = new Set(TIPOS.map(norm));

function isRealDate(d, m, y) {
  const dt = new Date(y, m - 1, d);
  return dt.getFullYear() === y && dt.getMonth() === m - 1 && dt.getDate() === d;
}

export function validateRows(rows) {
  const errors = [];
  const warnings = [];
  const invalidRows = [];

  rows.forEach((r, i) => {
    const fila = r.Linea || i + 2;
    const archivo = r._file || '';
    let invalid = false;
    const err = (campo, valor, extra = {}) => {
      errors.push({ archivo, fila, campo, valor, ...extra });
      invalid = true;
    };
    const warn = (campo, valor, extra = {}) => {
      warnings.push({ archivo, fila, campo, valor, ...extra });
    };

    if (!/^\d{2}\/\d{2}\/\d{4}$/.test(r.Fecha)) err('Fecha', r.Fecha);
    else {
      const [d, m, y] = r.Fecha.split('/').map(Number);
      if (!isRealDate(d, m, y)) err('Fecha', r.Fecha);
    }

    if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(r.Hora)) err('Hora', r.Hora);
    if (!FUNCIONES_NORM.has(norm(r.Funcion))) err('Funcion', r.Funcion);
    if (!TIPOS_NORM.has(norm(r.TipoBoleto))) err('TipoBoleto', r.TipoBoleto);
    if (!r.Complejo) err('Complejo', r.Complejo);
    if (!r.Sala) err('Sala', r.Sala);

    const cant = parseNumber(r.Cantidad);
    if (!Number.isInteger(cant) || cant < 0) err('Cantidad', r.Cantidad);

    const pb = parseNumber(r.PrecioBoleto);
    const ps = parseNumber(r.PrecioServicio);
    const tp = parseNumber(r.TotalPrecio);
    const tr = parseNumber(r.TotalRecaudacion);
    for (const [campo, raw, value] of [
      ['PrecioBoleto', r.PrecioBoleto, pb],
      ['PrecioServicio', r.PrecioServicio, ps],
      ['TotalPrecio', r.TotalPrecio, tp],
      ['TotalRecaudacion', r.TotalRecaudacion, tr],
    ]) {
      if (!Number.isFinite(value)) err(campo, raw);
      else if (value < 0) err(campo, raw);
    }

    // Otros Cargos es opcional; si no viene se asume 0. Si viene, debe ser numérico.
    let oc = 0;
    if (String(r.OtrosCargos ?? '').trim() !== '') {
      oc = parseNumber(r.OtrosCargos);
      if (!Number.isFinite(oc) || oc < 0) err('OtrosCargos', r.OtrosCargos);
    }

    let ocIva = 0;
    if (String(r.OtrosIva ?? '').trim() !== '') {
      ocIva = parseNumber(r.OtrosIva);
      if (!Number.isFinite(ocIva) || ocIva < 0) warn('OtrosIva', r.OtrosIva);
      else ocIva = ocIva || 0;
    }

    if ([pb, ps, oc, tp].every(Number.isFinite)) {
      // Evidencia empírica (fila Otros Cargos 11.22 IVA 16 → Total 169.99):
      // PrecioBoleto y PrecioServicio ya incluyen IVA; OtrosCargos se suma con su IVA.
      const esperado = pb + ps + oc * (1 + ocIva / 100);
      if (Math.abs(tp - esperado) > 0.03 && norm(r.TipoBoleto) !== 'SIN FUNCION') {
        warn('TotalPrecio≠Boleto+Servicio+Otros*(1+IVA)', tp, { esperado: +esperado.toFixed(2) });
      }
    }

    if (Number.isFinite(tp) && Number.isFinite(tr) && Number.isInteger(cant) && cant >= 0) {
      const esperadoTr = tp * cant;
      if (Math.abs(tr - esperadoTr) > Math.max(1, Math.abs(esperadoTr) * 0.001)) {
        err('TotalRecaudación≠Precio*Cantidad', tr, { esperado: +esperadoTr.toFixed(2) });
      }
    }

    if (r.ParseWarning) warn('Formato', r.ParseWarning);
    if (invalid) invalidRows.push(r);
  });

  return { errors, warnings, invalidRows };
}

export function semanaCine(fechaDDMMAAAA) {
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(fechaDDMMAAAA || '');
  if (!m) return '';
  const dt = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  const delta = (dt.getDay() - 5 + 7) % 7; // viernes=5
  dt.setDate(dt.getDate() - delta);
  const y = dt.getFullYear();
  const mm = String(dt.getMonth() + 1).padStart(2, '0');
  const dd = String(dt.getDate()).padStart(2, '0');
  return `${y}-${mm}-${dd}`;
}

const FUNCIONES_BASE = ['VESPERTINA', 'INTERMEDIA', 'NOCHE'];

export function aggregateBySala(rows) {
  const agg = {};
  for (const r of rows) {
    if (norm(r.Funcion) === 'SIN FUNCION') continue; // no genera espectadores ni funciones
    const key = norm(r.Complejo) + '|' + String(r.Sala).trim();
    agg[key] ??= { complejo: r.Complejo.trim(), sala: String(r.Sala).trim(), total: 0, totalBase: 0, diasBase: new Set(), dias: new Set(), funciones: new Set(), porSemana: {} };
    const c = parseNumber(r.Cantidad);
    if (!Number.isFinite(c)) continue;
    agg[key].total += c;
    if (FUNCIONES_BASE.includes(norm(r.Funcion))) {
      agg[key].totalBase += c;
      agg[key].diasBase.add(r.Fecha);
    }
    agg[key].dias.add(r.Fecha);
    agg[key].funciones.add(r.Fecha + ' ' + r.Hora);
    const s = semanaCine(r.Fecha);
    agg[key].porSemana[s] = (agg[key].porSemana[s] || 0) + c;
  }
  return agg;
}

export const PELVEN_FACTORES = [
  0.9524, 0.9048, 0.8571, 0.8095, 0.7619, 0.7143, 0.6667, 0.6190, 0.5714, 0.5238,
  0.4762, 0.4286, 0.3810, 0.3333, 0.2857, 0.2381, 0.1905, 0.1429, 0.0952, 0.0476
];

// Redondeo reglamentario a mayor a partir de .70:
// Si la parte decimal es >= 0.70 se redondea hacia arriba, de lo contrario hacia abajo.
export function roundAt70(val) {
  if (typeof val !== 'number' || Number.isNaN(val) || !Number.isFinite(val)) return 0;
  const intPart = Math.floor(val);
  const dec = +(val - intPart).toFixed(8);
  return (dec >= 0.70 - 1e-9) ? intPart + 1 : intPart;
}

export function calcPelven(agg, mapMaestro, { semanasPorAnio = 52, factorDiario = 4.2, factoresPorComplejo = {}, incluirTodosCircuitos = false } = {}) {
  const out = [];
  const factorDe = (complejo) => {
    const hit = Object.entries(factoresPorComplejo).find(([k]) => norm(complejo).includes(norm(k)));
    return hit ? Number(hit[1]) : factorDiario;
  };
  for (const [key, a] of Object.entries(agg)) {
    const m = mapMaestro.get(key);
    const dias = a.dias.size || 1;
    const f = factorDe(a.complejo);
    const rawBase = a.diasBase.size ? (a.totalBase / a.diasBase.size) * f : 0;
    const base = roundAt70(rawBase);
    const row = {
      nro: m?.nro ?? null,
      complejo: a.complejo, sala: a.sala, totalEspectadores: a.total, totalBase: a.totalBase,
      diasActivos: dias, diasBase: a.diasBase.size, semanasConData: Object.keys(a.porSemana).length,
      funcionesDistintas: a.funciones.size, base, rawBase, factorUsado: f,
      circuito: m?.circuito ?? '', administradora: m?.administradora ?? '',
      estado: m?.estado ?? '', ciudad: m?.ciudad ?? '', aforo: m?.aforo ?? 0, enMaestro: !!m,
    };
    row['f21'] = base;
    PELVEN_FACTORES.forEach((fac, idx) => {
      row['f' + (20 - idx)] = roundAt70(base * fac);
    });
    out.push(row);
  }

  // Añadir salas del maestro sin data (base 0)
  const circuitosPresentes = new Set(out.filter((r) => r.enMaestro && r.circuito).map((r) => norm(r.circuito)));
  for (const m of mapMaestro.values()) {
    const key = norm(m.complejo) + '|' + String(m.sala);
    const debeIncluir = incluirTodosCircuitos || circuitosPresentes.has(norm(m.circuito));
    if (!agg[key] && debeIncluir) {
      const row = {
        nro: m.nro,
        complejo: m.complejo, sala: String(m.sala), totalEspectadores: 0, totalBase: 0,
        diasActivos: 0, diasBase: 0, semanasConData: 0, funcionesDistintas: 0,
        base: 0, rawBase: 0, factorUsado: factorDe(m.complejo),
        circuito: m.circuito, administradora: m.administradora, estado: m.estado, ciudad: m.ciudad, aforo: m.aforo,
        enMaestro: true, sinData: true
      };
      row['f21'] = 0;
      PELVEN_FACTORES.forEach((_, idx) => { row['f' + (20 - idx)] = 0; });
      out.push(row);
    }
  }

  // Ordenamiento oficial CNAC (reproduce el orden exacto del Pelven oficial):
  // 1. Si ambas salas están en el Maestro, respetar el orden correlativo oficial (Nro 1 a 323).
  // 2. Si hay salas nuevas/sin registro, agrupar por Circuito -> Administradora -> Complejo -> Sala.
  out.sort((x, y) => {
    if (x.nro != null && y.nro != null) return x.nro - y.nro;
    if (x.nro != null) return -1;
    if (y.nro != null) return 1;

    const circComp = norm(x.circuito).localeCompare(norm(y.circuito));
    if (circComp !== 0) return circComp;

    const adminComp = norm(x.administradora).localeCompare(norm(y.administradora));
    if (adminComp !== 0) return adminComp;

    const compComp = norm(x.complejo).localeCompare(norm(y.complejo));
    if (compComp !== 0) return compComp;

    return (Number(x.sala) || 0) - (Number(y.sala) || 0);
  });

  return out;
}

export async function buildWorkbook(filas, anio) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(`publicar ${anio}`, {
    views: [{ showGridLines: true }]
  });

  // Anchos de columna oficiales exactos (1..29)
  const colWidths = [10.5, 16.2, 48.0, 40.8, 18.7, 17.4, 6.1, 6.4, 13.9, ...Array(20).fill(10.5)];
  colWidths.forEach((w, i) => {
    ws.getColumn(i + 1).width = w;
  });

  const fontTitle = { name: 'Arial', size: 10, bold: false, color: { argb: 'FF000000' } };
  const fontHeader = { name: 'Arial', size: 10, bold: false, color: { argb: 'FFFFFFFF' } };
  const fontData = { name: 'Arial', size: 10, bold: false, color: { argb: 'FF000000' } };

  const fillHeader = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF611729' } // Granate / burdeos oficial (97, 23, 41)
  };

  const fillHighlight = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFDEE6EF' } // Azul pastel suave oficial (222, 230, 239)
  };

  const thinBorder = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } }
  };

  // Filas 1..4: Textos de providencia
  const r1 = ws.addRow(['Promedio de Salas / Pelven']);
  const r2 = ws.addRow([`Providencia N° 002, publicada en Gaceta Oficial N° 38.402 de fecha 21/03/2006`]);
  const r3 = ws.addRow(['Providencia para la determinación del promedio de sala para las Obras Cinematográficas Nacionales']);
  const r4 = ws.addRow([`Aplica para el año ${anio}`]);
  [r1, r2, r3, r4].forEach((row) => {
    row.height = 15;
    row.getCell(1).font = fontTitle;
  });

  // Fila 5: Cabeceras principales
  const r5 = ws.addRow(['Nro', 'Circuito', 'Administradora', 'Complejos', 'Estado ', 'Ciudad', 'Salas', 'Aforo', 'FUNCIONES', ...Array(20).fill('')]);
  // Fila 6: Números de funciones (21 a 1)
  const r6 = ws.addRow(['', '', '', '', '', '', '', '', 21, ...Array.from({ length: 20 }, (_, i) => 20 - i)]);
  // Fila 7: Porcentajes / factores
  const r7 = ws.addRow(['', '', '', '', '', '', '', '', '100 % PELVEN', ...PELVEN_FACTORES]);

  [r5, r6, r7].forEach((row) => {
    row.height = 21;
  });

  // Casillas combinadas oficiales
  ws.mergeCells(5, 1, 7, 1); // A5:A7 (Nro)
  ws.mergeCells(5, 2, 7, 2); // B5:B7 (Circuito)
  ws.mergeCells(5, 3, 7, 3); // C5:C7 (Administradora)
  ws.mergeCells(5, 4, 7, 4); // D5:D7 (Complejos)
  ws.mergeCells(5, 5, 7, 5); // E5:E7 (Estado )
  ws.mergeCells(5, 6, 7, 6); // F5:F7 (Ciudad)
  ws.mergeCells(5, 7, 7, 7); // G5:G7 (Salas)
  ws.mergeCells(5, 8, 7, 8); // H5:H7 (Aforo)
  ws.mergeCells(5, 9, 5, 29); // I5:AC5 (FUNCIONES)

  // Aplicar formato de cabecera a TODAS las celdas de las filas 5, 6 y 7
  for (let r = 5; r <= 7; r++) {
    for (let c = 1; c <= 29; c++) {
      const cell = ws.getCell(r, c);
      cell.fill = fillHeader;
      cell.font = fontHeader;
      cell.border = thinBorder;
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      if (r === 7 && c >= 10) {
        cell.numFmt = '0.00%';
      }
    }
  }

  // Filas de datos (a partir de fila 8)
  filas.forEach((r, i) => {
    const rowValues = [
      r.nro != null ? r.nro : i + 1,
      r.circuito,
      r.administradora,
      r.complejo,
      r.estado,
      r.ciudad,
      Number(r.sala) || r.sala,
      Number(r.aforo) || 0,
      r.base,
      ...Array.from({ length: 20 }, (_, k) => r['f' + (20 - k)])
    ];
    const row = ws.addRow(rowValues);
    row.height = 24;

    for (let c = 1; c <= 29; c++) {
      const cell = row.getCell(c);
      cell.font = fontData;
      cell.border = thinBorder;

      if (c === 1) { // Nro
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.numFmt = '0';
      } else if (c >= 2 && c <= 4) { // Circuito, Administradora, Complejos
        cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
      } else if (c === 5 || c === 6) { // Estado, Ciudad
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      } else if (c === 7) { // Salas
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.numFmt = '0';
      } else if (c === 8) { // Aforo
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.numFmt = '#,##0';
      } else if (c === 9) { // FUNCIONES (Base 21) -> Destacado azul pastel suave
        cell.fill = fillHighlight;
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.numFmt = '#,##0';
      } else { // Columnas 20 a 1
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.numFmt = '#,##0';
      }
    }
  });

  return wb.xlsx.writeBuffer();
}
