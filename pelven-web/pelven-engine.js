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
const norm = (s) => String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase(); // sin tildes: 'C.U. EL MARQUÉS' == 'C.U. EL MARQUES'
const num = (s) => {
  if (s == null || s === '') return 0;
  const v = parseFloat(String(s).replace(/\./g, '').replace(',', '.'));
  return Number.isFinite(v) ? v : NaN;
};

// Parser universal para texto delimitado (CSV con comas, punto y coma, tabuladores, y soporte de comillas)
export function parseDelimitedText(text) {
  const lines = text.split(/\r?\n/).filter((l) => l.trim().length > 0);
  if (!lines.length) return [];

  // Detección automática del separador: tab (\t), punto y coma (;), coma (,), tubería (|)
  const firstLine = lines[0];
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
    const result = [];
    let current = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      if (char === '"' || char === "'") {
        if (inQuotes && line[i + 1] === char) {
          current += char;
          i++; // comilla escapada doble
        } else {
          inQuotes = !inQuotes;
        }
      } else if (char === delim && !inQuotes) {
        result.push(current.trim());
        current = '';
      } else {
        current += char;
      }
    }
    result.push(current.trim());
    return result;
  }

  const headers = splitLine(lines[0], bestDelim).map((h) => h.replace(/^["']|["']$/g, '').trim());
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const values = splitLine(lines[i], bestDelim);
    const obj = {};
    headers.forEach((h, idx) => {
      let val = values[idx] ?? '';
      val = val.replace(/^["']|["']$/g, '').trim();
      obj[h] = val;
    });
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
    rows = sheetData.map(normalizeRow);
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
        const o = {};
        header.forEach((h, i) => (o[h] = vals[i] ?? ''));
        rows.push(normalizeRow(o));
      }
    } catch {
      const wb = XLSX.read(buf, { type: 'buffer' });
      const sheetName = wb.SheetNames[0];
      const sheetData = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '', raw: false });
      rows = sheetData.map(normalizeRow);
    }
  }
  // CASO 3: Archivo XLS binario antiguo (OLE2 BIFF8)
  else if (isOle) {
    const wb = XLSX.read(buf, { type: 'buffer' });
    const sheetName = wb.SheetNames[0];
    const sheetData = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '', raw: false });
    rows = sheetData.map(normalizeRow);
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
  fechas.sort((a, b) => (toKey(a) < toKey(b) ? -1 : 1));
  const semanas = new Set(rows.map((r) => semanaCine(r.Fecha)).filter(Boolean));
  return { rows, semanas, rango: fechas.length ? [fechas[0], fechas[fechas.length - 1]] : [] };
}

function pick(o, ...names) {
  for (const n of names) {
    const k = Object.keys(o).find((k) => norm(k) === norm(n));
    if (k) return o[k];
  }
  return '';
}

function normalizeRow(o) {
  return {
    Fecha: String(pick(o, 'Fecha') ?? '').trim(),
    Complejo: String(pick(o, 'Complejo', 'Complejo Cinematográfico') ?? '').trim(),
    Sala: String(pick(o, 'Salas', 'N° de Sala', 'Sala') ?? '').trim(),
    Registro: String(pick(o, 'N° Registro de la Obra', 'N Registro') ?? '').trim(),
    Obra: String(pick(o, 'Obra Cinematografica', 'Título de la Obra Cinematográfica') ?? '').trim(),
    Formato: String(pick(o, 'Formato de Proyeccion', 'Formato de Proyección') ?? '').trim(),
    Hora: String(pick(o, 'Hora', 'Hora de Exhibición') ?? '').trim(),
    Funcion: String(pick(o, 'Función', 'Funcion') ?? '').trim(),
    TipoBoleto: String(pick(o, 'Tipo de Boleto') ?? '').trim(),
    PrecioBoleto: pick(o, 'Precio Boleto', 'Precio de Boleto Bs.'),
    PrecioServicio: pick(o, 'Precio Servicio', 'Precio Servicio Bs.'),
    OtrosCargos: pick(o, 'Otros Cargos', 'Otros Cargos Bs.'),
    OtrosIva: pick(o, 'Otros Cargos IVA', 'Otros Cargos Iva Alícuota (%)'),
    TotalPrecio: pick(o, 'Total Precio', 'Total Precio de Boleto Bs.'),
    Cantidad: pick(o, 'Cantidad', 'N° de Espectadores'),
    TotalRecaudacion: pick(o, 'Total Recaudación', 'Total Recaudacion Bs.'),
  };
}

export function validateRows(rows) {
  const errors = [];
  const warnings = [];
  rows.forEach((r, i) => {
    const n = i + 2;
    if (!/^\d{2}\/\d{2}\/\d{4}$/.test(r.Fecha)) errors.push({ fila: n, campo: 'Fecha', valor: r.Fecha });
    else {
      const [d, m, y] = r.Fecha.split('/').map(Number);
      if (m < 1 || m > 12 || d < 1 || d > 31) errors.push({ fila: n, campo: 'Fecha', valor: r.Fecha });
    }
    if (!/^\d{2}:\d{2}$/.test(r.Hora)) errors.push({ fila: n, campo: 'Hora', valor: r.Hora });
    if (!FUNCIONES.map(norm).includes(norm(r.Funcion))) errors.push({ fila: n, campo: 'Funcion', valor: r.Funcion });
    if (!TIPOS.map(norm).includes(norm(r.TipoBoleto))) errors.push({ fila: n, campo: 'TipoBoleto', valor: r.TipoBoleto });
    const cant = Number(String(r.Cantidad).replace(',', '.'));
    if (!Number.isInteger(cant) || cant < 0) errors.push({ fila: n, campo: 'Cantidad', valor: r.Cantidad });
    const pb = num(r.PrecioBoleto), ps = num(r.PrecioServicio), oc = num(r.OtrosCargos), tp = num(r.TotalPrecio), tr = num(r.TotalRecaudacion);
    if ([pb, ps, oc, tp, tr].some((v) => Number.isNaN(v))) errors.push({ fila: n, campo: 'Montos', valor: `${r.PrecioBoleto}|${r.PrecioServicio}|${r.OtrosCargos}|${r.TotalPrecio}|${r.TotalRecaudacion}` });
    else {
      // Evidencia empírica (fila Otros Cargos 11.22 IVA 16 → Total 169.99):
      // PrecioBoleto y PrecioServicio ya incluyen IVA; OtrosCargos se suma con su IVA.
      const ocIva = Number(String(r.OtrosIva ?? 0).replace(',', '.')) || 0;
      const esperado = pb + ps + oc * (1 + ocIva / 100);
      if (Math.abs(tp - esperado) > 0.03 && norm(r.TipoBoleto) !== 'SIN FUNCION') warnings.push({ fila: n, campo: 'TotalPrecio≠Boleto+Servicio+Otros*(1+IVA)', valor: tp, esperado: +esperado.toFixed(2) });
      if (Math.abs(tr - tp * cant) > Math.max(1, tp * cant * 0.001)) errors.push({ fila: n, campo: 'TotalRecaudación≠Precio*Cantidad', valor: tr });
    }
  });
  return { errors, warnings };
}

export function semanaCine(fechaDDMMAAAA) {
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(fechaDDMMAAAA || '');
  if (!m) return '';
  const dt = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  const delta = (dt.getDay() - 5 + 7) % 7; // viernes=5
  dt.setDate(dt.getDate() - delta);
  return dt.toISOString().slice(0, 10);
}

const FUNCIONES_BASE = ['VESPERTINA', 'INTERMEDIA', 'NOCHE'];

export function aggregateBySala(rows) {
  const agg = {};
  for (const r of rows) {
    if (norm(r.Funcion) === 'SIN FUNCION') continue; // no genera espectadores ni funciones
    const key = norm(r.Complejo) + '|' + String(r.Sala).trim();
    agg[key] ??= { complejo: r.Complejo.trim(), sala: String(r.Sala).trim(), total: 0, totalBase: 0, diasBase: new Set(), dias: new Set(), funciones: new Set(), porSemana: {} };
    const c = Number(String(r.Cantidad).replace(',', '.')) || 0;
    agg[key].total += c;
    if (FUNCIONES_BASE.includes(norm(r.Funcion).normalize('NFD').replace(/[\u0300-\u036f]/g, ''))) {
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
      complejo: a.complejo, sala: a.sala, totalEspectadores: a.total, diasActivos: dias,
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
        complejo: m.complejo, sala: String(m.sala), totalEspectadores: 0, diasActivos: 0, funcionesDistintas: 0,
        base: 0, rawBase: 0, factorUsado: factorDiario,
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
    if (x.nro && y.nro) return x.nro - y.nro;
    if (x.nro) return -1;
    if (y.nro) return 1;

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
      i + 1,
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
