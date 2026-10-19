import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import {
  parseDelimitedText,
  parseNumber,
  normalizeRow,
  validateRows,
  aggregateBySala,
  calcPelven,
  roundAt70,
  semanaCine,
  buildWorkbook,
  parseTaquillaFile,
  findDuplicateRows,
} from '../pelven-engine.js';

const HEADER = [
  'Fecha',
  'Complejo',
  'Salas',
  'N° Registro de la Obra',
  'Obra Cinematografica',
  'Formato de Proyeccion',
  'Hora',
  'Función',
  'Tipo de Boleto',
  'Precio Boleto',
  'Precio Servicio',
  'Otros Cargos',
  'Otros Cargos IVA',
  'Total Precio',
  'Cantidad',
  'Total Recaudación',
];

const BASE_ROW = {
  Fecha: '31/05/2025',
  Complejo: 'C.U. Lider',
  Salas: '2',
  'N° Registro de la Obra': '1',
  Obra: "HOPE ON THE STAGE' in JAPAN: LIVE VIEWING (Sub)",
  'Formato de Proyeccion': 'D-Cinema 3D',
  Hora: '19:00',
  Función: 'Intermedia',
  'Tipo de Boleto': 'Promoción',
  'Precio Boleto': '409,57',
  'Precio Servicio': '409,57',
  'Otros Cargos': '0',
  'Otros Cargos IVA': '16',
  'Total Precio': '819,14',
  Cantidad: '5',
  'Total Recaudación': '4095,7',
};

function makeRow(overrides = {}) {
  return normalizeRow({ ...BASE_ROW, ...overrides });
}

test('parseDelimitedText no confunde apóstrofos con comillas', () => {
  const text = [HEADER.join('\t'), Object.values(BASE_ROW).join('\t')].join('\n');
  const raw = parseDelimitedText(text);
  assert.equal(raw.length, 1);
  assert.equal(raw[0]['Obra Cinematografica'], "HOPE ON THE STAGE' in JAPAN: LIVE VIEWING (Sub)");
  assert.equal(raw[0]['Hora'], '19:00');
  assert.equal(raw[0]['Función'], 'Intermedia');
  assert.equal(raw[0]['Cantidad'], '5');
  assert.equal(raw[0]['Total Recaudación'], '4095,7');
  assert.equal(raw[0].__line, 2);

  const normalized = normalizeRow(raw[0]);
  const { errors } = validateRows([normalized]);
  assert.deepEqual(errors, []);
});

test('parseDelimitedText maneja comillas dobles y delimitadores embebidos', () => {
  const csv = 'Fecha,Complejo,Sala\n"02/01/2025","C.U. ""El Marqués""",1\n';
  const [row] = parseDelimitedText(csv);
  assert.equal(row.Fecha, '02/01/2025');
  assert.equal(row.Complejo, 'C.U. "El Marqués"');
  assert.equal(row.Sala, '1');
});

test('parseDelimitedText conserva el número de línea original con líneas vacías', () => {
  const text = [HEADER.join('\t'), '', Object.values(BASE_ROW).join('\t')].join('\n');
  const rows = parseDelimitedText(text);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].__line, 3);
});

test('parseNumber interpreta formatos es-VE, en-US y números nativos', () => {
  assert.equal(parseNumber('169.99'), 169.99);
  assert.equal(parseNumber('78,49'), 78.49);
  assert.equal(parseNumber('1.234,56'), 1234.56);
  assert.equal(parseNumber('1,234.56'), 1234.56);
  assert.equal(parseNumber('1.234'), 1234);
  assert.equal(parseNumber('1,234,567'), 1234567);
  assert.equal(parseNumber('-78,49'), -78.49);
  assert.equal(parseNumber('16%'), 16);
  assert.equal(parseNumber(169.99), 169.99);
  assert.ok(Number.isNaN(parseNumber('')));
  assert.ok(Number.isNaN(parseNumber('   ')));
  assert.ok(Number.isNaN(parseNumber('N/A')));
  assert.ok(Number.isNaN(parseNumber(null)));
});

test('validateRows rechaza campos obligatorios vacíos', () => {
  const { errors, invalidRows } = validateRows([makeRow({ Cantidad: '' })]);
  assert.ok(errors.some((e) => e.campo === 'Cantidad'));
  assert.equal(invalidRows.length, 1);
});

test('validateRows rechaza fechas y horas imposibles', () => {
  const res = validateRows([makeRow({ Fecha: '31/02/2025', Hora: '99:99' })]);
  assert.ok(res.errors.some((e) => e.campo === 'Fecha'));
  assert.ok(res.errors.some((e) => e.campo === 'Hora'));
  assert.equal(res.invalidRows.length, 1);
});

test('validateRows exige complejo y sala', () => {
  const res = validateRows([makeRow({ Complejo: '', Salas: '' })]);
  assert.ok(res.errors.some((e) => e.campo === 'Complejo'));
  assert.ok(res.errors.some((e) => e.campo === 'Sala'));
});

test('validateRows reporta archivo y línea de origen', () => {
  const row = { ...makeRow({ Cantidad: '' }), Linea: 201938, _file: 'cines-unidos.xls' };
  const { errors } = validateRows([row]);
  const err = errors.find((e) => e.campo === 'Cantidad');
  assert.equal(err.archivo, 'cines-unidos.xls');
  assert.equal(err.fila, 201938);
});

test('parseTaquillaFile lee celdas numéricas de xlsx sin corromper montos', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pelven-test-'));
  const file = path.join(dir, 'reporte.xlsx');
  try {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('datos');
    ws.addRow(HEADER);
    ws.addRow(['02/01/2025', 'C.U. EL MARQUES', '1', '1', 'PELICULA', '2D', '15:00', 'Vespertina', 'Completo', 78.49, 78.49, 11.22, 16, 169.99, 10, 1699.9]);
    await wb.xlsx.writeFile(file);

    const parsed = await parseTaquillaFile(file, 'reporte.xlsx');
    assert.equal(parsed.rows.length, 1);
    assert.equal(parsed.rows[0].PrecioBoleto, 78.49);
    assert.equal(parsed.rows[0].Cantidad, 10);
    assert.equal(parsed.rows[0].TotalRecaudacion, 1699.9);
    const { errors } = validateRows(parsed.rows);
    assert.deepEqual(errors, []);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('aggregateBySala y calcPelven usan solo Vespertina+Intermedia+Noche', () => {
  const rows = [makeRow({ Cantidad: '10' }), makeRow({ Cantidad: '20', Función: 'Matinee' })];
  const agg = aggregateBySala(rows);
  const out = calcPelven(agg, new Map(), {});
  const r = out[0];
  assert.equal(r.totalEspectadores, 30);
  assert.equal(r.totalBase, 10);
  assert.equal(r.diasBase, 1);
  assert.equal(r.base, 42);
  assert.equal(r.f21, 42);
  assert.equal(r.f20, 40);
});

test('calcPelven usa la fracción exacta n/21 y no factores pre-redondeados', () => {
  const diasBase = new Set(Array.from({ length: 30 }, (_, i) => `${String(i + 1).padStart(2, '0')}/01/2025`));
  const agg = {
    'C.U. PRUEBA|1': {
      complejo: 'C.U. PRUEBA',
      sala: '1',
      total: 50000,
      totalBase: 50000,
      diasBase,
      dias: new Set(diasBase),
      funciones: new Set(),
      porSemana: {},
    },
  };
  const out = calcPelven(agg, new Map(), { factorDiario: 4.2 });
  assert.equal(out[0].base, 7000);
  // 7000 * 0.9524 = 6666.8 (redondearía a 6667); 7000 * 20/21 = 6666.666… → 6666.
  assert.equal(out[0].f20, 6666);
  assert.equal(out[0].f10, 3333);
  assert.equal(out[0].f1, 333);
});

test('findDuplicateRows detecta filas idénticas y no las marca como error', () => {
  const a = { ...makeRow(), Linea: 10, _file: 'a.xls' };
  const b = { ...makeRow(), Linea: 20, _file: 'b.xls' };
  const c = { ...makeRow({ Cantidad: '6' }), Linea: 21, _file: 'b.xls' };

  const res = findDuplicateRows([a, b, c]);
  assert.equal(res.grupos, 1);
  assert.equal(res.filasExtra, 1);
  assert.equal(res.detalle[0].ocurrencias, 2);
  assert.equal(res.detalle[0].filas[0].archivo, 'a.xls');
  assert.equal(res.detalle[0].filas[1].archivo, 'b.xls');

  const sinDuplicados = findDuplicateRows([a, c]);
  assert.equal(sinDuplicados.grupos, 0);
  assert.equal(sinDuplicados.filasExtra, 0);
});

test('roundAt70 aplica el criterio reglamentario', () => {
  assert.equal(roundAt70(10.69), 10);
  assert.equal(roundAt70(10.7), 11);
  assert.equal(roundAt70(10.699999999), 11); // dentro de la tolerancia de 1e-9
  assert.equal(roundAt70(0), 0);
});

test('semanaCine agrupa por viernes', () => {
  assert.equal(semanaCine('02/01/2025'), '2024-12-27');
  assert.equal(semanaCine('03/01/2025'), '2025-01-03');
  assert.equal(semanaCine(''), '');
});

test('buildWorkbook conserva el Nro oficial de la fila', async () => {
  const row = {
    nro: 7,
    circuito: 'CINES UNIDOS',
    administradora: 'ADMIN',
    complejo: 'C.U. PRUEBA',
    estado: 'MIRANDA',
    ciudad: 'CARACAS',
    sala: '1',
    aforo: 100,
    base: 10,
  };
  for (let n = 20; n >= 1; n--) row[`f${n}`] = n;
  const buf = await buildWorkbook([row], 2026);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  const ws = wb.worksheets[0];
  assert.equal(ws.getCell(8, 1).value, 7);
  assert.equal(ws.getCell(8, 9).value, 10);
});
