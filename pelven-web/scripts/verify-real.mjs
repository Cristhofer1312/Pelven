// Verificación opcional contra el reporte real de Cines Unidos.
// Uso: npm run verify:real [ruta-al-archivo]
// El archivo no está en Git por su tamaño (78 MB); se espera en la raíz del repo.
import fs from 'node:fs';
import path from 'node:path';
import { parseTaquillaFile, validateRows, aggregateBySala, calcPelven } from '../pelven-engine.js';

const defaultFile = path.resolve(process.cwd(), '..', 'COMPAÑIA ANONIMA EMPRESA CINES UNIDOS 20250102-20251231.xls');
const input = process.argv[2] ? path.resolve(process.argv[2]) : defaultFile;

if (!fs.existsSync(input)) {
  console.error(`No se encontró el archivo: ${input}`);
  console.error('Pásalo como argumento: npm run verify:real -- "ruta/al/reporte.xls"');
  process.exit(2);
}

const norm = (s) => String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();

const t0 = Date.now();
const parsed = await parseTaquillaFile(input, path.basename(input));
const parseMs = Date.now() - t0;
const { errors, warnings, invalidRows } = validateRows(parsed.rows);

console.log(`Archivo: ${path.basename(input)}`);
console.log(`Parseo: ${parsed.rows.length.toLocaleString()} filas en ${(parseMs / 1000).toFixed(1)} s`);
console.log(`Rango: ${parsed.rango.join(' al ')} · ${parsed.semanas.size} semanas`);
console.log(`Validación: ${errors.length} errores · ${warnings.length} avisos · ${invalidRows.length} filas inválidas`);

if (errors.length) {
  console.error('\nPrimeros errores:');
  for (const e of errors.slice(0, 10)) console.error(`  ${e.archivo || 'archivo'}:${e.fila} ${e.campo} = ${JSON.stringify(e.valor)}`);
  process.exit(1);
}

const agg = aggregateBySala(parsed.rows);
const maestro = JSON.parse(fs.readFileSync(path.resolve(process.cwd(), 'maestro-seed.json'), 'utf-8'));
const map = new Map(maestro.map((r) => [norm(r.complejo) + '|' + String(r.sala), r]));
const out = calcPelven(agg, map, {});

let compared = 0;
let sumPct = 0;
let maxPct = 0;
let maxItem = null;
for (const r of out) {
  const m = map.get(norm(r.complejo) + '|' + String(r.sala));
  if (!m || m.base2026 == null || !r.rawBase) continue;
  const pct = (Math.abs(r.rawBase - m.base2026) / m.base2026) * 100;
  compared++;
  sumPct += pct;
  if (pct > maxPct) {
    maxPct = pct;
    maxItem = { complejo: r.complejo, sala: r.sala, calculado: r.rawBase, oficial: m.base2026 };
  }
}

console.log(`Salas: ${Object.keys(agg).length} detectadas · ${out.length} filas de salida`);
if (compared) {
  console.log(`Base vs maestro: error medio ${(sumPct / compared).toFixed(3)} % · máximo ${maxPct.toFixed(3)} %`);
  if (maxItem) console.log(`  mayor desviación: ${maxItem.complejo} sala ${maxItem.sala} (calc ${maxItem.calculado.toFixed(2)} vs oficial ${maxItem.oficial.toFixed(2)})`);
}
console.log('\nVerificación OK');
