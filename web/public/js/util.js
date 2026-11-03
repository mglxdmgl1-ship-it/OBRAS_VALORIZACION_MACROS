// Utilidades generales (sin DOM): redondeo como Excel, fechas ISO, formatos en español (Perú).

/** Redondeo como ROUND() de Excel (mitad lejos de cero, sobre 15 dígitos significativos). */
export function redondear(x, d = 2) {
  if (x === null || x === undefined || x === '' || !isFinite(x)) return 0;
  const n = Number(Number(x).toPrecision(15));
  const s = Math.sign(n) || 1;
  return s * Number(Math.round(Number(Math.abs(n) + 'e' + d)) + 'e-' + d);
}
export const r2 = (x) => redondear(x, 2);
export const r4 = (x) => redondear(x, 4);

/** Convierte texto o número a número (acepta "1,234.50", "1 234,50", "12.5%"). */
export function num(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  let s = String(v).trim().replace(/\s|S\/\.?/g, '');
  if (!s) return 0;
  const pct = s.endsWith('%');
  if (pct) s = s.slice(0, -1);
  if (/,\d{1,3}$/.test(s) && s.indexOf('.') >= 0 && s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g, '').replace(',', '.');
  else if (/^-?\d+,\d+$/.test(s)) s = s.replace(',', '.');
  else s = s.replace(/,/g, '');
  const n = parseFloat(s);
  if (!isFinite(n)) return 0;
  return pct ? n / 100 : n;
}

export function esNumero(v) {
  if (typeof v === 'number') return isFinite(v);
  if (typeof v !== 'string') return false;
  return /^\s*-?[\d.,]+\s*%?\s*$/.test(v) && /\d/.test(v);
}

// ------------------------------------------------------------------ formatos
const nf = {};
function fmtN(d) {
  if (!nf[d]) nf[d] = new Intl.NumberFormat('en-US', { minimumFractionDigits: d, maximumFractionDigits: d });
  return nf[d];
}
/** 1234.5 → "1,234.50" */
export function fmt(v, d = 2, vacioSiCero = false) {
  if (v === null || v === undefined || v === '' || !isFinite(v)) return '';
  if (vacioSiCero && Math.abs(v) < 1e-12) return '';
  const x = redondear(v, d);
  return fmtN(d).format(Object.is(x, -0) ? 0 : x);
}
export const soles = (v) => (v === null || v === undefined ? '' : 'S/ ' + fmt(v, 2));
export const pct = (v, d = 2) => (v === null || v === undefined || !isFinite(v) ? '' : fmt(v * 100, d) + ' %');
/** Cantidad con decimales "inteligentes": hasta 4, sin ceros sobrantes. */
export function cant(v, max = 4) {
  if (v === null || v === undefined || v === '' || !isFinite(v)) return '';
  const x = redondear(v, max);
  return new Intl.NumberFormat('en-US', { minimumFractionDigits: 0, maximumFractionDigits: max }).format(x);
}

// ------------------------------------------------------------------ fechas ISO (AAAA-MM-DD, sin zona horaria)
export const MESES = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SETIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
export const MESES_C = ['ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO', 'SET', 'OCT', 'NOV', 'DIC'];
export const DIAS = ['DOMINGO', 'LUNES', 'MARTES', 'MIÉRCOLES', 'JUEVES', 'VIERNES', 'SÁBADO'];

export function toDate(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d));
}
export function iso(dt) {
  return dt.toISOString().slice(0, 10);
}
export function hoy() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}
export function sumarDias(f, n) {
  const d = toDate(f);
  d.setUTCDate(d.getUTCDate() + n);
  return iso(d);
}
export function difDias(a, b) {
  return Math.round((toDate(b) - toDate(a)) / 86400000);
}
export function diaSemana(f) {
  return toDate(f).getUTCDay();
}
export function mesDe(f) {
  return f.slice(0, 7);
}
export function diasDelMes(ym) {
  const [y, m] = ym.split('-').map(Number);
  return new Date(Date.UTC(y, m, 0)).getUTCDate();
}
export function finDeMes(ym) {
  return `${ym}-${String(diasDelMes(ym)).padStart(2, '0')}`;
}
export function mesSiguiente(ym) {
  let [y, m] = ym.split('-').map(Number);
  m++;
  if (m > 12) { m = 1; y++; }
  return `${y}-${String(m).padStart(2, '0')}`;
}
export function mesesEntre(a, b) {
  const out = [];
  let ym = mesDe(a);
  const fin = mesDe(b);
  while (ym <= fin && out.length < 240) { out.push(ym); ym = mesSiguiente(ym); }
  return out;
}
export function nombreMes(ym, corto = false) {
  const [y, m] = ym.split('-').map(Number);
  return (corto ? MESES_C : MESES)[m - 1] + ' ' + y;
}
/** "2026-10-15" → "15/10/2026" */
export function fFecha(f) {
  if (!f) return '';
  const [y, m, d] = String(f).slice(0, 10).split('-');
  return d && m && y ? `${d}/${m}/${y}` : String(f);
}
/** Acepta "15/10/2026", "2026-10-15", número de serie de Excel o Date → ISO. */
export function aISO(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) return isNaN(v) ? null : iso(new Date(Date.UTC(v.getUTCFullYear(), v.getUTCMonth(), v.getUTCDate())));
  if (typeof v === 'number' && v > 20000 && v < 80000) return sumarDias('1899-12-30', Math.floor(v));
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  m = s.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})/);
  if (m) {
    const y = m[3].length === 2 ? '20' + m[3] : m[3];
    return `${y}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  }
  return null;
}

// ------------------------------------------------------------------ texto
export function esc(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
export function normal(s) {
  return String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/\s+/g, ' ').trim();
}
export function clave(s) {
  return normal(s).replace(/[^A-Z0-9]/g, '');
}
/** Normaliza un código de ítem: "01.01.01", quita espacios y punto final. */
export function normItem(s) {
  return String(s ?? '').trim().replace(/\s+/g, '').replace(/\.+$/, '');
}
export function nivelItem(item) {
  return item ? item.split('.').length : 0;
}
/** Comparación natural de ítems ("01.2" < "01.10"). */
export function cmpItem(a, b) {
  const pa = a.split('.'), pb = b.split('.');
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    if (pa[i] === undefined) return -1;
    if (pb[i] === undefined) return 1;
    const na = parseInt(pa[i], 10), nb = parseInt(pb[i], 10);
    if (!isNaN(na) && !isNaN(nb) && na !== nb) return na - nb;
    if (pa[i] !== pb[i]) return pa[i] < pb[i] ? -1 : 1;
  }
  return 0;
}
export function uid(p = '') {
  return p + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}
export function clonar(o) {
  return typeof structuredClone === 'function' ? structuredClone(o) : JSON.parse(JSON.stringify(o));
}

/** Aplica la regla de redondeo para el almacén. */
export function redondearCantidad(x, modo = 'arriba', paso = 1) {
  if (!isFinite(x)) return 0;
  const p = paso > 0 ? paso : 1;
  const v = Number((x / p).toPrecision(12));
  switch (modo) {
    case 'arriba': return redondear(Math.ceil(v - 1e-9) * p, 4);
    case 'normal': return redondear(Math.round(v) * p, 4);
    case 'abajo': return redondear(Math.floor(v + 1e-9) * p, 4);
    case 'dos': return redondear(x, 2);
    default: return redondear(x, 4);
  }
}
