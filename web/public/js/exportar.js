// Exportación genérica de reportes a Excel (con fórmulas, ExcelJS) y a PDF (jsPDF + AutoTable).
//
// Un reporte es:
// { archivo, titulo, subtitulo, codigo, orientacion:'l'|'p', firmas:true, notas:[],
//   secciones: [{ titulo, hoja, columnas:[{id,t,w,tipo,g}], filas:[{v:{},est,nivel,hijos:[i],fx:{col:fn}}],
//                 fx:{col:(r,L,ctx)=>'FORMULA'}, sumar:[cols] }] }
// tipos: txt, n2, n4, s (soles), pct, int, fecha.  est: titulo | total | sub | alerta | ok
import { fmt, cant, fFecha, toDate, r2, hoy } from './util.js';
import { descargar, nombreArchivo } from './ui.js';
import { S } from './store.js';

const AZUL = '1F3A5F', AZUL2 = 'DCE6F1', GRIS = 'F2F2F2', AMAR = 'FFF2CC';

function cabeceraObra() {
  const d = S.obra?.datos || {};
  return {
    entidad: [d.entidad, d.dependencia].filter(Boolean).join(' – ') || 'ENTIDAD',
    proyecto: d.nombre || '',
    linea: [d.cui && 'CUI: ' + d.cui, d.meta && 'META: ' + d.meta, [d.distrito, d.provincia, d.region].filter(Boolean).join(' – '), d.modalidad].filter(Boolean).join('   ·   '),
    firmas: (S.obra?.firmas || []).filter(Boolean),
    responsables: S.obra?.responsables || [],
  };
}
function nombreFirma(cargo) {
  const r = (S.obra?.responsables || []).find((x) => String(x.cargo || '').toUpperCase() === String(cargo).toUpperCase());
  return r ? r.nombre : '';
}

const FMT_XL = { n2: '#,##0.00;[Red]-#,##0.00', n4: '#,##0.0000', s: '#,##0.00;[Red]-#,##0.00', pct: '0.00%', int: '#,##0', fecha: 'dd/mm/yyyy' };
function valorTexto(v, tipo) {
  if (v === null || v === undefined || v === '') return '';
  switch (tipo) {
    case 'n2': case 's': return typeof v === 'number' ? fmt(v, 2) : String(v);
    case 'n4': return typeof v === 'number' ? cant(v, 4) : String(v);
    case 'pct': return typeof v === 'number' ? fmt(v * 100, 2) + '%' : String(v);
    case 'int': return typeof v === 'number' ? fmt(v, 0) : String(v);
    case 'fecha': return fFecha(v);
    default: return String(v);
  }
}
const colLetra = (n) => { let s = ''; n++; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

// =================================================================== EXCEL
export async function aExcel(R) {
  if (!window.ExcelJS) throw new Error('No se cargó la librería de Excel.');
  const wb = new window.ExcelJS.Workbook();
  wb.creator = 'SGO – Costos, presupuestos y valorizaciones';
  wb.created = new Date();
  const cab = cabeceraObra();
  const hojas = new Map();
  // ---------- primera pasada: posición de cada fila
  const layout = [];
  for (const [si, sec] of R.secciones.entries()) {
    const nombreHoja = (sec.hoja || R.hoja || R.codigo || 'Reporte').replace(/[\\/*?:[\]]/g, ' ').slice(0, 31);
    if (!hojas.has(nombreHoja)) hojas.set(nombreHoja, { fila: 1, ws: null, primera: true, ncol: 0 });
    const h = hojas.get(nombreHoja);
    if (h.primera) { h.fila = 7; h.primera = false; } else h.fila += 1;
    const conGrupo = sec.columnas.some((c) => c.g);
    const tituloFila = sec.titulo ? h.fila++ : null;
    const filaCab = h.fila;
    h.fila += conGrupo ? 2 : 1;
    const filas = sec.filas.map(() => h.fila++);
    h.ncol = Math.max(h.ncol, sec.columnas.length);
    layout.push({ si, hoja: nombreHoja, tituloFila, filaCab, conGrupo, filas });
  }
  const ctx = {
    ref: (si, fi, col, abs = true) => {
      const L = layout[si], sec = R.secciones[si];
      const c = sec.columnas.findIndex((x) => x.id === col);
      const d = abs ? '$' : '';
      return `'${L.hoja}'!${d}${colLetra(c)}${d}${L.filas[fi]}`;
    },
  };
  // ---------- hojas con encabezado de la obra
  for (const [nombre, h] of hojas) {
    const ws = wb.addWorksheet(nombre, {
      pageSetup: { paperSize: 9, orientation: R.orientacion === 'p' ? 'portrait' : 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.6, header: 0.3, footer: 0.3 } },
      views: [{ showGridLines: false }],
    });
    ws.headerFooter.oddFooter = '&L&8' + (R.titulo || '') + '&R&8Página &P de &N';
    h.ws = ws;
    const n = Math.max(h.ncol, 6);
    const ult = colLetra(n - 1);
    const put = (fila, txt, est) => {
      ws.mergeCells(`A${fila}:${ult}${fila}`);
      const c = ws.getCell(`A${fila}`);
      c.value = txt;
      Object.assign(c, est);
    };
    put(1, cab.entidad, { font: { bold: true, size: 10, color: { argb: 'FF' + AZUL } } });
    put(2, cab.proyecto, { font: { bold: true, size: 9 }, alignment: { wrapText: true, vertical: 'top' } });
    ws.getRow(2).height = 26;
    put(3, cab.linea, { font: { size: 8, color: { argb: 'FF555555' } } });
    put(4, (R.codigo ? R.codigo + '  ' : '') + (R.titulo || ''), { font: { bold: true, size: 13, color: { argb: 'FF' + AZUL } }, alignment: { horizontal: 'center' } });
    put(5, R.subtitulo || '', { font: { size: 9, italic: true }, alignment: { horizontal: 'center' } });
  }
  // ---------- secciones
  for (const [si, sec] of R.secciones.entries()) {
    const L = layout[si];
    const ws = hojas.get(L.hoja).ws;
    const cols = sec.columnas;
    const Lt = {};
    cols.forEach((c, i) => (Lt[c.id] = colLetra(i)));
    cols.forEach((c, i) => {
      const col = ws.getColumn(i + 1);
      col.width = Math.max(col.width || 0, c.w || (c.tipo === 'txt' ? 14 : 12));
    });
    if (L.tituloFila) {
      const c = ws.getCell(`A${L.tituloFila}`);
      c.value = sec.titulo;
      c.font = { bold: true, size: 10, color: { argb: 'FF' + AZUL } };
    }
    const estiloCab = (c) => {
      c.font = { bold: true, size: 8, color: { argb: 'FFFFFFFF' } };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + AZUL } };
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      c.border = borde('FFFFFFFF');
    };
    if (L.conGrupo) {
      const r1 = L.filaCab, r2_ = L.filaCab + 1;
      let i = 0;
      while (i < cols.length) {
        const g = cols[i].g;
        if (!g) {
          ws.mergeCells(`${colLetra(i)}${r1}:${colLetra(i)}${r2_}`);
          ws.getCell(`${colLetra(i)}${r1}`).value = cols[i].t;
          estiloCab(ws.getCell(`${colLetra(i)}${r1}`));
          i++;
        } else {
          let j = i;
          while (j + 1 < cols.length && cols[j + 1].g === g) j++;
          if (j > i) ws.mergeCells(`${colLetra(i)}${r1}:${colLetra(j)}${r1}`);
          ws.getCell(`${colLetra(i)}${r1}`).value = g;
          estiloCab(ws.getCell(`${colLetra(i)}${r1}`));
          for (let k = i; k <= j; k++) {
            const c = ws.getCell(`${colLetra(k)}${r2_}`);
            c.value = cols[k].t;
            estiloCab(c);
          }
          i = j + 1;
        }
      }
      ws.getRow(r1).height = 18;
      ws.getRow(r2_).height = 30;
    } else {
      cols.forEach((c, i) => {
        const cell = ws.getCell(`${colLetra(i)}${L.filaCab}`);
        cell.value = c.t;
        estiloCab(cell);
      });
      ws.getRow(L.filaCab).height = 32;
    }
    if (si === 0 && !sec.sinCongelar) ws.views = [{ state: 'frozen', ySplit: L.filaCab + (L.conGrupo ? 1 : 0), showGridLines: false }];
    // filas
    sec.filas.forEach((f, fi) => {
      const r = L.filas[fi];
      const row = ws.getRow(r);
      cols.forEach((c, ci) => {
        const cell = row.getCell(ci + 1);
        let v = f.v[c.id];
        let fx = (f.fx && f.fx[c.id]) || (!f.est && sec.fx && sec.fx[c.id]) || null;
        if (!fx && f.hijos && f.hijos.length && sec.sumar && sec.sumar.includes(c.id)) {
          const filasH = f.hijos.map((h) => L.filas[h]);
          fx = () => sumaFilas(Lt[c.id], filasH);
        }
        if (fx && f.sinFx !== true) {
          const formula = fx(r, Lt, ctx, f);
          if (formula) cell.value = { formula, result: typeof v === 'number' ? v : undefined };
          else cell.value = normalizarValor(v, c.tipo);
        } else cell.value = normalizarValor(v, c.tipo);
        if (FMT_XL[c.tipo]) cell.numFmt = FMT_XL[c.tipo];
        cell.font = { size: 8, bold: f.est === 'titulo' || f.est === 'total' || f.est === 'sub' };
        cell.alignment = { vertical: 'top', wrapText: c.tipo === 'txt' && (c.w || 0) >= 30, horizontal: c.tipo === 'txt' ? 'left' : 'right', indent: c.sangria && f.nivel ? Math.min(f.nivel - 1, 6) : 0 };
        cell.border = borde('FFBFBFBF');
        const fill = f.est === 'total' ? AZUL2 : f.est === 'titulo' ? GRIS : f.est === 'sub' ? 'EAF1FB' : f.est === 'alerta' ? AMAR : null;
        if (fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + fill } };
        if (f.color && f.color[c.id]) cell.font = { ...cell.font, color: { argb: 'FF' + f.color[c.id] } };
      });
    });
    if (sec.editables) {
      for (const id of sec.editables) {
        const ci = cols.findIndex((c) => c.id === id);
        if (ci < 0) continue;
        sec.filas.forEach((f, fi) => {
          if (f.est) return;
          ws.getCell(`${colLetra(ci)}${L.filas[fi]}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9DB' } };
        });
      }
    }
  }
  // ---------- notas y firmas
  for (const [, h] of hojas) {
    const ws = h.ws;
    let fila = ws.rowCount + 2;
    for (const n of R.notas || []) {
      ws.getCell(`A${fila}`).value = n;
      ws.getCell(`A${fila}`).font = { size: 8, italic: true, color: { argb: 'FF555555' } };
      fila++;
    }
    if (R.firmas !== false && cab.firmas.length) {
      fila += 3;
      const n = Math.max(h.ncol, 6);
      const paso = Math.floor(n / cab.firmas.length);
      cab.firmas.forEach((cargo, i) => {
        const c0 = i * paso, c1 = Math.min(n - 1, c0 + Math.max(paso - 1, 0));
        const a = `${colLetra(c0)}${fila}`, b = `${colLetra(c1)}${fila}`;
        if (c1 > c0) { ws.mergeCells(`${a}:${b}`); ws.mergeCells(`${colLetra(c0)}${fila + 1}:${colLetra(c1)}${fila + 1}`); ws.mergeCells(`${colLetra(c0)}${fila + 2}:${colLetra(c1)}${fila + 2}`); }
        const c = ws.getCell(a);
        c.value = '______________________________';
        c.alignment = { horizontal: 'center' };
        const nm = ws.getCell(`${colLetra(c0)}${fila + 1}`);
        nm.value = nombreFirma(cargo);
        nm.alignment = { horizontal: 'center' };
        nm.font = { size: 8 };
        const cg = ws.getCell(`${colLetra(c0)}${fila + 2}`);
        cg.value = cargo;
        cg.alignment = { horizontal: 'center' };
        cg.font = { bold: true, size: 8 };
      });
    }
  }
  const buf = await wb.xlsx.writeBuffer();
  descargar(nombreArchivo((R.archivo || R.titulo || 'reporte') + '_' + hoy()) + '.xlsx', new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
}
/** SUM de varias filas de una columna: agrupa filas consecutivas en rangos y respeta el límite de 255 argumentos. */
export function sumaFilas(col, filas) {
  const fs = [...new Set(filas)].sort((a, b) => a - b);
  const partes = [];
  for (let i = 0; i < fs.length; i++) {
    let j = i;
    while (j + 1 < fs.length && fs[j + 1] === fs[j] + 1) j++;
    partes.push(j > i ? `${col}${fs[i]}:${col}${fs[j]}` : `${col}${fs[i]}`);
    i = j;
  }
  if (!partes.length) return '0';
  const grupos = [];
  for (let i = 0; i < partes.length; i += 200) grupos.push(`SUM(${partes.slice(i, i + 200).join(',')})`);
  return grupos.length === 1 ? grupos[0] : `SUM(${grupos.join(',')})`;
}
function borde(argb) {
  const s = { style: 'thin', color: { argb } };
  return { top: s, left: s, bottom: s, right: s };
}
function normalizarValor(v, tipo) {
  if (v === undefined || v === null) return null;
  if (tipo === 'fecha' && typeof v === 'string' && /^\d{4}-\d{2}-\d{2}/.test(v)) return toDate(v.slice(0, 10));
  if (typeof v === 'number') return isFinite(v) ? v : null;
  return v;
}

// =================================================================== PDF
export function aPDF(R) {
  const J = window.jspdf?.jsPDF;
  if (!J) throw new Error('No se cargó la librería de PDF.');
  const doc = new J({ orientation: R.orientacion === 'p' ? 'portrait' : 'landscape', unit: 'mm', format: R.papel || 'a4' });
  if (typeof doc.autoTable !== 'function') throw new Error('No se cargó el complemento de tablas del PDF.');
  const W = doc.internal.pageSize.getWidth(), H = doc.internal.pageSize.getHeight();
  const mx = 10;
  const cab = cabeceraObra();
  const tope = 33;
  const encabezado = () => {
    doc.setDrawColor(31, 58, 95);
    doc.setFillColor(31, 58, 95);
    doc.rect(mx, 7, W - 2 * mx, 1.2, 'F');
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8);
    doc.setTextColor(31, 58, 95);
    doc.text(cab.entidad.slice(0, 150), mx, 12);
    doc.setTextColor(40);
    doc.setFontSize(7);
    const pr = doc.splitTextToSize(cab.proyecto, W - 2 * mx);
    doc.text(pr.slice(0, 2), mx, 15.5);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(90);
    doc.text(cab.linea.slice(0, 200), mx, 15.5 + Math.min(pr.length, 2) * 3);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(11);
    doc.setTextColor(31, 58, 95);
    doc.text(((R.codigo ? R.codigo + '  ' : '') + (R.titulo || '')).slice(0, 120), W / 2, 26, { align: 'center' });
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    doc.setTextColor(60);
    if (R.subtitulo) doc.text(String(R.subtitulo).slice(0, 180), W / 2, 30, { align: 'center' });
  };
  let y = tope;
  for (const sec of R.secciones) {
    const cols = sec.columnas;
    const conGrupo = cols.some((c) => c.g);
    const head = [];
    if (conGrupo) {
      const h1 = [], h2 = [];
      let i = 0;
      while (i < cols.length) {
        if (!cols[i].g) { h1.push({ content: cols[i].t, rowSpan: 2 }); i++; continue; }
        let j = i;
        while (j + 1 < cols.length && cols[j + 1].g === cols[i].g) j++;
        h1.push({ content: cols[i].g, colSpan: j - i + 1 });
        for (let k = i; k <= j; k++) h2.push(cols[k].t);
        i = j + 1;
      }
      head.push(h1, h2);
    } else head.push(cols.map((c) => c.t));
    const body = sec.filas.map((f) => cols.map((c, ci) => {
      let t = valorTexto(f.v[c.id], c.tipo);
      if (ci === 0 || c.sangria) { /* sin cambios */ }
      if (c.sangria && f.nivel > 1) t = '  '.repeat(Math.min(f.nivel - 1, 6)) + t;
      return t;
    }));
    const total = cols.reduce((s, c) => s + (c.w || 12), 0);
    const util = W - 2 * mx;
    const columnStyles = {};
    cols.forEach((c, i) => {
      columnStyles[i] = { cellWidth: ((c.w || 12) / total) * util, halign: c.tipo === 'txt' ? 'left' : 'right' };
    });
    if (sec.titulo) {
      if (y > H - 30) { doc.addPage(); y = tope; }
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(9);
      doc.setTextColor(31, 58, 95);
      doc.text(sec.titulo, mx, y + 3);
      y += 5;
    }
    const fuente = sec.fuente || (cols.length > 14 ? 5.5 : cols.length > 10 ? 6.3 : 7.2);
    doc.autoTable({
      head, body, startY: y, margin: { left: mx, right: mx, top: tope, bottom: 14 },
      theme: 'grid',
      styles: { fontSize: fuente, cellPadding: 0.9, overflow: 'linebreak', lineColor: [190, 190, 190], lineWidth: 0.1, textColor: 30 },
      headStyles: { fillColor: [31, 58, 95], textColor: 255, fontStyle: 'bold', halign: 'center', valign: 'middle', fontSize: fuente },
      columnStyles,
      didParseCell: (d) => {
        if (d.section !== 'body') return;
        const f = sec.filas[d.row.index];
        if (!f) return;
        if (f.est === 'titulo') { d.cell.styles.fontStyle = 'bold'; d.cell.styles.fillColor = [242, 242, 242]; }
        if (f.est === 'total') { d.cell.styles.fontStyle = 'bold'; d.cell.styles.fillColor = [220, 230, 241]; }
        if (f.est === 'sub') { d.cell.styles.fontStyle = 'bold'; d.cell.styles.fillColor = [234, 241, 251]; }
        if (f.est === 'alerta') d.cell.styles.fillColor = [255, 242, 204];
        const col = cols[d.column.index];
        if (f.color && f.color[col.id]) { const h = f.color[col.id]; d.cell.styles.textColor = [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)]; }
        if (f.fondo && f.fondo[col.id]) { const h = f.fondo[col.id]; d.cell.styles.fillColor = [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)]; }
      },
      didDrawPage: () => encabezado(),
    });
    y = doc.lastAutoTable.finalY + 6;
  }
  if (R.notas?.length) {
    doc.setFont('helvetica', 'italic');
    doc.setFontSize(7);
    doc.setTextColor(80);
    for (const n of R.notas) {
      const ls = doc.splitTextToSize(n, W - 2 * mx);
      if (y + ls.length * 3 > H - 14) { doc.addPage(); encabezado(); y = tope; }
      doc.text(ls, mx, y);
      y += ls.length * 3 + 1;
    }
  }
  if (R.firmas !== false && cab.firmas.length) {
    if (y + 28 > H - 12) { doc.addPage(); encabezado(); y = tope; }
    y += 18;
    const n = cab.firmas.length;
    const ancho = (W - 2 * mx) / n;
    doc.setTextColor(30);
    cab.firmas.forEach((cargo, i) => {
      const cx = mx + ancho * i + ancho / 2;
      doc.setDrawColor(60);
      doc.line(cx - 30, y, cx + 30, y);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(7);
      const nm = nombreFirma(cargo);
      if (nm) doc.text(nm, cx, y + 3.5, { align: 'center' });
      doc.setFont('helvetica', 'bold');
      doc.text(cargo, cx, y + (nm ? 7 : 3.5), { align: 'center' });
    });
  }
  const np = doc.internal.getNumberOfPages();
  for (let i = 1; i <= np; i++) {
    doc.setPage(i);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7);
    doc.setTextColor(120);
    doc.text(`Página ${i} de ${np}`, W - mx, H - 6, { align: 'right' });
    doc.text(`Generado el ${fFecha(hoy())} · SGO Costos y Valorizaciones`, mx, H - 6);
  }
  doc.save(nombreArchivo((R.archivo || R.titulo || 'reporte') + '_' + hoy()) + '.pdf');
}

export async function exportar(R, formato) {
  if (formato === 'pdf') return aPDF(R);
  return aExcel(R);
}
export { r2 };
