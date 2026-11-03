// Equipo mecánico: equipos y O/S, partes diarios, partes desde A.C.U. × metrado y control 4.15.
import { S, cambio } from '../store.js';
import { esc, fmt, cant, num, hoy, mesDe, nombreMes, fFecha, uid, diasDelMes } from '../util.js';
import { icono, formulario, confirmar, aviso } from '../ui.js';
import { rControlEquipos, rPartes } from '../reportes.js';
import { cab, btnExport, enlazarExport, mesesDeObra, listaPartidasDatalist, datalistInsumos } from './comun.js';
import { refrescar } from '../app.js';

let tab = 'equipos', cMes = null, gen = { desde: null, hasta: null, props: null, key: '' }, fMes = '';

export function render(cont, params) {
  if (params.tab) tab = params.tab;
  if (params.desde) { gen = { desde: params.desde, hasta: params.hasta || params.desde, props: null, key: '' }; tab = 'generar'; }
  const o = S.obra;
  cont.innerHTML = `
    ${cab('Equipo mecánico', `${o.equipos.length} equipo(s) · ${o.partes.length} parte(s) diario(s). Horas a valorizar = horómetro (si se registra) o el horario de mañana y tarde; también puede digitarlas.`,
      `<button class="btn btn-primario" id="b-parte">${icono('mas')}Parte diario</button><button class="btn" id="b-eq">${icono('mas')}Equipo</button>`)}
    <div class="pestanas">${[['equipos', 'Equipos y O/S'], ['partes', 'Partes diarios'], ['generar', 'Partes desde metrados (A.C.U.)'], ['control', 'Control 4.15']].map(([k, t]) => `<button data-tab="${k}" class="${tab === k ? 'activo' : ''}">${t}</button>`).join('')}</div>
    <div id="cuerpo"></div>${listaPartidasDatalist()}${datalistInsumos('dl-eq', ['EQ', 'SC'])}`;
  cont.querySelectorAll('[data-tab]').forEach((b) => (b.onclick = () => { tab = b.dataset.tab; refrescar(); }));
  cont.querySelector('#b-parte').onclick = () => editarParte(null);
  cont.querySelector('#b-eq').onclick = () => editarEquipo(null);
  ({ equipos, partes, generar, control })[tab](cont.querySelector('#cuerpo'));
}

const CAMPOS_EQ = [
  { id: 'codigo', et: 'Código', req: true, ph: 'E-01' }, { id: 'descripcion', et: 'Descripción', req: true, col: 2 },
  { id: 'marca', et: 'Marca' }, { id: 'modelo', et: 'Modelo' }, { id: 'serie', et: 'Serie / placa' },
  { id: 'proveedor', et: 'Proveedor', col: 2 }, { id: 'os', et: 'O/S (N° y fecha)' },
  { id: 'unidad', et: 'Unidad', tipo: 'select', ops: ['H.M.', 'H.E.', 'DÍA', 'MES', 'VIAJE'] }, { id: 'cantidad', et: 'Cantidad contratada', tipo: 'number' }, { id: 'costo', et: 'Costo unitario (S/)', tipo: 'number' },
  { id: 'inicio', et: 'Inicio de la O/S (en obra)', tipo: 'date', ayuda: 'Habilita la valorización de las partidas que usan el equipo' }, { id: 'fin_os', et: 'Fin de la O/S', tipo: 'date' },
  { id: 'estado', et: 'Estado de la O/S', tipo: 'select', ops: ['CONFIRMADO', 'POR CONFIRMAR'] },
  { id: 'condicion', et: 'Condición', ph: 'MÁQUINA SECA / CON COMBUSTIBLE / PROPIO', col: 2 }, { id: 'operador', et: 'Operador' },
  { id: 'insumo_acu', et: 'Insumo del A.C.U. que corresponde', lista: 'dl-eq', col: 3, ayuda: 'Para el control NO VALORIZAR y para generar partes desde los metrados' },
];
function editarEquipo(i) {
  const o = S.obra;
  const e = i !== null ? o.equipos[i] : null;
  formulario({
    titulo: e ? 'Editar equipo ' + e.codigo : 'Nuevo equipo', campos: CAMPOS_EQ, val: e || { codigo: 'E-' + String(o.equipos.length + 1).padStart(2, '0'), unidad: 'H.M.', estado: 'CONFIRMADO' }, ancho: '860px',
    alGuardar: (v) => {
      if (o.equipos.some((x, k) => x.codigo === v.codigo && k !== i)) throw new Error('Código repetido.');
      const ins = v.insumo_acu ? S.motor.insumo(v.insumo_acu) : null;
      if (ins) v.insumo_acu = ins.descripcion;
      if (e) { if (e.codigo !== v.codigo) o.partes.forEach((p) => { if (p.equipo === e.codigo) p.equipo = v.codigo; }); Object.assign(e, v); }
      else o.equipos.push(v);
      cambio(); aviso('Equipo guardado.'); refrescar();
    },
  });
}
function equipos(box) {
  const o = S.obra, m = S.motor;
  const ym = mesDe(hoy());
  const ctl = new Map(m.controlEquipos(ym).map((x) => [x.codigo, x]));
  box.innerHTML = `<div class="tabla-cont"><table class="t"><thead><tr><th>Código</th><th>Equipo</th><th>Proveedor / O/S</th><th>Und</th><th class="n">Contratado</th><th class="n">Costo U.</th><th class="n">Monto O/S</th><th>Inicio – fin O/S</th><th class="n">Acumulado</th><th class="n">Saldo O/S</th><th>Insumo del A.C.U.</th><th></th></tr></thead><tbody>
    ${o.equipos.map((e, i) => { const c = ctl.get(e.codigo); return `<tr><td class="mono">${esc(e.codigo)}</td><td><b>${esc(e.descripcion)}</b><div class="tenue" style="font-size:11px">${esc([e.marca, e.modelo, e.serie].filter(Boolean).join(' · '))}</div></td><td style="font-size:12px">${esc(e.proveedor || '')}<div class="tenue">${esc(e.os || '')}</div></td><td>${esc(e.unidad)}</td>
      <td class="n">${cant(e.cantidad, 2)}</td><td class="n">${fmt(e.costo)}</td><td class="n">${fmt((Number(e.cantidad) || 0) * (Number(e.costo) || 0))}</td><td style="font-size:12px">${fFecha(e.inicio || e.os_fecha)} – ${fFecha(e.fin_os)} ${e.estado === 'POR CONFIRMAR' ? '<span class="badge badge-alerta">por confirmar</span>' : ''}</td>
      <td class="n">${cant(c?.acumulado, 2)}</td><td class="n ${c?.excede ? 'neg' : ''}">${cant(c?.saldo, 2)}</td><td style="font-size:12px">${esc(e.insumo_acu || '')}${e.insumo_acu && !m.insumo(e.insumo_acu) ? ' <span class="badge badge-error">no existe</span>' : ''}</td>
      <td class="acc"><button class="btn-icono" data-ed="${i}">${icono('editar')}</button><button class="btn-icono peligro" data-del="${i}">${icono('borrar')}</button></td></tr>`; }).join('') || '<tr><td colspan="12" class="vacio">Sin equipos registrados.</td></tr>'}
  </tbody></table></div>`;
  box.querySelectorAll('[data-ed]').forEach((b) => (b.onclick = () => editarEquipo(Number(b.dataset.ed))));
  box.querySelectorAll('[data-del]').forEach((b) => (b.onclick = async () => {
    const e = o.equipos[Number(b.dataset.del)];
    const n = o.partes.filter((p) => p.equipo === e.codigo).length;
    if (!(await confirmar(`¿Eliminar el equipo ${esc(e.codigo)}?${n ? ` También se eliminarán sus ${n} parte(s).` : ''}`, { peligro: true }))) return;
    o.partes = o.partes.filter((p) => p.equipo !== e.codigo);
    o.equipos.splice(Number(b.dataset.del), 1);
    cambio(); refrescar();
  }));
}

function editarParte(id) {
  const o = S.obra;
  if (!o.equipos.length) { aviso('Primero registre el equipo.', 'alerta'); return; }
  const p = id ? o.partes.find((x) => x.id === id) : null;
  const ult = o.partes.slice().sort((a, b) => (a.fecha < b.fecha ? 1 : -1))[0];
  const campos = [
    { id: 'fecha', et: 'Fecha', tipo: 'date', req: true }, { id: 'equipo', et: 'Equipo', tipo: 'select', ops: o.equipos.map((e) => [e.codigo, e.codigo + ' · ' + e.descripcion]) }, { id: 'n_parte', et: 'N° de parte' },
    { id: 'operador', et: 'Operador', col: 3 },
    { id: 'm_de', et: 'Mañana de (hh:mm)', ph: '07:00' }, { id: 'm_a', et: 'Mañana a', ph: '12:00' }, { id: 'h_ini', et: 'Horómetro inicial', tipo: 'number' },
    { id: 't_de', et: 'Tarde de', ph: '13:00' }, { id: 't_a', et: 'Tarde a', ph: '16:00' }, { id: 'h_fin', et: 'Horómetro final', tipo: 'number' },
    { id: 'horas', et: 'Horas a valorizar (opcional)', tipo: 'number', ayuda: 'Vacío = horómetro o horario' }, { id: 'combustible', et: 'Combustible (gal)', tipo: 'number' }, { id: 'partida', et: 'Partida', lista: 'dl-partidas' },
    { id: 'trabajo', et: 'Trabajo realizado', col: 2 }, { id: 'metrado', et: 'Metrado', tipo: 'number' }, { id: 'observacion', et: 'Observación', col: 3 },
  ];
  formulario({
    titulo: p ? 'Editar parte diario' : 'Nuevo parte diario', campos, ancho: '860px',
    val: p || { fecha: hoy(), equipo: ult?.equipo || o.equipos[0].codigo, operador: ult?.operador || '', m_de: '07:00', m_a: '12:00', t_de: '13:00', t_a: '16:00', h_ini: ult?.h_fin ?? null },
    alGuardar: (v) => {
      const h = S.motor.horasParte(v);
      if (p) Object.assign(p, v); else o.partes.push({ id: uid('P'), ...v });
      cambio();
      aviso(`Parte guardado: ${fmt(h.valorizar)} h.${h.alertas.length ? ' ' + h.alertas.join('; ') : ''}`, h.alertas.length ? 'alerta' : 'ok');
      refrescar();
    },
  });
}
function partes(box) {
  const o = S.obra, m = S.motor;
  const lista = o.partes.filter((p) => !fMes || mesDe(p.fecha) === fMes).slice().sort((a, b) => (a.fecha < b.fecha ? 1 : a.fecha > b.fecha ? -1 : 0));
  box.innerHTML = `<div class="filtros"><select id="p-mes"><option value="">Todos los meses</option>${mesesDeObra().map((x) => `<option value="${x}" ${x === fMes ? 'selected' : ''}>${nombreMes(x)}</option>`).join('')}</select><span style="flex:1"></span>${btnExport('partes')}</div>
    <div class="tabla-cont"><table class="t"><thead><tr><th>Fecha</th><th>Equipo</th><th>N° parte</th><th>Operador</th><th>Horario</th><th class="n">Horómetro</th><th class="n">Horas</th><th class="n">Gal</th><th>Partida</th><th>Trabajo</th><th class="n">Metrado</th><th></th></tr></thead><tbody>
    ${lista.map((p) => { const h = m.horasParte(p); return `<tr class="${h.alertas.length ? 'e-revisar' : ''}" title="${esc(h.alertas.join('\n'))}"><td>${fFecha(p.fecha)}</td><td class="mono">${esc(p.equipo)}</td><td>${esc(p.n_parte || '')}</td><td style="font-size:12px">${esc(p.operador || '')}</td><td style="font-size:12px">${esc([p.m_de && `${p.m_de}-${p.m_a}`, p.t_de && `${p.t_de}-${p.t_a}`].filter(Boolean).join(' / '))}</td>
      <td class="n" style="font-size:12px">${p.h_ini !== null && p.h_ini !== undefined && p.h_ini !== '' ? `${cant(p.h_ini, 1)} → ${cant(p.h_fin, 1)}` : ''}</td><td class="n"><b>${fmt(h.valorizar)}</b></td><td class="n">${p.combustible ? fmt(p.combustible) : ''}</td><td class="item">${esc(p.partida || '')}</td><td style="font-size:12px"><span class="desc-trunc" style="max-width:260px">${esc(p.trabajo || '')}</span></td><td class="n">${p.metrado ? cant(p.metrado, 2) : ''}</td>
      <td class="acc"><button class="btn-icono" data-ed="${esc(p.id)}">${icono('editar')}</button><button class="btn-icono peligro" data-del="${esc(p.id)}">${icono('borrar')}</button></td></tr>`; }).join('') || '<tr><td colspan="12" class="vacio">Sin partes diarios.</td></tr>'}
  </tbody></table></div>`;
  enlazarExport(box, { partes: () => rPartes(lista.slice().reverse()) });
  box.querySelector('#p-mes').onchange = (e) => { fMes = e.target.value; refrescar(); };
  box.querySelectorAll('[data-ed]').forEach((b) => (b.onclick = () => editarParte(b.dataset.ed)));
  box.querySelectorAll('[data-del]').forEach((b) => (b.onclick = async () => { if (!(await confirmar('¿Eliminar este parte?', { peligro: true }))) return; o.partes = o.partes.filter((p) => p.id !== b.dataset.del); cambio(); refrescar(); }));
}

function generar(box) {
  const o = S.obra, m = S.motor;
  const fs = m.fechasConMetrado();
  if (!gen.desde) { gen.desde = fs.length ? fs[0].slice(0, 8) + '01' : hoy(); gen.hasta = fs.length ? fs[fs.length - 1] : hoy(); }
  const props = m.proponerPartes(gen.desde, gen.hasta);
  const exist = new Set(o.partes.map((p) => `${p.fecha}|${p.equipo}|${p.partida}`));
  const key = `${gen.desde}|${gen.hasta}|${props.length}`;
  if (!gen.props || gen.key !== key) { gen.props = props.map((p) => ({ ...p, dup: exist.has(`${p.fecha}|${p.equipo}|${p.partida}`), ok: !exist.has(`${p.fecha}|${p.equipo}|${p.partida}`) })); gen.key = key; }
  const P = gen.props;
  box.innerHTML = `<div class="nota">Horas de cada equipo = metrado del día × horas por unidad del A.C.U. (solo equipos enlazados a un insumo del A.C.U.). Puede editar las horas antes de registrar los partes.</div>
    <div class="filtros"><label>Desde <input type="date" id="g-d" value="${gen.desde}"></label><label>Hasta <input type="date" id="g-h" value="${gen.hasta}"></label><span style="flex:1"></span><button class="btn btn-primario" id="g-ok" ${P.length ? '' : 'disabled'}>${icono('guardar')}Registrar partes marcados</button></div>
    ${P.length ? `<div class="tabla-cont"><table class="t" id="t-g"><thead><tr><th></th><th>Fecha</th><th>Equipo</th><th>Partida</th><th>Trabajo</th><th class="n">Metrado</th><th class="n">Horas</th><th></th></tr></thead><tbody>
    ${P.map((p, i) => `<tr data-i="${i}" class="${p.dup ? 'tenue' : ''}"><td><input type="checkbox" data-ok ${p.ok ? 'checked' : ''}></td><td>${fFecha(p.fecha)}</td><td class="mono">${esc(p.equipo)}</td><td class="item">${esc(p.partida)}</td><td style="font-size:12px"><span class="desc-trunc">${esc(p.trabajo)}</span></td><td class="n">${cant(p.metrado, 2)}</td><td><input class="n" data-h value="${cant(p.horas, 2)}" style="width:80px"></td><td>${p.dup ? '<span class="badge badge-alerta">ya existe</span>' : ''}</td></tr>`).join('')}
    </tbody></table></div>` : '<div class="panel"><div class="vacio"><h3>Nada que generar</h3><p>No hay metrados en el rango con equipos del A.C.U. enlazados a un equipo registrado.</p></div></div>'}`;
  box.querySelector('#g-d').onchange = (e) => { gen.desde = e.target.value; gen.props = null; refrescar(); };
  box.querySelector('#g-h').onchange = (e) => { gen.hasta = e.target.value; gen.props = null; refrescar(); };
  box.querySelector('#t-g tbody')?.addEventListener('change', (e) => { const tr = e.target.closest('tr'); const p = P[Number(tr.dataset.i)]; if (e.target.matches('[data-ok]')) p.ok = e.target.checked; if (e.target.matches('[data-h]')) p.horas = num(e.target.value); });
  box.querySelector('#g-ok')?.addEventListener('click', () => {
    const sel = P.filter((p) => p.ok && p.horas > 0);
    sel.forEach((p) => o.partes.push({ id: uid('P'), fecha: p.fecha, equipo: p.equipo, n_parte: '', operador: o.equipos.find((e) => e.codigo === p.equipo)?.operador || '', m_de: '', m_a: '', t_de: '', t_a: '', h_ini: null, h_fin: null, horas: p.horas, combustible: null, partida: p.partida, trabajo: p.trabajo, metrado: p.metrado, observacion: 'Según A.C.U. × metrado' }));
    cambio({ guardarYa: true }); gen.props = null;
    aviso(`${sel.length} parte(s) registrado(s).`);
    tab = 'partes'; refrescar();
  });
}

function control(box) {
  const m = S.motor;
  const meses = mesesDeObra();
  if (!cMes) cMes = meses.filter((x) => S.obra.partes.some((p) => p.fecha.startsWith(x))).pop() || meses[meses.length - 1] || mesDe(hoy());
  const nd = diasDelMes(cMes);
  const datos = m.controlEquipos(cMes);
  box.innerHTML = `<div class="filtros"><select id="c-mes">${meses.map((x) => `<option value="${x}" ${x === cMes ? 'selected' : ''}>${nombreMes(x)}</option>`).join('')}</select><span style="flex:1"></span>${btnExport('c415', '4.15')}</div>
    <div class="tabla-cont"><table class="t grilla"><thead><tr><th>Equipo</th><th>Und</th><th class="n">Contratado</th><th class="n">Anterior</th>${Array.from({ length: nd }, (_, i) => `<th class="n">${i + 1}</th>`).join('')}<th class="n">Mes</th><th class="n">Acumulado</th><th class="n">Saldo O/S</th><th class="n">Valorizado S/</th></tr></thead><tbody>
    ${datos.map((e) => `<tr class="${e.excede ? 'e-novalorizar' : ''}"><td><span class="mono">${esc(e.codigo)}</span> ${esc(e.descripcion)}</td><td>${esc(e.unidad)}</td><td class="n">${cant(e.cantidad, 2)}</td><td class="n">${cant(e.anterior, 2)}</td>${e.dias.map((h) => `<td class="n" style="font-size:11px">${h ? cant(h, 2) : ''}</td>`).join('')}<td class="n"><b>${cant(e.mes, 2)}</b></td><td class="n">${cant(e.acumulado, 2)}</td><td class="n ${e.saldo < 0 ? 'neg' : ''}">${cant(e.saldo, 2)}</td><td class="n">${fmt(e.valorizado)}</td></tr>`).join('') || `<tr><td colspan="${nd + 8}" class="vacio">Sin equipos.</td></tr>`}
    </tbody><tfoot><tr class="tot"><td colspan="${nd + 7}">TOTAL VALORIZADO DEL MES</td><td class="n">${fmt(datos.reduce((s, e) => s + e.valorizado, 0))}</td></tr></tfoot></table></div>`;
  enlazarExport(box, { c415: () => rControlEquipos(cMes) });
  box.querySelector('#c-mes').onchange = (e) => { cMes = e.target.value; refrescar(); };
}
