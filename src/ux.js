/* Asistente de 4 pasos: navegación, revisión de cobertura y ayudas para la operadora. */
const U = { paso: 1, max: 1, local: false };
const MESES_CORTOS = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
const PLATAFORMAS = [['Niubiz', 'Niubiz'], ['AFRUS', 'AFRUS'], ['MC Procesos', 'MC / Izipay']];

function enSesion() { return typeof M !== 'undefined' && !!M.token; }

function irA(n) {
  U.paso = n; U.max = Math.max(U.max, n);
  for (let k = 1; k <= 4; k++) $('#paso' + k).hidden = k !== n;
  uxActualizar();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function pasoPermitido(k) {
  if (k <= U.max) return true;
  if (k === 4 && enSesion() && M.lotes && M.lotes.length) return true;
  return false;
}

/* Cobertura: qué meses tienen pagos exitosos por plataforma. */
function cobertura() {
  const P = params();
  const anio = P.desde ? P.desde.y : new Date().getFullYear();
  const hoy = new Date();
  const mesHoy = hoy.getFullYear() > anio ? 12 : hoy.getFullYear() < anio ? 0 : hoy.getMonth() + 1;
  const filas = []; const faltan = [];
  for (const [clave, nombre] of PLATAFORMAS) {
    const n = Array(13).fill(0);
    for (const p of S.pagos.values()) if (p.ok && p.plataforma === clave && p.f && p.f.y === anio) n[p.f.m]++;
    const primero = n.findIndex((x, i) => i > 0 && x > 0);
    const celdas = [];
    const huecos = [];
    for (let m = 1; m <= 12; m++) {
      if (n[m]) celdas.push(`<span class="mes ok" title="${n[m]} pagos">✓</span>`);
      else if (m > mesHoy) celdas.push('<span class="mes futuro">·</span>');
      else if (primero < 0 || m < primero) celdas.push('<span class="mes futuro" title="Sin pagos">–</span>');
      else { celdas.push('<span class="mes falta" title="Sin pagos este mes">!</span>'); huecos.push(MESES_CORTOS[m - 1]); }
    }
    if (primero < 0) faltan.push(`${nombre}: no hay ningún reporte`);
    else if (huecos.length) faltan.push(`${nombre}: ${huecos.join(', ')}`);
    filas.push(`<tr><td>${nombre}</td>${celdas.map((c) => `<td>${c}</td>`).join('')}</tr>`);
  }
  return { anio, faltan, html: `<div class="tabla-scroll" style="margin-bottom:6px"><table class="cob"><thead><tr><th>Pagos ${anio}</th>${MESES_CORTOS.map((m) => `<th>${m}</th>`).join('')}</tr></thead><tbody>${filas.join('')}</tbody></table></div>
    <p class="small muted" style="margin:0 0 14px"><span class="mes ok" style="width:20px;height:20px;line-height:20px">✓</span> hay pagos · <span class="mes falta" style="width:20px;height:20px;line-height:20px">!</span> no hay pagos ese mes · <span class="mes futuro">·</span> todavía no llega</p>` };
}

function pintarEstado1() {
  const hayPagos = [...S.pagos.values()].some((p) => p.ok);
  if (!hayPagos) {
    $('#cobertura').innerHTML = '';
    $('#estado-general').innerHTML = $('#cargando').hidden ? `<li><span class="ico warn">!</span><div><b>Todavía no hay reportes de pago.</b><div class="det">${enSesion() ? 'Guarda los reportes de Niubiz, AFRUS y MC/Izipay en la carpeta <b>01_REPORTES</b> y pulsa <b>Leer la carpeta otra vez</b>.' : 'Arrástralos en <b>Agregar un archivo</b>.'}</div></div></li>` : '';
    $('#motivo-1').textContent = $('#cargando').hidden ? 'Faltan los reportes de pago.' : '';
    return;
  }
  const c = cobertura();
  $('#cobertura').innerHTML = c.html;
  const it = [];
  const exitosos = [...S.pagos.values()].filter((p) => p.ok && p.f && p.f.y === c.anio).length;
  it.push(c.faltan.length
    ? ['warn', 'Revisa si faltan reportes', `Meses sin pagos — ${c.faltan.map(esc).join(' · ')}. Si falta un reporte, descárgalo de la plataforma, guárdalo en <b>01_REPORTES</b> y pulsa <b>Leer la carpeta otra vez</b>. Si de verdad no hubo pagos ese mes, puedes continuar.`]
    : ['ok', `Reportes completos: ${exitosos.toLocaleString('es-PE')} pagos exitosos en ${c.anio}`, 'Ya se quitaron los pagos rechazados, fallidos, reembolsados y repetidos.']);
  it.push(S.base.size
    ? ['ok', `Base de correos: ${S.base.size.toLocaleString('es-PE')} personas`, 'De aquí salen los correos y nombres de los donantes.']
    : ['warn', 'No se encontró la base de correos', 'Debe estar en <b>02_BASE_DONANTES</b>. Sin ella, muchas personas quedarán sin correo.']);
  it.push(S.plantillas.personas
    ? ['ok', 'Plantilla oficial lista', 'El certificado se arma sobre el formato oficial de Crea+.']
    : ['mal', 'Falta la plantilla del certificado', 'Debe estar en <b>05_HERRAMIENTA</b> con el nombre <b>plantilla_personas.pdf</b>.']);
  if (S.emitidos.size) it.push(['info', `${S.emitidos.size.toLocaleString('es-PE')} personas ya tienen su certificado de este año`, 'No se les volverá a crear.']);
  const raros = S.archivos.filter((a) => a.tipo === 'DESCONOCIDO');
  if (raros.length) it.push(['mal', `${raros.length} archivo(s) no se reconocieron y no se usan`, raros.map((a) => esc(a.nombre)).join(', ')]);
  if (enSesion() && M.lotes) {
    const pend = M.lotes.reduce((s, l) => s + pendientesDe(l).pend.length, 0);
    if (pend) it.push(['warn', `Hay ${pend.toLocaleString('es-PE')} certificados ya creados que faltan enviar`, '<button class="sec" data-ir="4" style="margin-top:6px">Ir a enviarlos</button>']);
  }
  const ico = { ok: '✓', warn: '!', mal: '✕', info: 'i' };
  $('#estado-general').innerHTML = it.map(([t, tit, det]) => `<li><span class="ico ${t}">${ico[t]}</span><div><b>${tit}</b><div class="det">${det}</div></div></li>`).join('');
  $('#motivo-1').textContent = '';
}

function uxActualizar() {
  for (const b of document.querySelectorAll('.pbtn')) {
    const k = +b.dataset.p;
    b.classList.toggle('actual', k === U.paso);
    b.classList.toggle('hecho', k !== U.paso && pasoPermitido(k));
    b.querySelector('.n').textContent = k !== U.paso && k < U.max ? '✓' : k;
  }
  const listos = S.donantes.filter((d) => d.incluir && d.estado === 'LISTO').length;
  const muestras = S.muestrasVistas.size;
  const ok2 = listos > 0 && muestras >= 3;
  $('#btn-a-paso3').disabled = !ok2;
  $('#motivo-2').textContent = !listos ? 'No hay donantes listos.' : muestras < 3 ? `Mira ${3 - muestras} ejemplo(s) más.` : '';
  $('#btn-muestra').textContent = muestras ? 'Ver otro ejemplo' : 'Ver un ejemplo';
  const creado = !$('#post-crear').hidden;
  $('#btn-a-paso4').hidden = !creado;
  $('#form-crear').hidden = creado;
  if (creado) { $('#btn-generar-sp').hidden = true; $('#motivo-3').textContent = ''; }
  else if (enSesion()) $('#btn-generar-sp').hidden = false;
}

function uxSesion(dentro) {
  const ti = /ti\b/.test(location.hash);
  $('#config-ti').hidden = !ti;
  $('#link-local').hidden = !ti;
  const activo = dentro || U.local;
  $('#bienvenida').hidden = activo;
  $('#pasos').hidden = !activo;
  if (!activo) { for (let k = 1; k <= 4; k++) $('#paso' + k).hidden = true; return; }
  $('#det-zona').open = !dentro;
  irA(U.paso);
  if (dentro && !M.leido && !M.leyendo) actualizarDesdeSharePoint();
}

function uxIniciar() {
  const _pintarArchivos = pintarArchivos;
  pintarArchivos = function () { _pintarArchivos(); pintarEstado1(); };
  const _pintarLotes = pintarLotes;
  pintarLotes = function () { _pintarLotes(); if (U.paso === 1) pintarEstado1(); uxActualizar(); };
  $('#pasos').addEventListener('click', (e) => { const b = e.target.closest('.pbtn'); if (b && pasoPermitido(+b.dataset.p)) irA(+b.dataset.p); });
  document.addEventListener('click', (e) => { const b = e.target.closest('[data-ir]'); if (b) irA(+b.dataset.ir); });
  $('#btn-a-paso3').addEventListener('click', () => { $('#post-crear').hidden = true; actualizarPaso3(); irA(3); });
  $('#btn-a-paso4').addEventListener('click', () => irA(4));
  $('#btn-muestra').addEventListener('click', () => {
    const listos = S.donantes.map((d, i) => [d, i]).filter(([d]) => d.incluir && d.estado === 'LISTO');
    const nuevos = listos.filter(([d]) => !S.muestrasVistas.has(d.doc));
    const pool = nuevos.length ? nuevos : listos;
    if (!pool.length) { aviso('No hay donantes listos para mostrar.', 'warn'); return; }
    // Preferir ejemplos variados: el que tiene más líneas, uno con varias plataformas y uno al azar.
    const orden = [(a, b) => b[0].lineas.length - a[0].lineas.length, (a, b) => b[0].plataformas.length - a[0].plataformas.length, () => Math.random() - 0.5];
    const pick = [...pool].sort(orden[Math.min(S.muestrasVistas.size, 2)])[0];
    verMuestra(pick[1]);
  });
  $('#btn-modo-local').addEventListener('click', () => { U.local = true; uxSesion(false); });
  window.addEventListener('hashchange', () => uxSesion(enSesion()));
  if (typeof M === 'undefined') uxSesion(false);
}
window.addEventListener('DOMContentLoaded', uxIniciar);
