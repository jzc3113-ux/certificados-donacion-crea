/* Certificados de donación Crea+ — app local (todo se procesa en este navegador). */
'use strict';
const { PDFDocument, StandardFonts, rgb } = PDFLib;

/* ------------------------------------------------------------------ constantes */
const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto',
  'septiembre', 'octubre', 'noviembre', 'diciembre'];
const MESES_CORTO = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'set', 'oct', 'nov', 'dic'];
const SERIE = '002-000021,';
const PAGE_W = 595.304, PAGE_H = 841.89;
const MAX_LINEAS = 8;

/* Posiciones medidas sobre la plantilla oficial (puntos PDF, origen abajo-izquierda). */
const LAYOUT = {
  personas: {
    size: 9.41, numero: { x: 439.2, y: 604.66 }, nombre: { x: 154.2, y: 456.38, maxW: 365 },
    doc: { x: 154.2, y: 432.45 },
    filas: { ys: [392.72, 374.42, 356.11, 337.81, 319.5, 301.19, 282.89, 264.58], cantCx: 107.5, descX: 154.2, descMaxW: 280, valorRx: 521.2, monedaX: 447.8 },
    total: { rx: 521.0, y: 236.55, size: 9.91 }, ciudad: { x: 69.7, y: 161.76 }, fecha: { x: 101.0, y: 161.76 },
  },
  empresas: {
    size: 9.21, numero: { x: 435.2, y: 612.81 }, nombre: { x: 165.7, y: 467.66, maxW: 345 },
    doc: { x: 165.7, y: 444.24 }, direccion: { x: 165.7, y: 421.19, size: 8.29, maxW: 345 },
    filas: { ys: [381.99, 357.16, 339.25, 321.33, 303.42, 285.5, 267.58, 249.67], cantCx: 115.9, descX: 165.7, descMaxW: 270, valorRx: 515.3, monedaX: 443.5 },
    total: { rx: 515.2, y: 222.21, size: 9.6 }, ciudad: { x: 69.7, y: 149.03 }, fecha: { x: 109.9, y: 149.03 },
  },
};

/* ------------------------------------------------------------------ estado */
const S = {
  archivos: [],          // {nombre, tipo, filas, exitosos, mensaje}
  pagos: new Map(),      // id -> pago (todos los estados)
  duplicados: 0,
  base: new Map(),       // doc -> {nombres:Set, correos:Set}
  basePorCorreo: new Map(),
  emitidos: new Map(),   // doc -> numero (lotes anteriores de esta app: bloquean)
  previos: new Map(),    // doc -> [números] (certificados del Excel: solo informan)
  maxEmitido: 0,
  donantes: [],          // resultado de la consolidación
  sinDoc: [],            // pagos exitosos sin documento
  plantillas: { personas: null, empresas: null }, // Uint8Array
  muestrasVistas: new Set(),
  metric: null,          // fuente para medir textos
  pagina: 0,
};

/* ------------------------------------------------------------------ utilidades */
const $ = (s) => document.querySelector(s);
const limpiar = (v) => String(v == null ? '' : v).replace(/ /g, ' ').replace(/\s+/g, ' ').trim();
const sinTildes = (s) => limpiar(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const importe = (n) => n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const soles = (n) => 'S/ ' + importe(n);
const correoValido = (c) => /^[^\s@,;]+@[^\s@,;]+\.[a-z]{2,}$/i.test(c);

function numero(v) {
  let s = limpiar(v).replace(/[^\d.,-]/g, '');
  if (!s) return NaN;
  if (s.includes(',') && s.includes('.')) s = s.replace(/,/g, '');
  else if (s.includes(',')) s = /,\d{1,2}$/.test(s) ? s.replace(',', '.') : s.replace(/,/g, '');
  return parseFloat(s);
}

function normDoc(v) {
  let d = limpiar(v).replace(/[^0-9]/g, '');
  if (d.length === 7) d = '0' + d;
  return d;
}

/* Fecha → {y,m,d}. Acepta dd/mm/yyyy, yyyy-mm-dd (con hora) y número serial de Excel. */
function fecha(v) {
  const s = limpiar(v);
  if (!s) return null;
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return { y: +m[3], m: +m[2], d: +m[1] };
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return { y: +m[1], m: +m[2], d: +m[3] };
  if (/^\d{5}(\.\d+)?$/.test(s)) {
    const dt = new Date(Date.UTC(1899, 11, 30) + Math.floor(+s) * 86400000);
    return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
  }
  return null;
}
const fkey = (f) => f ? f.y * 10000 + f.m * 100 + f.d : 0;
const ftexto = (f) => `${f.d} de ${MESES[f.m - 1]} de ${f.y}`;

function aviso(texto, tipo = 'info') {
  const el = document.createElement('div');
  el.className = 'aviso ' + tipo;
  el.textContent = texto;
  const caja = $('#avisos');
  caja.prepend(el);
  while (caja.children.length > 4) caja.lastChild.remove();
  if (tipo !== 'error') setTimeout(() => el.remove(), 9000);
}

/* ------------------------------------------------------------------ lectura CSV / XLSX */
async function decodificar(buf) {
  try { return new TextDecoder('utf-8', { fatal: true }).decode(buf).replace(/^﻿/, ''); }
  catch (e) { return new TextDecoder('windows-1252').decode(buf); }
}

function separador(linea) {
  const c = { ',': 0, ';': 0, '\t': 0 };
  let q = false;
  for (const ch of linea) { if (ch === '"') q = !q; else if (!q && ch in c) c[ch]++; }
  return Object.entries(c).sort((a, b) => b[1] - a[1])[0][0];
}

/* Recorre el CSV fila por fila sin guardar todo en memoria. */
function recorrerCSV(texto, alFila) {
  const nl = texto.indexOf('\n');
  const sep = separador(nl > 0 ? texto.slice(0, nl) : texto);
  let fila = [], campo = '', q = false, i = 0;
  const n = texto.length;
  while (i < n) {
    const ch = texto[i];
    if (q) {
      if (ch === '"') { if (texto[i + 1] === '"') { campo += '"'; i++; } else q = false; }
      else campo += ch;
    } else if (ch === '"') q = true;
    else if (ch === sep) { fila.push(campo); campo = ''; }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && texto[i + 1] === '\n') i++;
      fila.push(campo); campo = '';
      if (fila.length > 1 || fila[0] !== '') alFila(fila);
      fila = [];
    } else campo += ch;
    i++;
  }
  if (campo !== '' || fila.length) { fila.push(campo); alFila(fila); }
}

async function leerXLSX(buf) {
  const zip = await JSZip.loadAsync(buf);
  const parser = new DOMParser();
  const xml = async (p) => { const f = zip.file(p); return f ? parser.parseFromString(await f.async('string'), 'application/xml') : null; };
  const compartidas = [];
  const ss = await xml('xl/sharedStrings.xml');
  if (ss) for (const si of ss.getElementsByTagName('si')) {
    let t = ''; for (const x of si.getElementsByTagName('t')) t += x.textContent; compartidas.push(t);
  }
  const wb = await xml('xl/workbook.xml');
  const rels = await xml('xl/_rels/workbook.xml.rels');
  const destino = {};
  for (const r of rels.getElementsByTagName('Relationship')) destino[r.getAttribute('Id')] = r.getAttribute('Target');
  const hojas = [];
  for (const sh of wb.getElementsByTagName('sheet')) {
    const rid = sh.getAttribute('r:id') || sh.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    let ruta = destino[rid] || '';
    ruta = ruta.startsWith('/') ? ruta.slice(1) : 'xl/' + ruta.replace(/^\.\//, '');
    const doc = await xml(ruta);
    if (!doc) continue;
    const filas = [];
    for (const row of doc.getElementsByTagName('row')) {
      const f = [];
      for (const c of row.getElementsByTagName('c')) {
        const ref = c.getAttribute('r') || '';
        const letras = ref.replace(/\d+/g, '');
        let col = 0; for (const L of letras) col = col * 26 + (L.charCodeAt(0) - 64); col -= 1;
        const t = c.getAttribute('t');
        let v = '';
        if (t === 'inlineStr') { for (const x of c.getElementsByTagName('t')) v += x.textContent; }
        else {
          const ve = c.getElementsByTagName('v')[0];
          v = ve ? ve.textContent : '';
          if (t === 's') v = compartidas[+v] || '';
        }
        f[col < 0 ? f.length : col] = v;
      }
      for (let k = 0; k < f.length; k++) if (f[k] === undefined) f[k] = '';
      filas.push(f);
    }
    hojas.push({ nombre: sh.getAttribute('name'), filas });
  }
  return hojas;
}

/* ------------------------------------------------------------------ reconocimiento de formatos */
const FIRMAS = {
  NIUBIZ: ['id.transaccion', 'estado cargo', 'importe cargo', 'num. iden. beneficiario'],
  AFRUS: ['id', 'status', 'amount', 'identification'],
  MC: ['numero recibo', 'estado transaccion', 'accion', 'importe cargo'],
};

function indice(cab) {
  const ix = {};
  cab.forEach((c, i) => { const k = sinTildes(c); if (k && !(k in ix)) ix[k] = i; });
  return ix;
}

function reconocer(cab) {
  const ix = indice(cab);
  for (const [tipo, cols] of Object.entries(FIRMAS)) if (cols.every((c) => c in ix)) return { tipo, ix };
  if ('numero_certificado' in ix && ('documento' in ix || 'dni' in ix)) return { tipo: 'REGISTRO', ix };
  const claves = Object.keys(ix);
  const doc = claves.find((k) => /^(dni|dni\/ruc|ruc|documento|nro\.? ?documento|n° documento|numero de documento|identificacion|doc|dni_ruc|dni o ruc)$/.test(k) || /^dni\b/.test(k));
  const correo = claves.find((k) => /(correo|e-?mail|mail)/.test(k));
  if (doc && correo) {
    const nombre = claves.find((k) => /^(nombre completo|nombre_completo|nombres y apellidos|nombre|nombres|razon social|senor\(es\)\/razon social)$/.test(k));
    const ape = claves.find((k) => /^(apellidos?|apellido paterno)$/.test(k));
    return { tipo: 'BASE', ix, cols: { doc: ix[doc], correo: ix[correo], nombre: nombre != null ? ix[nombre] : null, ape: ape != null ? ix[ape] : null } };
  }
  return null;
}

/* ------------------------------------------------------------------ normalización por plataforma */
function agregarPago(p) {
  if (S.pagos.has(p.id)) { S.duplicados++; return false; }
  S.pagos.set(p.id, p);
  return true;
}

function nombreMC(v) {
  const s = limpiar(v);
  if (!s.includes(',')) return s;
  const [ap, no] = s.split(',');
  return limpiar(no + ' ' + ap);
}

function procesarFila(tipo, ix, r, archivo) {
  const g = (k) => limpiar(r[ix[k]]);
  if (tipo === 'NIUBIZ') {
    const estado = g('estado cargo').toUpperCase();
    const doc = normDoc(g('num. iden. beneficiario'));
    const f = fecha(g('fecha proceso pago')) || fecha(g('fecha estimada pago'));
    const idt = g('id.transaccion'), aut = g('cod.autorizacion');
    const ok = estado === 'AUTORIZADO' && !!aut;
    const id = 'NIUBIZ:' + (idt || ['SIN-ID', g('nro. lote'), doc, g('fecha proceso pago'), g('nro. reintento'), g('importe cargo'), aut].join('|'));
    return { id, plataforma: 'Niubiz', ok, estado, doc, f, monto: numero(g('importe cargo')), nombre: g('nom. beneficiario') || g('nom. tarjeta habiente'), correo: '', archivo };
  }
  if (tipo === 'AFRUS') {
    const estado = g('status').toUpperCase();
    const f = fecha(g('created_at_utc_x_usuario')) || fecha(g('created_at')) ||
      (ix.year != null ? { y: +g('year'), m: +g('month'), d: +g('day') } : null);
    const moneda = sinTildes(g('currency'));
    return {
      id: 'AFRUS:' + g('id'), plataforma: 'AFRUS', ok: estado === 'SUCCEEDED' && (!moneda || moneda.includes('sol')), estado,
      doc: normDoc(g('identification')), f, monto: numero(g('amount')),
      nombre: limpiar(g('first_name') + ' ' + g('last_name')), correo: g('email').toLowerCase(), campana: g('campaign'), archivo,
    };
  }
  if (tipo === 'MC') {
    const estado = g('estado transaccion');
    const accion = g('accion');
    const ok = accion.toLowerCase() === 'aprobado' && estado.toLowerCase().startsWith('deposit');
    const cs = normDoc(g('codigo servicio'));
    return {
      id: 'MC:' + g('numero recibo'), plataforma: 'MC Procesos', ok, estado: estado + ' / ' + accion,
      doc: /^\d{8}$/.test(cs) ? cs : '', f: fecha(g('fecha deposito')) || fecha(g('fecha autorizacion')),
      monto: numero(g('importe cargo')), nombre: nombreMC(g('tarjetahabiente') || g('usuario')), correo: '', archivo,
    };
  }
  return null;
}

async function cargarArchivo(file) {
  const buf = await file.arrayBuffer();
  const esXlsx = /\.xls[xm]$/i.test(file.name);
  const info = { nombre: file.name, tipo: '?', filas: 0, exitosos: 0, mensaje: '' };
  const tratarTabla = (cabecera, recorrer) => {
    const rec = reconocer(cabecera);
    if (!rec) return false;
    info.tipo = rec.tipo;
    if (rec.tipo === 'BASE') {
      recorrer((r) => {
        const doc = normDoc(r[rec.cols.doc]);
        if (!doc) return;
        info.filas++;
        const e = S.base.get(doc) || { nombres: new Set(), correos: new Set() };
        let nom = rec.cols.nombre != null ? limpiar(r[rec.cols.nombre]) : '';
        if (rec.cols.ape != null) nom = limpiar(nom + ' ' + limpiar(r[rec.cols.ape]));
        if (nom) e.nombres.add(nom);
        for (const c of limpiar(r[rec.cols.correo]).toLowerCase().split(/[;,\s]+/)) if (c) {
          e.correos.add(c);
          const l = S.basePorCorreo.get(c) || new Set(); l.add(doc); S.basePorCorreo.set(c, l);
        }
        S.base.set(doc, e);
      });
      return true;
    }
    if (rec.tipo === 'REGISTRO') {
      const kdoc = 'documento' in rec.ix ? 'documento' : 'dni';
      const esLote = 'correo' in rec.ix && 'archivo' in rec.ix; // manifiesto generado por esta app
      info.tipo = esLote ? 'REGISTRO' : 'PREVIOS';
      recorrer((r) => {
        const num = limpiar(r[rec.ix.numero_certificado]); const doc = normDoc(r[rec.ix[kdoc]]);
        if (!num || !doc) return;
        info.filas++;
        if (esLote) S.emitidos.set(doc, num);
        else { const l = S.previos.get(doc) || []; l.push(num); S.previos.set(doc, l); }
        if (num.startsWith(SERIE)) { const n = parseInt(num.slice(SERIE.length), 10); if (n > S.maxEmitido) S.maxEmitido = n; }
      });
      return true;
    }
    recorrer((r) => {
      const p = procesarFila(rec.tipo, rec.ix, r, file.name);
      if (!p || !p.id || /:$/.test(p.id)) return;
      info.filas++;
      if (agregarPago(p) && p.ok) info.exitosos++;
    });
    return true;
  };

  if (esXlsx) {
    const hojas = await leerXLSX(buf);
    let alguna = false;
    for (const h of hojas) {
      for (let i = 0; i < Math.min(25, h.filas.length); i++) {
        if (reconocer(h.filas[i])) {
          tratarTabla(h.filas[i], (fn) => { for (let k = i + 1; k < h.filas.length; k++) fn(h.filas[k]); });
          alguna = true; break;
        }
      }
    }
    if (!alguna) info.tipo = 'DESCONOCIDO';
  } else {
    const texto = await decodificar(buf);
    let cab = null, filas = [], ok = false;
    // Buscar la fila de encabezados entre las primeras 25 líneas.
    const primeras = [];
    recorrerCSV(texto.slice(0, 20000), (r) => { if (primeras.length < 25) primeras.push(r); });
    const pos = primeras.findIndex((r) => reconocer(r));
    if (pos >= 0) {
      cab = primeras[pos];
      ok = tratarTabla(cab, (fn) => { let k = 0; recorrerCSV(texto, (r) => { if (k++ > pos) fn(r); }); });
    }
    if (!ok) info.tipo = 'DESCONOCIDO';
  }
  const nombres = { NIUBIZ: 'Niubiz', AFRUS: 'AFRUS', MC: 'MC Procesos / Izipay', BASE: 'Base de donantes', REGISTRO: 'Lote anterior de esta app (bloquea reenvío)', PREVIOS: 'Certificados emitidos antes en Excel (solo aviso)', DESCONOCIDO: 'No reconocido' };
  info.tipoTexto = nombres[info.tipo] || info.tipo;
  if (info.tipo === 'DESCONOCIDO') info.mensaje = 'No reconozco las columnas de este archivo. No se usó.';
  S.archivos.push(info);
  return info;
}

/* ------------------------------------------------------------------ consolidación */
function params() {
  return {
    desde: fecha($('#p-desde').value), hasta: fecha($('#p-hasta').value),
    emision: fecha($('#p-emision').value), ciudad: limpiar($('#p-ciudad').value) || 'Lima,',
    concepto: limpiar($('#p-concepto').value) || 'Donación Amigo Crea+',
    inmobiliaria: limpiar($('#p-inmob').value) || 'Amigos Crea+',
  };
}

function mesesTexto(meses, anio, concepto, font, size, maxW) {
  const ms = [...new Set(meses)].sort((a, b) => a - b);
  const pref = `${concepto} – Período ${anio} (`;
  const lista = ms.map((m) => MESES[m - 1]).join(', ');
  const rangos = []; let ini = ms[0], prev = ms[0];
  for (const m of ms.slice(1).concat([99])) {
    if (m === prev + 1) { prev = m; continue; }
    rangos.push(ini === prev ? MESES[ini - 1] : (prev === ini + 1 ? `${MESES[ini - 1]}, ${MESES[prev - 1]}` : `${MESES[ini - 1]} a ${MESES[prev - 1]}`));
    ini = prev = m;
  }
  const corto = ms.map((m) => MESES_CORTO[m - 1]).join(', ');
  const opciones = [pref + lista + ')', pref + rangos.join(', ') + ')', pref + corto + ')'];
  if (!font) return opciones[0];
  // Preferir el texto más legible que quepa, aceptando reducir la letra hasta 8.4 pt.
  for (const tam of [size, 8.4]) for (const o of opciones) if (font.widthOfTextAtSize(o, tam) <= maxW) return o;
  return opciones[2];
}

function consolidar() {
  const P = params();
  const kd = fkey(P.desde), kh = fkey(P.hasta);
  const porDoc = new Map();
  S.sinDoc = [];
  let fueraPeriodo = 0;
  for (const p of S.pagos.values()) {
    if (!p.ok) continue;
    if (!p.f || fkey(p.f) < kd || fkey(p.f) > kh) { fueraPeriodo++; continue; }
    if (!(p.monto > 0)) continue;
    let doc = p.doc;
    if (!doc && p.correo && S.basePorCorreo.has(p.correo) && S.basePorCorreo.get(p.correo).size === 1) doc = [...S.basePorCorreo.get(p.correo)][0];
    if (!doc) { S.sinDoc.push(p); continue; }
    if (!porDoc.has(doc)) porDoc.set(doc, []);
    porDoc.get(doc).push(p);
  }
  const anterior = new Map(S.donantes.map((d) => [d.doc, d]));
  S.donantes = [];
  for (const [doc, pagos] of porDoc) {
    pagos.sort((a, b) => fkey(a.f) - fkey(b.f));
    const base = S.base.get(doc);
    const nombresPlat = pagos.map((p) => p.nombre).filter(Boolean);
    const correosPlat = [...new Set(pagos.map((p) => p.correo).filter(Boolean))];
    const prev = anterior.get(doc);
    const d = {
      doc, pagos,
      tipo: doc.length === 11 ? 'empresas' : 'personas',
      nombre: prev && prev.editadoNombre ? prev.nombre : (base && base.nombres.size ? [...base.nombres][0] : (nombresPlat[nombresPlat.length - 1] || '')),
      correos: base && base.correos.size ? [...base.correos] : correosPlat,
      correo: '', direccion: prev ? prev.direccion || '' : '',
      editadoNombre: prev ? prev.editadoNombre : false, editadoCorreo: prev ? prev.editadoCorreo : false,
      incluir: prev ? prev.incluir : true,
      emitido: S.emitidos.get(doc) || '',
      previos: S.previos.get(doc) || [],
      total: Math.round(pagos.reduce((s, p) => s + p.monto, 0) * 100) / 100,
      plataformas: [...new Set(pagos.map((p) => p.plataforma))].join(', '),
      variosNombresBase: base ? base.nombres.size > 1 : false,
    };
    d.correo = prev && prev.editadoCorreo ? prev.correo : (d.correos.length === 1 ? d.correos[0] : '');
    // Líneas: una por importe distinto.
    const porMonto = new Map();
    for (const p of pagos) { const k = p.monto.toFixed(2); if (!porMonto.has(k)) porMonto.set(k, []); porMonto.get(k).push(p); }
    d.lineas = [...porMonto.entries()].sort((a, b) => fkey(a[1][0].f) - fkey(b[1][0].f)).map(([k, ps]) => ({
      cantidad: ps.length, valor: +k, meses: ps.map((p) => p.f.m), anio: ps[0].f.y,
    }));
    if (prev === undefined && d.emitido) d.incluir = false;
    validar(d);
    S.donantes.push(d);
  }
  S.donantes.sort((a, b) => a.nombre.localeCompare(b.nombre, 'es'));
  return { fueraPeriodo };
}

function validar(d) {
  const obs = [];
  if (d.tipo === 'personas' && !/^\d{8}$/.test(d.doc)) obs.push(`Documento de ${d.doc.length} dígitos: no parece DNI`);
  if (d.tipo === 'empresas' && !d.direccion) obs.push('Empresa: falta la dirección');
  if (!d.nombre) obs.push('Sin nombre');
  if (d.variosNombresBase && !d.editadoNombre) obs.push('La base tiene más de un nombre para este documento');
  if (!d.correo) obs.push(d.correos.length > 1 ? `Hay ${d.correos.length} correos distintos: elige uno` : 'Sin correo');
  else if (!correoValido(d.correo)) obs.push('Correo con formato inválido');
  if (d.lineas.length > MAX_LINEAS) obs.push(`Tiene ${d.lineas.length} importes distintos (máximo ${MAX_LINEAS})`);
  if (new Set(d.lineas.map((l) => l.anio)).size > 1) obs.push('Pagos de más de un año en el mismo certificado');
  if (S.metric && d.nombre && S.metric.widthOfTextAtSize(d.nombre, 7) > LAYOUT[d.tipo].nombre.maxW) obs.push('Nombre demasiado largo para el certificado');
  if (d.emitido) obs.push(`Ya tiene certificado emitido: ${d.emitido}`);
  d.obs = obs;
  d.nota = d.previos.length ? `Tenía certificado(s) anterior(es): ${d.previos.join(', ')}. Este certificado anual lo(s) reemplaza.` : '';
  d.estado = obs.length ? (d.emitido && obs.length === 1 ? 'EMITIDO' : 'REVISAR') : 'LISTO';
}

/* ------------------------------------------------------------------ PDF */
function filtrarTexto(font, t) {
  const permitidos = new Set(font.getCharacterSet());
  let out = '';
  for (const ch of String(t)) {
    if (permitidos.has(ch.codePointAt(0))) out += ch;
    else { const base = ch.normalize('NFD').replace(/[̀-ͯ]/g, ''); out += [...base].every((c) => permitidos.has(c.codePointAt(0))) ? base : '?'; }
  }
  return out;
}

function ajustar(font, texto, size, maxW, min = 6.5) {
  let s = size;
  while (s > min && font.widthOfTextAtSize(texto, s) > maxW) s -= 0.2;
  return s;
}

async function pdfCertificado(d, numeroCert, P, plantillaDoc) {
  const L = LAYOUT[d.tipo];
  const doc = await PDFDocument.create();
  doc.setTitle(`Certificado de donación ${numeroCert}`);
  doc.setAuthor('Asociación Civil CREA MAS PERÚ');
  doc.setCreator('Certificados de donación Crea+');
  const [fondo] = await doc.embedPages([plantillaDoc.getPage(0)]);
  const page = doc.addPage([PAGE_W, PAGE_H]);
  page.drawPage(fondo, { x: 0, y: 0, width: PAGE_W, height: PAGE_H });
  const reg = await doc.embedFont(StandardFonts.Helvetica);
  const neg = await doc.embedFont(StandardFonts.HelveticaBold);
  const negro = rgb(0, 0, 0);
  const T = (txt, x, y, size, font = reg) => page.drawText(filtrarTexto(font, txt), { x, y, size, font, color: negro });
  const Tder = (txt, rx, y, size, font = reg) => { const t = filtrarTexto(font, txt); T(t, rx - font.widthOfTextAtSize(t, size), y, size, font); };
  const Tcen = (txt, cx, y, size, font = reg) => { const t = filtrarTexto(font, txt); T(t, cx - font.widthOfTextAtSize(t, size) / 2, y, size, font); };

  T(numeroCert, L.numero.x, L.numero.y, L.size);
  const nom = filtrarTexto(reg, d.nombre);
  T(nom, L.nombre.x, L.nombre.y, ajustar(reg, nom, L.size, L.nombre.maxW));
  T(d.doc, L.doc.x, L.doc.y, L.size);
  if (L.direccion) { const dir = filtrarTexto(reg, d.direccion || ''); T(dir, L.direccion.x, L.direccion.y, ajustar(reg, dir, L.direccion.size, L.direccion.maxW)); }
  d.lineas.slice(0, MAX_LINEAS).forEach((l, i) => {
    const y = L.filas.ys[i];
    Tcen(String(l.cantidad), L.filas.cantCx, y, L.size);
    const desc = filtrarTexto(reg, mesesTexto(l.meses, l.anio, P.concepto, reg, L.size, L.filas.descMaxW));
    T(desc, L.filas.descX, y, ajustar(reg, desc, L.size, L.filas.descMaxW));
    T('S/', L.filas.monedaX, y, L.size);
    Tder(importe(l.valor), L.filas.valorRx, y, L.size);
  });
  T('S/', L.filas.monedaX, L.total.y, L.total.size);
  Tder(importe(d.total), L.total.rx, L.total.y, L.total.size);
  T(`${P.ciudad} ${ftexto(P.emision)}`, L.ciudad.x, L.ciudad.y, L.size, neg);
  return doc.save({ useObjectStreams: true });
}

/* ------------------------------------------------------------------ interfaz */
function pintarArchivos() {
  const tb = $('#tabla-archivos tbody');
  tb.innerHTML = S.archivos.map((a) => `<tr class="${a.tipo === 'DESCONOCIDO' ? 'mal' : ''}"><td>${esc(a.nombre)}</td><td>${esc(a.tipoTexto)}</td>
    <td class="num">${a.filas.toLocaleString('es-PE')}</td><td class="num">${['NIUBIZ', 'AFRUS', 'MC'].includes(a.tipo) ? a.exitosos.toLocaleString('es-PE') : '—'}</td><td>${esc(a.mensaje)}</td></tr>`).join('');
  $('#resumen-archivos').textContent = S.archivos.length
    ? `Pagos leídos: ${S.pagos.size.toLocaleString('es-PE')} (duplicados descartados: ${S.duplicados}). Personas en la base de donantes: ${S.base.size.toLocaleString('es-PE')}.`
    : '';
  const hayPagos = [...S.pagos.values()].some((p) => p.ok);
  $('#btn-revisar').disabled = !hayPagos;
}

function filtrados() {
  const f = $('#filtro').value, q = sinTildes($('#buscar').value);
  return S.donantes.filter((d) => (f === 'TODOS' || (f === 'EXCLUIDOS' ? !d.incluir : d.estado === f && d.incluir)) &&
    (!q || sinTildes(d.nombre + ' ' + d.doc + ' ' + d.correo).includes(q)));
}

function pintarRevision() {
  const c = { LISTO: 0, REVISAR: 0, EMITIDO: 0 }; let excl = 0, totalListo = 0;
  for (const d of S.donantes) { if (!d.incluir) excl++; else { c[d.estado]++; if (d.estado === 'LISTO') totalListo += d.total; } }
  const f = $('#filtro').value;
  const tarj = (val, n, txt, cls) => `<button class="metrica ${cls} ${f === val ? 'sel' : ''}" data-f="${val}"><b>${n.toLocaleString('es-PE')}</b><span>${txt}</span></button>`;
  $('#contadores').innerHTML = tarj('LISTO', c.LISTO, `listos · ${soles(totalListo)}`, 'ok') + tarj('REVISAR', c.REVISAR, 'necesitan tu ayuda', c.REVISAR ? 'warn' : '') +
    tarj('EMITIDO', c.EMITIDO, 'ya tienen certificado', '') + tarj('EXCLUIDOS', excl, 'dejados fuera', '');
  $('#ayuda-filtro').hidden = !(f === 'REVISAR' && c.REVISAR);
  $('#ayuda-filtro').innerHTML = 'A estas personas les falta algo (lo dice la columna <b>Qué falta</b>). Corrígelo en la misma fila (por ejemplo, escribe el correo) y pasarán a <b>Listos</b>. Si no puedes resolverlo, desmarca <b>Incluir</b>: no recibirán certificado por ahora.';
  const lista = filtrados();
  const POR = 50; const pags = Math.max(1, Math.ceil(lista.length / POR));
  S.pagina = Math.min(S.pagina, pags - 1);
  const vis = lista.slice(S.pagina * POR, S.pagina * POR + POR);
  $('#tabla-donantes tbody').innerHTML = vis.map((d) => {
    const i = S.donantes.indexOf(d);
    const correoCampo = d.correos.length > 1 && !d.editadoCorreo
      ? `<select data-i="${i}" class="ed-correo-sel"><option value="">— elegir —</option>${d.correos.map((x) => `<option ${x === d.correo ? 'selected' : ''}>${esc(x)}</option>`).join('')}</select>`
      : `<input data-i="${i}" class="ed-correo" value="${esc(d.correo)}" placeholder="correo@ejemplo.com">`;
    return `<tr class="${!d.incluir ? 'excluido' : d.estado === 'LISTO' ? 'listo' : 'revisar'}">
      <td><input type="checkbox" data-i="${i}" class="ed-incluir" ${d.incluir ? 'checked' : ''} title="Incluir"></td>
      <td><span class="estado ${d.estado}">${d.estado === 'LISTO' ? 'Listo' : d.estado === 'EMITIDO' ? 'Ya tiene' : 'Falta algo'}</span></td>
      <td><input data-i="${i}" class="ed-nombre" value="${esc(d.nombre)}"></td>
      <td>${esc(d.doc)}</td><td>${correoCampo}</td>
      <td class="num">${d.pagos.length}</td><td class="num">${soles(d.total)}</td>
      <td class="small" style="white-space:nowrap">${d.lineas.map((l) => `${l.cantidad} × ${soles(l.valor)}`).join('<br>')}<br><span class="muted">${esc(d.plataformas)}</span></td>
      <td class="small">${d.tipo === 'empresas' ? `<input data-i="${i}" class="ed-dir" value="${esc(d.direccion)}" placeholder="Dirección de la empresa"><br>` : ''}${d.obs.map(esc).join('<br>')}${d.nota ? `<br><span class="muted">${esc(d.nota)}</span>` : ''}</td>
      <td><button class="sec ver" data-i="${i}" style="padding:5px 10px">Ver PDF</button></td></tr>`;
  }).join('') || '<tr><td colspan="10" class="muted">No hay registros con este filtro.</td></tr>';
  $('#paginas').innerHTML = pags > 1 ? Array.from({ length: pags }, (_, k) => `<button class="pag ${k === S.pagina ? 'act' : ''}" data-p="${k}">${k + 1}</button>`).join('') : '';
  $('#sin-doc').innerHTML = S.sinDoc.length ? `<details><summary>${S.sinDoc.length} pagos exitosos no tienen DNI y no se pudieron asignar a nadie (no se incluyen)</summary>
    <table><thead><tr><th>Plataforma</th><th>Fecha</th><th>Monto</th><th>Nombre</th><th>Correo</th></tr></thead><tbody>${S.sinDoc.slice(0, 300).map((p) => `<tr><td>${p.plataforma}</td><td>${p.f ? ftexto(p.f) : ''}</td><td>${soles(p.monto)}</td><td>${esc(p.nombre)}</td><td>${esc(p.correo)}</td></tr>`).join('')}</tbody></table></details>` : '';
  actualizarPaso3();
}

function actualizarPaso3() {
  const listos = S.donantes.filter((d) => d.incluir && d.estado === 'LISTO').length;
  const plantillasOk = S.plantillas.personas && (S.plantillas.empresas || !S.donantes.some((d) => d.incluir && d.estado === 'LISTO' && d.tipo === 'empresas'));
  const sig = siguienteNumero();
  const muestras = S.muestrasVistas.size;
  $('#estado-muestras').innerHTML = [0, 1, 2].map((k) => `<span class="punto ${k < muestras ? 'ok' : ''}">${k < muestras ? '✓' : k + 1}</span>`).join('') + `<span class="muted small" style="margin-left:6px">${Math.min(muestras, 3)} de 3 ejemplos vistos</span>`;
  const numOk = $('#confirmo-num').checked;
  const ok = listos > 0 && plantillasOk && sig !== null && muestras >= 3 && numOk && $('#confirmo').checked;
  $('#btn-generar').disabled = !ok;
  $('#btn-generar').textContent = listos ? `Crear ${listos} y descargar ZIP` : 'Crear y descargar ZIP';
  const sp = $('#btn-generar-sp');
  if (sp) { sp.disabled = !ok; sp.textContent = listos ? `Crear ${listos.toLocaleString('es-PE')} certificados` : 'Crear certificados'; }
  $('#caja-plantillas').hidden = !!plantillasOk;
  $('#estado-plantillas').innerHTML = `<b>Falta la plantilla de ${S.plantillas.personas ? 'empresas' : 'personas'}.</b> Normalmente se lee sola de la carpeta 05_HERRAMIENTA; si no está ahí, elígela aquí.`;
  $('#resumen-crear').innerHTML = listos ? `Se crearán <b>${listos.toLocaleString('es-PE')} certificados</b>${sig !== null ? `, del N.º <b>${sig}</b> al <b>${sig + listos - 1}</b>` : ''}.` : 'No hay donantes listos. Vuelve al paso 2.';
  if (sig !== null && listos) $('#rango-numeros').textContent = `→ hasta el ${SERIE}${sig + listos - 1}`;
  else $('#rango-numeros').textContent = '';
  $('#texto-confirmo-num').innerHTML = sig !== null ? `Revisé el Excel oficial: el N.º <b>${SERIE}${sig}</b> está libre (no tiene nombre en la hoja REPORTE).` : 'Escribe arriba el número del primer certificado.';
  const motivo = !listos ? 'No hay donantes listos.' : !plantillasOk ? 'Falta una plantilla.' : sig === null ? 'Falta el número del primer certificado.' : muestras < 3 ? `Faltan ${3 - muestras} ejemplos por mirar (paso 2).` : !numOk ? 'Confirma el número en el Excel.' : !$('#confirmo').checked ? 'Marca la casilla de confirmación.' : '';
  $('#motivo-3').textContent = motivo;
  if (typeof uxActualizar === 'function') uxActualizar();
}

function siguienteNumero() {
  const v = limpiar($('#p-numero').value);
  if (!v) return null;
  const n = parseInt(v.split(',').pop(), 10);
  return Number.isFinite(n) && n > 0 ? n : null;
}

async function plantillaDe(tipo) {
  if (!S.plantillas[tipo]) throw new Error('Falta la plantilla de ' + tipo);
  return PDFDocument.load(S.plantillas[tipo]);
}

async function verMuestra(i) {
  const d = S.donantes[i];
  if (!S.plantillas[d.tipo]) {
    aviso(`Falta la plantilla de ${d.tipo} (plantilla_${d.tipo}.pdf). Debe estar en la carpeta 05_HERRAMIENTA de SharePoint.`, 'warn');
    return;
  }
  try {
    const bytes = await pdfCertificado(d, SERIE + 'MUESTRA', params(), await plantillaDe(d.tipo));
    const url = URL.createObjectURL(new Blob([bytes], { type: 'application/pdf' }));
    const sinVisor = navigator.pdfViewerEnabled === false;
    $('#visor').hidden = sinVisor; $('#visor-alt').hidden = !sinVisor;
    if (!sinVisor) $('#visor').src = url;
    const dl = $('#descargar-muestra'); dl.href = url; dl.download = `MUESTRA - ${d.nombre || d.doc}.pdf`;
    $('#modal').hidden = false;
    S.muestrasVistas.add(d.doc); actualizarPaso3();
  } catch (e) { aviso('No se pudo crear la muestra: ' + e.message, 'error'); }
}

function csv(filas) {
  return '﻿' + filas.map((f) => f.map((c) => {
    const s = String(c == null ? '' : c);
    return /[;"\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }).join(';')).join('\r\n');
}

const fechaIso = (f) => `${f.y}-${String(f.m).padStart(2, '0')}-${String(f.d).padStart(2, '0')}`;

/* Crea todos los PDF y archivos de control de un lote (no descarga ni sube nada). */
async function prepararLote() {
  const P = params();
  const listos = S.donantes.filter((d) => d.incluir && d.estado === 'LISTO');
  const n = siguienteNumero();
  const plant = { personas: await plantillaDe('personas'), empresas: S.plantillas.empresas ? await plantillaDe('empresas') : null };
  const manifiesto = [['numero_certificado', 'tipo', 'documento', 'nombre', 'correo', 'periodo_desde', 'periodo_hasta', 'cantidad_pagos', 'monto_total', 'moneda', 'archivo', 'fecha_emision', 'plataformas']];
  const sustento = [['numero_certificado', 'documento', 'plataforma', 'id_pago', 'fecha_pago', 'monto', 'estado_original', 'archivo_origen']];
  const reporte = [];
  const pdfs = [];
  $('#progreso').hidden = false;
  for (let k = 0; k < listos.length; k++) {
    const d = listos[k];
    const num = SERIE + (n + k);
    const archivo = $('#p-docnombre').checked ? `CERTIFICADO DE DONACIÓN - ${num} - ${d.doc}.pdf` : `CERTIFICADO DE DONACIÓN - ${num}.pdf`;
    const bytes = await pdfCertificado(d, num, P, plant[d.tipo]);
    pdfs.push({ archivo, bytes, num, d });
    manifiesto.push([num, d.tipo === 'empresas' ? 'EMPRESA' : 'PERSONA', d.doc, d.nombre, d.correo, fechaIso(P.desde), fechaIso(P.hasta), d.pagos.length, d.total.toFixed(2), 'PEN', archivo, fechaIso(P.emision), d.plataformas]);
    for (const p of d.pagos) sustento.push([num, d.doc, p.plataforma, p.id, fechaIso(p.f), p.monto.toFixed(2), p.estado, p.archivo]);
    const fila = [d.tipo === 'empresas' ? 'Donación EMPRESAS' : 'Donación PERSONAS', num, P.inmobiliaria, d.nombre, d.doc, d.tipo === 'empresas' ? d.direccion : '-'];
    for (let j = 0; j < MAX_LINEAS; j++) {
      const l = d.lineas[j];
      fila.push(l ? mesesTexto(l.meses, l.anio, P.concepto, S.metric, LAYOUT[d.tipo].size, LAYOUT[d.tipo].filas.descMaxW) : '', l ? l.cantidad : '', l ? l.valor.toFixed(2) : '');
    }
    fila.push('SOLES', d.total.toFixed(2), P.ciudad, ftexto(P.emision));
    reporte.push(fila);
    if (k % 5 === 0 || k === listos.length - 1) {
      $('#barra').value = Math.round(((k + 1) / listos.length) * 100);
      $('#progreso-texto').textContent = `Generando ${k + 1} de ${listos.length}…`;
      await new Promise((r) => setTimeout(r, 0));
    }
  }
  const marca = new Date().toISOString().slice(0, 16).replace(/[-:T]/g, '');
  const nombreLote = `lote_${marca}_${n}-${n + listos.length - 1}`;
  const control = {
    'manifiesto_envio.csv': csv(manifiesto),
    'sustento_pagos.csv': csv(sustento),
    'filas_para_REPORTE.txt': '﻿' + reporte.map((f) => f.join('\t')).join('\r\n'),
    'LEEME.txt': `Lote generado el ${new Date().toLocaleString('es-PE')}\r\nCertificados: ${listos.length}\r\nNúmeros: ${SERIE}${n} a ${SERIE}${n + listos.length - 1}\r\nPeríodo: ${fechaIso(P.desde)} a ${fechaIso(P.hasta)}\r\n\r\n` +
      `1. filas_para_REPORTE.txt: pegar en la hoja REPORTE del libro oficial, columna B de la fila del número ${SERIE}${n}.\r\n` +
      `2. manifiesto_envio.csv: lista para el envío de correos.\r\n3. sustento_pagos.csv: qué pagos respaldan cada certificado.\r\n`,
  };
  return { listos, n, pdfs, control, nombreLote, manifiesto };
}

/* Después de generar: bloquear a esos donantes y avanzar la numeración. */
function cerrarLote(L) {
  L.listos.forEach((d, k) => { d.emitido = SERIE + (L.n + k); S.emitidos.set(d.doc, d.emitido); d.incluir = false; validar(d); });
  $('#confirmo').checked = false; $('#confirmo-num').checked = false;
  S.muestrasVistas.clear();
  $('#p-numero').value = String(L.n + L.listos.length);
  guardarAjustes();
  pintarRevision();
}

/* Modo sin conexión: descarga un ZIP. */
async function generar() {
  const L = await prepararLote();
  const zip = new JSZip();
  const carpeta = zip.folder('certificados');
  for (const p of L.pdfs) carpeta.file(p.archivo, p.bytes);
  for (const [nom, txt] of Object.entries(L.control)) zip.file(nom, txt);
  const blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `certificados_${L.nombreLote}.zip`;
  document.body.appendChild(a); a.click(); a.remove();
  $('#progreso-texto').textContent = `Listo: ${L.listos.length} certificados generados. Se descargó el archivo ${a.download}.`;
  cerrarLote(L);
  irA(4);
  aviso(`Se generaron ${L.listos.length} certificados. El siguiente número disponible ahora es ${SERIE}${L.n + L.listos.length}.`, 'ok');
}

/* ------------------------------------------------------------------ almacenamiento local */
function guardarAjustes() {
  try {
    const a = {}; for (const id of ['p-desde', 'p-hasta', 'p-ciudad', 'p-concepto', 'p-inmob', 'p-numero']) a[id] = $('#' + id).value;
    localStorage.setItem('cd_ajustes', JSON.stringify(a));
  } catch (e) { /* almacenamiento no disponible */ }
}
function b64(bytes) { let s = ''; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000)); return btoa(s); }
function deb64(s) { const b = atob(s); const u = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); return u; }

async function cargarPlantilla(file, tipo) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const pdf = await PDFDocument.load(bytes);
  const p = pdf.getPage(0); const { width, height } = p.getSize();
  if (Math.abs(width - PAGE_W) > 2 || Math.abs(height - PAGE_H) > 2) throw new Error('La plantilla no tiene tamaño A4.');
  S.plantillas[tipo] = bytes;
  try { localStorage.setItem('cd_plantilla_' + tipo, b64(bytes)); } catch (e) { aviso('No se pudo recordar la plantilla en este navegador; habrá que cargarla otra vez la próxima vez.', 'warn'); }
}

/* ------------------------------------------------------------------ inicio */
async function iniciar() {
  const hoy = new Date();
  $('#p-emision').value = hoy.toISOString().slice(0, 10);
  try {
    const a = JSON.parse(localStorage.getItem('cd_ajustes') || '{}');
    for (const [k, v] of Object.entries(a)) if ($('#' + k) && v) $('#' + k).value = v;
    for (const t of ['personas', 'empresas']) { const s = localStorage.getItem('cd_plantilla_' + t); if (s) S.plantillas[t] = deb64(s); }
  } catch (e) { /* sin almacenamiento */ }
  const tmp = await PDFDocument.create();
  S.metric = await tmp.embedFont(StandardFonts.Helvetica);

  const zona = $('#zona');
  const recibir = async (files) => {
    $('#cargando').hidden = false;
    for (const f of files) {
      try {
        if (/\.pdf$/i.test(f.name)) {
          const tipo = /empresa/i.test(f.name) ? 'empresas' : 'personas';
          await cargarPlantilla(f, tipo);
          continue;
        }
        const info = await cargarArchivo(f);
        if (info.tipo === 'DESCONOCIDO') aviso(`${f.name}: no se reconoció este archivo y no se usa.`, 'error');
      } catch (e) { aviso(`${f.name}: no se pudo leer (${e.message}).`, 'error'); }
    }
    $('#cargando').hidden = true;
    if (S.maxEmitido && !$('#p-numero').value) $('#p-numero').value = String(S.maxEmitido + 1);
    pintarArchivos(); actualizarPaso3();
  };
  zona.addEventListener('dragover', (e) => { e.preventDefault(); zona.classList.add('sobre'); });
  zona.addEventListener('dragleave', () => zona.classList.remove('sobre'));
  zona.addEventListener('drop', (e) => { e.preventDefault(); zona.classList.remove('sobre'); recibir([...e.dataTransfer.files]); });
  $('#input-archivos').addEventListener('change', (e) => { recibir([...e.target.files]); e.target.value = ''; });

  $('#btn-revisar').addEventListener('click', () => {
    consolidar();
    $('#filtro').value = S.donantes.some((d) => d.incluir && d.estado === 'REVISAR') ? 'REVISAR' : 'LISTO';
    S.pagina = 0;
    pintarRevision();
    irA(2);
  });
  $('#contadores').addEventListener('click', (e) => { const b = e.target.closest('[data-f]'); if (b) { $('#filtro').value = b.dataset.f; S.pagina = 0; pintarRevision(); } });
  $('#btn-reiniciar').addEventListener('click', () => { if (confirm('¿Borrar todos los archivos cargados y empezar de nuevo? (las plantillas y ajustes se conservan)')) location.reload(); });
  for (const id of ['filtro', 'buscar']) $('#' + id).addEventListener('input', () => { S.pagina = 0; pintarRevision(); });
  $('#paginas').addEventListener('click', (e) => { if (e.target.dataset.p) { S.pagina = +e.target.dataset.p; pintarRevision(); } });
  $('#tabla-donantes').addEventListener('change', (e) => {
    const i = e.target.dataset.i; if (i === undefined) return;
    const d = S.donantes[+i];
    if (e.target.classList.contains('ed-incluir')) d.incluir = e.target.checked;
    if (e.target.classList.contains('ed-nombre')) { d.nombre = limpiar(e.target.value); d.editadoNombre = true; }
    if (e.target.classList.contains('ed-correo') || e.target.classList.contains('ed-correo-sel')) { d.correo = limpiar(e.target.value).toLowerCase(); d.editadoCorreo = !!d.correo; }
    if (e.target.classList.contains('ed-dir')) d.direccion = limpiar(e.target.value);
    validar(d); pintarRevision();
  });
  $('#tabla-donantes').addEventListener('click', (e) => { if (e.target.classList.contains('ver')) verMuestra(+e.target.dataset.i); });
  $('#cerrar-modal').addEventListener('click', () => { $('#modal').hidden = true; });
  for (const id of ['p-desde', 'p-hasta', 'p-concepto', 'p-ciudad', 'p-inmob', 'p-numero', 'p-emision']) $('#' + id).addEventListener('change', () => {
    guardarAjustes();
    if (S.donantes.length && ['p-desde', 'p-hasta'].includes(id)) { consolidar(); pintarRevision(); } else actualizarPaso3();
  });
  $('#confirmo').addEventListener('change', actualizarPaso3);
  $('#confirmo-num').addEventListener('change', actualizarPaso3);
  $('#p-numero').addEventListener('input', () => { $('#confirmo-num').checked = false; actualizarPaso3(); });
  $('#input-plantillas').addEventListener('change', async (e) => {
    for (const f of e.target.files) {
      try {
        const tipo = /empresa/i.test(f.name) ? 'empresas' : 'personas';
        await cargarPlantilla(f, tipo);
        aviso(`Plantilla de ${tipo} cargada (${f.name}).`, 'ok');
      } catch (err) { aviso(`${f.name}: no es una plantilla válida (${err.message}).`, 'error'); }
    }
    e.target.value = '';
    actualizarPaso3();
  });
  $('#btn-generar').addEventListener('click', async () => {
    $('#btn-generar').disabled = true;
    try { await generar(); } catch (e) { aviso('Error al generar: ' + e.message, 'error'); }
    actualizarPaso3();
  });
  actualizarPaso3();
}
window.addEventListener('DOMContentLoaded', iniciar);
