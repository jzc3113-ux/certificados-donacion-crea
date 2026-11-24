/* Conexión con Microsoft 365: inicio de sesión (OAuth 2.0 + PKCE), SharePoint y envío de correos.
   Todo se ejecuta en el navegador con la cuenta de quien inicia sesión; no hay servidor ni claves. */
'use strict';

const CFG_DEFECTO = {
  clientId: '', tenantId: '',
  host: 'creamas.sharepoint.com', sitio: '/sites/Recaudacion2026DataLake', biblioteca: 'inbox', carpeta: 'Certificados 2026',
  buzon: 'amigos@creamas.org', pausaMs: 2500,
};
const RUTAS = {
  reportes: '01_REPORTES', base: '02_BASE_DONANTES', control: '03_CONTROL', lotes: '04_LOTES', herramienta: '05_HERRAMIENTA',
  registro: '03_CONTROL/Registro certificados 2026.xlsx',
};
const REG_COLS = ['fecha_registro', 'lote', 'numero_certificado', 'documento', 'nombre', 'correo', 'monto_total', 'archivo', 'estado', 'fecha_envio', 'detalle'];
const REGISTRO_PLANTILLA_B64 = '/*__REGISTRO_B64__*/';
const GRAPH = 'https://graph.microsoft.com/v1.0';

const M = { cfg: { ...CFG_DEFECTO }, token: null, yo: null, siteId: null, driveBase: null, lotes: [], registro: null, enviando: false };

function cargarCfg() {
  try { Object.assign(M.cfg, JSON.parse(localStorage.getItem('cd_m365') || '{}')); } catch (e) { /* sin almacenamiento */ }
  if (window.CD_CFG) for (const [k, v] of Object.entries(window.CD_CFG)) if (v && !M.cfg[k]) M.cfg[k] = v;
}
function guardarCfg() { try { localStorage.setItem('cd_m365', JSON.stringify(M.cfg)); } catch (e) { /* nada */ } }
const urlRetorno = () => location.origin + location.pathname.replace(/index\.html$/, '');

/* ------------------------------------------------------------ inicio de sesión (PKCE) */
const b64url = (buf) => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
function aleatorio(n = 48) { const a = new Uint8Array(n); crypto.getRandomValues(a); return b64url(a); }
const ALCANCE = 'openid profile offline_access https://graph.microsoft.com/.default';

async function iniciarSesion() {
  if (!M.cfg.clientId || !M.cfg.tenantId) { aviso('Falta configurar el Id. de aplicación y el Id. de directorio (Configuración TI).', 'error'); return; }
  const verificador = aleatorio(48);
  const reto = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verificador)));
  const estado = aleatorio(16);
  sessionStorage.setItem('cd_pkce', JSON.stringify({ verificador, estado }));
  const q = new URLSearchParams({
    client_id: M.cfg.clientId, response_type: 'code', redirect_uri: urlRetorno(), response_mode: 'query',
    scope: ALCANCE, code_challenge: reto, code_challenge_method: 'S256', state: estado, prompt: 'select_account',
  });
  location.assign(`https://login.microsoftonline.com/${M.cfg.tenantId}/oauth2/v2.0/authorize?${q}`);
}

async function pedirToken(cuerpo) {
  const r = await fetch(`https://login.microsoftonline.com/${M.cfg.tenantId}/oauth2/v2.0/token`, {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: M.cfg.clientId, scope: ALCANCE, ...cuerpo }),
  });
  const j = await r.json();
  if (!r.ok) throw new Error(j.error_description || j.error || 'No se pudo iniciar sesión');
  M.token = { acceso: j.access_token, renovar: j.refresh_token, vence: Date.now() + (j.expires_in - 300) * 1000 };
  sessionStorage.setItem('cd_token', JSON.stringify(M.token));
}

async function completarInicio() {
  const q = new URLSearchParams(location.search);
  if (q.get('error')) { aviso('Inicio de sesión cancelado: ' + (q.get('error_description') || q.get('error')), 'error'); history.replaceState(null, '', urlRetorno()); return; }
  if (q.get('code')) {
    const pk = JSON.parse(sessionStorage.getItem('cd_pkce') || '{}');
    history.replaceState(null, '', urlRetorno());
    if (!pk.estado || pk.estado !== q.get('state')) { aviso('Respuesta de inicio de sesión no válida. Intenta de nuevo.', 'error'); return; }
    await pedirToken({ grant_type: 'authorization_code', code: q.get('code'), redirect_uri: urlRetorno(), code_verifier: pk.verificador });
    sessionStorage.removeItem('cd_pkce');
  } else {
    try { M.token = JSON.parse(sessionStorage.getItem('cd_token') || 'null'); } catch (e) { M.token = null; }
  }
}

async function tokenVigente() {
  if (!M.token) throw new Error('No has iniciado sesión.');
  if (Date.now() > M.token.vence) await pedirToken({ grant_type: 'refresh_token', refresh_token: M.token.renovar });
  return M.token.acceso;
}

function cerrarSesion() { sessionStorage.removeItem('cd_token'); M.token = null; M.yo = null; pintarSesion(); }

/* ------------------------------------------------------------ Graph */
async function graph(metodo, ruta, cuerpo, tipo = 'application/json', intento = 0) {
  const h = { Authorization: 'Bearer ' + await tokenVigente() };
  if (cuerpo !== undefined) h['Content-Type'] = tipo;
  const r = await fetch(ruta.startsWith('http') ? ruta : GRAPH + ruta, {
    method: metodo, headers: h, body: cuerpo === undefined ? undefined : (tipo === 'application/json' ? JSON.stringify(cuerpo) : cuerpo),
  });
  if ((r.status === 429 || r.status === 503 || r.status === 504) && intento < 5) {
    const espera = (parseInt(r.headers.get('Retry-After') || '5', 10) || 5) * 1000;
    await new Promise((ok) => setTimeout(ok, espera));
    return graph(metodo, ruta, cuerpo, tipo, intento + 1);
  }
  if (r.status === 204 || r.status === 202) return null;
  const txt = await r.text();
  let j = null; try { j = txt ? JSON.parse(txt) : null; } catch (e) { j = { texto: txt }; }
  if (!r.ok) { const e = new Error((j && j.error && j.error.message) || `Error ${r.status}`); e.status = r.status; e.code = j && j.error && j.error.code; throw e; }
  return j;
}
const enc = (ruta) => ruta.split('/').map(encodeURIComponent).join('/');
const rutaDrive = (rel) => `${M.driveBase}/root:/${enc(M.cfg.carpeta + (rel ? '/' + rel : ''))}`;

async function sitio() {
  if (M.siteId && M.driveBase) return M.siteId;
  const s = await graph('GET', `/sites/${M.cfg.host}:${M.cfg.sitio}`);
  M.siteId = s.id;
  let drive;
  if (M.cfg.biblioteca) {
    const lista = (await graph('GET', `/sites/${M.siteId}/drives?$select=id,name,webUrl`)).value || [];
    const b = M.cfg.biblioteca.toLowerCase();
    drive = lista.find((d) => (d.name || '').toLowerCase() === b || decodeURIComponent(d.webUrl || '').toLowerCase().endsWith('/' + b));
    if (!drive) throw new Error(`No existe la biblioteca "${M.cfg.biblioteca}" en el sitio. Bibliotecas disponibles: ${lista.map((d) => d.name).join(', ')}`);
  } else drive = await graph('GET', `/sites/${M.siteId}/drive`);
  M.driveBase = `/drives/${drive.id}`;
  return M.siteId;
}

async function listar(rel) {
  const out = [];
  let url = `${rutaDrive(rel)}:/children?$top=999`;
  try {
    while (url) {
      const j = await graph('GET', url);
      for (const it of j.value) out.push({ ...it, rel: (rel ? rel + '/' : '') + it.name });
      url = j['@odata.nextLink'] || null;
    }
  } catch (e) { if (e.status === 404) return []; throw e; }
  return out;
}
async function listarTodo(rel) {
  const res = [];
  for (const it of await listar(rel)) {
    if (it.folder) res.push(...await listarTodo(it.rel)); else res.push(it);
  }
  return res;
}
async function descargar(it) {
  const url = it['@microsoft.graph.downloadUrl'];
  const r = url ? await fetch(url) : await fetch(GRAPH + `${M.driveBase}/items/${it.id}/content`, { headers: { Authorization: 'Bearer ' + await tokenVigente() } });
  if (!r.ok) throw new Error(`No se pudo descargar ${it.name}`);
  return new Uint8Array(await r.arrayBuffer());
}
async function subir(rel, bytes, sobrescribir = false) {
  return graph('PUT', `${rutaDrive(rel)}:/content?@microsoft.graph.conflictBehavior=${sobrescribir ? 'replace' : 'fail'}`, bytes, 'application/octet-stream');
}
const comoArchivo = (nombre, bytes) => ({ name: nombre, arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) });

/* ------------------------------------------------------------ registro (Excel en 03_CONTROL) */
async function abrirRegistro() {
  let item;
  try { item = await graph('GET', rutaDrive(RUTAS.registro)); }
  catch (e) {
    if (e.status !== 404) throw e;
    item = await subir(RUTAS.registro, deb64(REGISTRO_PLANTILLA_B64));
    aviso('Se creó el archivo "Registro certificados 2026.xlsx" en 03_CONTROL.', 'ok');
  }
  const base = `${M.driveBase}/items/${item.id}/workbook/tables/Envios`;
  const filas = (await graph('GET', `${base}/rows`)).value || [];
  const porNumero = new Map();
  const pruebaLotes = new Set();
  for (const f of filas) {
    const v = f.values[0];
    const reg = Object.fromEntries(REG_COLS.map((c, i) => [c, v[i]]));
    if (!reg.numero_certificado) continue;
    if (reg.estado === 'PRUEBA_ENVIADA') pruebaLotes.add(reg.lote);
    const previo = porNumero.get(reg.numero_certificado);
    // Prevalece el estado más avanzado: ENVIADO > ENVIANDO > ERROR > PRUEBA
    const rango = { REENVIADO: 5, ENVIADO: 4, ENVIANDO: 3, ERROR: 2, PRUEBA_ENVIADA: 1 };
    if (!previo || (rango[reg.estado] || 0) >= (rango[previo.estado] || 0)) porNumero.set(reg.numero_certificado, { ...reg, index: f.index });
  }
  M.registro = { base, porNumero, pruebaLotes };
  return M.registro;
}
async function registrar(valores) {
  const j = await graph('POST', `${M.registro.base}/rows`, { values: [REG_COLS.map((c) => valores[c] == null ? '' : String(valores[c]))] });
  return j.index;
}
async function actualizarFila(index, valores) {
  await graph('PATCH', `${M.registro.base}/rows/itemAt(index=${index})`, { values: [REG_COLS.map((c) => valores[c] == null ? '' : String(valores[c]))] });
}

/* ------------------------------------------------------------ preparar la carpeta (estructura fija) */
const ESTRUCTURA = ['01_REPORTES', '01_REPORTES/Niubiz', '01_REPORTES/AFRUS', '01_REPORTES/MC_Izipay', '01_REPORTES/Manuales', '01_REPORTES/Exclusiones',
  '02_BASE_DONANTES', '03_CONTROL', '04_LOTES', '05_HERRAMIENTA', '06_MANUAL'];

async function crearCarpeta(rel) {
  const partes = (M.cfg.carpeta + (rel ? '/' + rel : '')).split('/');
  const nombre = partes.pop();
  const padre = partes.length ? `${M.driveBase}/root:/${enc(partes.join('/'))}:/children` : `${M.driveBase}/root/children`;
  try { await graph('POST', padre, { name: nombre, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }); return true; }
  catch (e) { if (e.status === 409) return false; throw e; }
}

async function prepararCarpeta() {
  await sitio();
  let creadas = 0;
  if (await crearCarpeta('')) creadas++;
  for (const r of ESTRUCTURA) if (await crearCarpeta(r)) creadas++;
  let subidas = 0;
  for (const t of ['personas', 'empresas']) {
    if (!S.plantillas[t]) continue;
    try { await subir(`${RUTAS.herramienta}/plantilla_${t}.pdf`, S.plantillas[t]); subidas++; }
    catch (e) { if (e.status !== 409) throw e; }
  }
  const raiz = await graph('GET', rutaDrive(''));
  aviso(`Carpeta lista: ${creadas} carpetas nuevas${subidas ? `, ${subidas} plantilla(s) subidas` : ''}.`, 'ok');
  $('#enlace-carpeta').innerHTML = `Carpeta: <a href="${esc(raiz.webUrl)}" target="_blank" rel="noopener">${esc(M.cfg.carpeta)}</a>`;
  return raiz.webUrl;
}

/* ------------------------------------------------------------ actualizar desde SharePoint */
async function actualizarDesdeSharePoint() {
  if (M.leyendo) return;
  M.leyendo = true;
  $('#cargando').hidden = false;
  $('#estado-general').innerHTML = '';
  try {
    await sitio();
    const grupos = [RUTAS.reportes, RUTAS.base, RUTAS.control, RUTAS.herramienta];
    let n = 0;
    for (const g of grupos) {
      for (const it of await listarTodo(g)) {
        if (/^~\$/.test(it.name) || /Registro certificados 2026\.xlsx$/i.test(it.name)) continue;
        if (g === RUTAS.herramienta && !/^plantilla_(personas|empresas)\.pdf$/i.test(it.name)) continue;
        if (!/\.(csv|xlsx|xlsm|pdf)$/i.test(it.name)) continue;
        if (S.archivos.some((a) => a.ruta === it.rel)) continue;
        const bytes = await descargar(it);
        if (/\.pdf$/i.test(it.name)) { await cargarPlantilla(comoArchivo(it.name, bytes), /empresa/i.test(it.name) ? 'empresas' : 'personas'); continue; }
        const info = await cargarArchivo(comoArchivo(it.name, bytes));
        info.ruta = it.rel;
        n++;
      }
    }
    await leerLotes();
    await abrirRegistro();
    M.leido = true;
  } catch (e) {
    aviso('No se pudo leer la carpeta de SharePoint: ' + e.message + ' Revisa tu conexión y pulsa "Leer la carpeta otra vez".', 'error');
  }
  M.leyendo = false;
  $('#cargando').hidden = true;
  if (S.maxEmitido && !$('#p-numero').value) $('#p-numero').value = String(S.maxEmitido + 1);
  pintarArchivos(); actualizarPaso3(); pintarLotes();
}

/* Lotes anteriores: cada carpeta de 04_LOTES con su manifiesto. Bloquean reemisión y permiten reanudar envíos. */
async function leerLotes() {
  M.lotes = [];
  for (const c of await listar(RUTAS.lotes)) {
    if (!c.folder) continue;
    const hijos = await listar(c.rel);
    const man = hijos.find((h) => h.name === 'manifiesto_envio.csv');
    if (!man) continue;
    const texto = await decodificar(await descargar(man));
    const filas = [];
    let cab = null;
    recorrerCSV(texto, (r) => { if (!cab) { cab = r.map(sinTildes); return; } filas.push(Object.fromEntries(cab.map((k, i) => [k, limpiar(r[i])]))); });
    for (const f of filas) {
      S.emitidos.set(normDoc(f.documento), f.numero_certificado);
      if (f.numero_certificado.startsWith(SERIE)) { const k = parseInt(f.numero_certificado.slice(SERIE.length), 10); if (k > S.maxEmitido) S.maxEmitido = k; }
    }
    const pdfs = new Map((await listar(c.rel + '/certificados')).map((h) => [h.name, h]));
    M.lotes.push({ nombre: c.name, rel: c.rel, filas, pdfs });
  }
}

/* ------------------------------------------------------------ generar y guardar en SharePoint */
async function generarEnSharePoint() {
  await sitio();
  const L = await prepararLote();
  const rel = `${RUTAS.lotes}/${L.nombreLote}`;
  $('#progreso-texto').textContent = 'Guardando en SharePoint…';
  // Primero los archivos de control: si algo falla después, el lote ya bloquea la reemisión.
  for (const [nom, txt] of Object.entries(L.control)) await subir(`${rel}/${nom}`, new TextEncoder().encode(txt));
  for (let k = 0; k < L.pdfs.length; k++) {
    await subir(`${rel}/certificados/${L.pdfs[k].archivo}`, L.pdfs[k].bytes);
    $('#barra').value = Math.round(((k + 1) / L.pdfs.length) * 100);
    $('#progreso-texto').textContent = `Guardando PDF ${k + 1} de ${L.pdfs.length} en SharePoint…`;
  }
  cerrarLote(L);
  await leerLotes();
  if (!M.registro) await abrirRegistro();
  pintarLotes();
  $('#progreso').hidden = true;
  M.ultimoLote = L;
  $('#post-crear-titulo').textContent = `Listo: ${L.listos.length.toLocaleString('es-PE')} certificados creados (N.º ${L.n} al ${L.n + L.listos.length - 1})`;
  $('#post-crear-num').textContent = SERIE + L.n;
  $('#post-crear').hidden = false;
}

/* ------------------------------------------------------------ envío */
function pendientesDe(lote) {
  const reg = M.registro ? M.registro.porNumero : new Map();
  const pend = [], inciertos = [], enviados = [];
  for (const f of lote.filas) {
    const r = reg.get(f.numero_certificado);
    if (r && (r.estado === 'ENVIADO' || r.estado === 'REENVIADO')) enviados.push(f);
    else if (r && r.estado === 'ENVIANDO') inciertos.push(f);
    else pend.push(f);
  }
  return { pend, inciertos, enviados };
}

function cuerpoCorreo(f) {
  const nombre = esc(f.nombre);
  const anio = (f.periodo_desde || '').slice(0, 4) || '2026';
  return `<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:1.55;color:#1d2030;max-width:600px">
<p>Estimado/a ${nombre}:</p>
<p>Recibe un cordial saludo del equipo de Crea+.</p>
<p>Queremos agradecerte por la confianza y el compromiso que has demostrado al apoyar nuestras iniciativas educativas. Gracias a personas como tú, seguimos impulsando oportunidades de aprendizaje y desarrollo para niños, niñas y jóvenes.</p>
<p>Adjuntamos tu <b>certificado de donación ${esc(anio)}</b> (N.º ${esc(f.numero_certificado)}), emitido con la información validada de tus aportes. Este certificado comprende todas tus donaciones de ${esc(anio)} y reemplaza a cualquier certificado de este año que hayas recibido antes.</p>
<p>Te recomendamos conservar este documento para tus registros. Si tienes alguna consulta sobre la información consignada, puedes responder a este correo.</p>
<p>Gracias por ser parte de este esfuerzo compartido.</p>
<p>Con gratitud,<br><b>Equipo Crea+</b></p></div>`;
}

const PRUEBA_MAX = 3;
async function enviarLote(nombreLote, prueba) {
  if (M.enviando) return;
  const lote = M.lotes.find((l) => l.nombre === nombreLote);
  await abrirRegistro();
  let { pend, inciertos } = pendientesDe(lote);
  if (!pend.length) { aviso('Ya se enviaron todos los certificados de este grupo.', 'info'); return; }
  const destinoPrueba = M.yo && (M.yo.mail || M.yo.userPrincipalName);
  if (prueba) pend = pend.slice(0, PRUEBA_MAX);
  const minutos = Math.max(1, Math.round((pend.length * (M.cfg.pausaMs + 1500)) / 60000));
  const msg = prueba
    ? `Se enviarán ${pend.length} correos de PRUEBA a tu correo (${destinoPrueba}). Los donantes no reciben nada.`
    : `Vas a enviar ${pend.length} certificados a los DONANTES desde ${M.cfg.buzon}.\n\nTardará unos ${minutos} minutos. Una vez enviados, no se pueden recuperar.\n\n¿Enviar ahora?`;
  if (!confirm(msg)) return;
  M.enviando = true;
  const t = { ok: 0, error: 0 };
  $('#envio-progreso').hidden = false;
  for (let k = 0; k < pend.length; k++) {
    const f = pend[k];
    const base = { fecha_registro: new Date().toISOString(), lote: lote.nombre, numero_certificado: f.numero_certificado, documento: f.documento, nombre: f.nombre, correo: f.correo, monto_total: f.monto_total, archivo: f.archivo };
    let idx = null;
    try {
      if (!correoValido(f.correo)) throw new Error('Correo inválido');
      const it = lote.pdfs.get(f.archivo);
      if (!it) throw new Error('No se encontró el PDF en SharePoint');
      const bytes = await descargar(it);
      if (!bytes.length) throw new Error('PDF vacío');
      if (!prueba) idx = await registrar({ ...base, estado: 'ENVIANDO' }); // reserva antes de enviar
      await graph('POST', `/users/${encodeURIComponent(M.cfg.buzon)}/sendMail`, {
        message: {
          subject: `${prueba ? '[PRUEBA] ' : ''}Crea+ | Tu certificado de donación ${(f.periodo_desde || '2026').slice(0, 4)} N.º ${f.numero_certificado}`,
          body: { contentType: 'HTML', content: cuerpoCorreo(f) },
          toRecipients: [{ emailAddress: { address: prueba ? destinoPrueba : f.correo } }],
          attachments: [{ '@odata.type': '#microsoft.graph.fileAttachment', name: f.archivo, contentType: 'application/pdf', contentBytes: b64(bytes) }],
        },
        saveToSentItems: true,
      });
      if (prueba) await registrar({ ...base, estado: 'PRUEBA_ENVIADA', fecha_envio: new Date().toISOString(), detalle: 'a ' + destinoPrueba });
      else await actualizarFila(idx, { ...base, estado: 'ENVIADO', fecha_envio: new Date().toISOString() });
      t.ok++;
    } catch (e) {
      t.error++;
      try {
        if (idx !== null) await actualizarFila(idx, { ...base, estado: 'ERROR', detalle: e.message });
        else await registrar({ ...base, estado: 'ERROR', detalle: e.message });
      } catch (e2) { /* si el registro también falla, se informa abajo */ }
      aviso(`${f.numero_certificado}: no se envió (${e.message}).`, 'error');
    }
    $('#envio-barra').value = Math.round(((k + 1) / pend.length) * 100);
    const resta = Math.max(0, Math.round(((pend.length - k - 1) * (M.cfg.pausaMs + 1500)) / 60000));
    $('#envio-texto').textContent = `Enviando ${k + 1} de ${pend.length}… (enviados: ${t.ok}${t.error ? `, con problema: ${t.error}` : ''}). ${resta ? `Faltan unos ${resta} min. ` : ''}No cierres esta pestaña.`;
    if (k < pend.length - 1) await new Promise((ok) => setTimeout(ok, M.cfg.pausaMs));
  }
  M.enviando = false;
  await abrirRegistro();
  pintarLotes();
  $('#envio-texto').textContent = prueba
    ? `Terminado (prueba): se enviaron ${t.ok} correos a ${destinoPrueba}. Revisa tu bandeja: que el PDF se abra y que los datos estén bien.`
    : `Terminado: ${t.ok} enviados${t.error ? `, ${t.error} con problema (normalmente un correo mal escrito: búscalos abajo en "¿Un donante dice que no le llegó?", corrige el correo y reenvía)` : ''}${inciertos.length ? `. ${inciertos.length} quedaron en duda de un envío anterior: revisa Elementos enviados de ${M.cfg.buzon}` : ''}.`;
  aviso(prueba ? 'Prueba enviada a tu correo.' : `Envío terminado: ${t.ok} enviados${t.error ? `, ${t.error} con problema` : ''}.`, t.error ? 'warn' : 'ok');
}

/* ------------------------------------------------------------ reenviar a un donante */
async function reenviar(nombreLote, numero, correo) {
  if (M.enviando) return;
  const lote = M.lotes.find((l) => l.nombre === nombreLote);
  const f = lote && lote.filas.find((x) => x.numero_certificado === numero);
  if (!f) throw new Error('No se encontró ese certificado.');
  correo = limpiar(correo).toLowerCase();
  if (!correoValido(correo)) { aviso('Ese correo no parece válido. Revísalo.', 'warn'); return; }
  if (!confirm(`Se reenviará el certificado N.º ${numero} de ${f.nombre} a ${correo}. ¿Continuar?`)) return;
  M.enviando = true;
  try {
    await abrirRegistro();
    const it = lote.pdfs.get(f.archivo);
    if (!it) throw new Error('No se encontró el PDF en SharePoint');
    const bytes = await descargar(it);
    await graph('POST', `/users/${encodeURIComponent(M.cfg.buzon)}/sendMail`, {
      message: {
        subject: `Crea+ | Tu certificado de donación ${(f.periodo_desde || '2026').slice(0, 4)} N.º ${f.numero_certificado}`,
        body: { contentType: 'HTML', content: cuerpoCorreo(f) },
        toRecipients: [{ emailAddress: { address: correo } }],
        attachments: [{ '@odata.type': '#microsoft.graph.fileAttachment', name: f.archivo, contentType: 'application/pdf', contentBytes: b64(bytes) }],
      },
      saveToSentItems: true,
    });
    await registrar({ fecha_registro: new Date().toISOString(), lote: lote.nombre, numero_certificado: f.numero_certificado, documento: f.documento, nombre: f.nombre, correo, monto_total: f.monto_total, archivo: f.archivo, estado: 'REENVIADO', fecha_envio: new Date().toISOString(), detalle: correo !== f.correo ? 'correo corregido (antes ' + f.correo + ')' : '' });
    f.correo = correo;
    await abrirRegistro();
    aviso(`Certificado reenviado a ${correo}.`, 'ok');
  } finally { M.enviando = false; }
  pintarLotes(); pintarReenvio();
}

function pintarReenvio() {
  const caja = $('#resultados-reenvio'); if (!caja) return;
  const q = sinTildes($('#buscar-reenvio').value);
  if (q.length < 3) { caja.innerHTML = ''; return; }
  const res = [];
  for (const l of M.lotes) for (const f of l.filas) if (sinTildes(f.nombre + ' ' + f.documento).includes(q)) res.push({ l, f });
  if (!res.length) { caja.innerHTML = '<p class="muted">No hay certificados emitidos para esa búsqueda.</p>'; return; }
  const reg = M.registro ? M.registro.porNumero : new Map();
  const est = { ENVIADO: 'Enviado', REENVIADO: 'Reenviado', ERROR: 'No se pudo enviar', ENVIANDO: 'En duda', PRUEBA_ENVIADA: 'Sin enviar' };
  caja.innerHTML = `<table style="margin-top:10px"><thead><tr><th>Nombre</th><th>DNI</th><th>N.º</th><th>Estado</th><th>Correo</th><th></th></tr></thead><tbody>${res.slice(0, 20).map(({ l, f }, k) => {
    const r = reg.get(f.numero_certificado);
    return `<tr><td>${esc(f.nombre)}</td><td>${esc(f.documento)}</td><td>${esc(f.numero_certificado.slice(SERIE.length))}</td><td>${esc(r ? est[r.estado] || r.estado : 'Sin enviar')}</td>
      <td><input class="correo-reenvio" data-k="${k}" value="${esc(f.correo)}"></td><td><button class="sec btn-reenviar" data-k="${k}" data-l="${esc(l.nombre)}" data-n="${esc(f.numero_certificado)}">Reenviar</button></td></tr>`;
  }).join('')}</tbody></table>`;
}

/* ------------------------------------------------------------ interfaz */
function pintarSesion() {
  const dentro = !!M.token;
  $('#sesion').innerHTML = dentro
    ? `<span>👤 ${esc((M.yo && (M.yo.displayName || M.yo.mail)) || 'Sesión iniciada')}</span> <button id="btn-salir">Salir</button>`
    : '';
  if (dentro) $('#btn-salir').onclick = cerrarSesion;
  $('#btn-sp').hidden = !dentro;
  $('#btn-generar-sp').hidden = !dentro;
  $('#btn-generar').hidden = dentro;
  $('#paso4-sp').hidden = !dentro;
  $('#paso4-local').hidden = dentro;
  $('#buzon-texto').textContent = M.cfg.buzon;
  $('#ayuda-paso1').innerHTML = dentro
    ? 'La app lee sola la carpeta <b>' + esc(M.cfg.carpeta) + '</b> de SharePoint: reportes de pago, base de correos, certificados ya emitidos y plantillas.'
    : 'Modo sin conexión: arrastra los reportes, la base de donantes y las plantillas en <b>Agregar un archivo</b>.';
  if (typeof uxSesion === 'function') uxSesion(dentro);
}

function rangoLote(l) {
  const m = /^lote_(\d{4})(\d{2})(\d{2})\d{4}_(\d+)-(\d+)$/.exec(l.nombre);
  return m ? { desde: m[4], hasta: m[5], fecha: `${m[3]}/${m[2]}/${m[1]}` } : { desde: '', hasta: '', fecha: '' };
}

function pintarLotes() {
  const caja = $('#lista-lotes');
  if (!caja) return;
  if (!M.lotes.length) { caja.innerHTML = '<p class="caja info">Todavía no hay certificados creados. Créalos en el paso 3.</p>'; return; }
  const pruebas = M.registro ? M.registro.pruebaLotes : new Set();
  const orden = [...M.lotes].sort((a, b) => pendientesDe(b).pend.length - pendientesDe(a).pend.length || b.nombre.localeCompare(a.nombre));
  caja.innerHTML = orden.map((l) => {
    const p = pendientesDe(l); const r = rangoLote(l);
    const probado = pruebas.has(l.nombre) || M.probados.has(l.nombre);
    const titulo = `<h3>Certificados N.º ${r.desde} al ${r.hasta}</h3><p class="muted small" style="margin:0">${l.filas.length.toLocaleString('es-PE')} certificados · creados el ${r.fecha}</p>`;
    if (!p.pend.length) {
      return `<div class="lote">${titulo}<p class="caja ok" style="margin-bottom:0">✓ Todos enviados (${p.enviados.length.toLocaleString('es-PE')}).${p.inciertos.length ? ` ${p.inciertos.length} en duda: revisa Elementos enviados.` : ''}</p></div>`;
    }
    return `<div class="lote">${titulo}
      <p style="margin:10px 0 0"><b>${p.enviados.length.toLocaleString('es-PE')}</b> enviados · <b>${p.pend.length.toLocaleString('es-PE')}</b> por enviar${p.inciertos.length ? ` · <b>${p.inciertos.length}</b> en duda` : ''}</p>
      <div class="etapas">
        <div class="etapa ${probado ? 'hecha' : ''}"><h4>${probado ? '✓ ' : ''}a) Envíate una prueba</h4><p>Te llegan ${Math.min(PRUEBA_MAX, p.pend.length)} correos de ejemplo a ti. Revisa que el PDF se abra y que todo esté bien.</p>
          <button class="sec enviar-prueba" data-l="${esc(l.nombre)}">Enviarme ${Math.min(PRUEBA_MAX, p.pend.length)} de prueba</button></div>
        <div class="etapa"><h4>b) Envía a los donantes</h4><p>${probado ? 'Cuando hayas revisado la prueba en tu correo.' : 'Se activa después de la prueba.'}</p>
          <button class="enviar-lote" data-l="${esc(l.nombre)}" ${probado ? '' : 'disabled'}>Enviar a ${p.pend.length.toLocaleString('es-PE')} donantes</button></div>
      </div></div>`;
  }).join('');
}

function pintarCfg() {
  for (const k of Object.keys(CFG_DEFECTO)) { const el = $('#cfg-' + k); if (el) el.value = M.cfg[k]; }
}

async function iniciarM365() {
  M.probados = new Set();
  cargarCfg(); pintarCfg();
  $('#cfg-guardar').addEventListener('click', () => {
    for (const k of Object.keys(CFG_DEFECTO)) { const el = $('#cfg-' + k); if (el) M.cfg[k] = k === 'pausaMs' ? Math.max(2000, +el.value || 2500) : el.value.trim(); }
    guardarCfg(); M.siteId = null; M.driveBase = null; aviso('Configuración guardada.', 'ok');
  });
  $('#btn-sp').addEventListener('click', actualizarDesdeSharePoint);
  $('#btn-entrar-grande').addEventListener('click', iniciarSesion);
  $('#cfg-preparar').addEventListener('click', async () => {
    try { await prepararCarpeta(); } catch (e) { aviso('No se pudo preparar la carpeta: ' + e.message, 'error'); }
  });
  $('#btn-generar-sp').addEventListener('click', async () => {
    $('#btn-generar-sp').disabled = true;
    try { await generarEnSharePoint(); } catch (e) { aviso('No se pudieron guardar los certificados en SharePoint: ' + e.message, 'error'); }
    actualizarPaso3();
  });
  $('#btn-descargar-filas').addEventListener('click', () => {
    const L = M.ultimoLote; if (!L) return;
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([L.control['filas_para_REPORTE.txt']], { type: 'text/plain;charset=utf-8' }));
    a.download = `filas para el Excel - N ${L.n} al ${L.n + L.listos.length - 1}.txt`;
    document.body.appendChild(a); a.click(); a.remove();
  });
  $('#lista-lotes').addEventListener('click', async (e) => {
    const b = e.target.closest('.enviar-lote, .enviar-prueba'); if (!b) return;
    const prueba = b.classList.contains('enviar-prueba');
    b.disabled = true;
    try { await enviarLote(b.dataset.l, prueba); if (prueba) M.probados.add(b.dataset.l); } catch (err) { M.enviando = false; aviso('Error en el envío: ' + err.message, 'error'); }
    pintarLotes();
  });
  $('#buscar-reenvio').addEventListener('input', pintarReenvio);
  $('#resultados-reenvio').addEventListener('click', async (e) => {
    const b = e.target.closest('.btn-reenviar'); if (!b) return;
    const correo = $(`.correo-reenvio[data-k="${b.dataset.k}"]`).value;
    b.disabled = true;
    try { await reenviar(b.dataset.l, b.dataset.n, correo); } catch (err) { M.enviando = false; aviso('No se pudo reenviar: ' + err.message, 'error'); }
    b.disabled = false;
  });
  window.addEventListener('beforeunload', (e) => { if (M.enviando) { e.preventDefault(); e.returnValue = ''; } });
  try { await completarInicio(); } catch (e) { aviso(e.message, 'error'); }
  if (M.token) { try { M.yo = await graph('GET', '/me?$select=displayName,mail,userPrincipalName'); } catch (e) { M.token = null; sessionStorage.removeItem('cd_token'); aviso('La sesión venció. Vuelve a entrar.', 'warn'); } }
  pintarSesion();
}
window.addEventListener('DOMContentLoaded', iniciarM365);
