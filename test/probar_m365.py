"""Prueba del modo conectado con un SharePoint/Outlook SIMULADO (no toca Microsoft 365 real)."""
import json
import os
import re
import sys
import urllib.parse
from playwright.sync_api import sync_playwright

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
D = os.path.join(BASE, 'test', 'datos')
RAIZ = 'Certificados 2026'
PLANT = os.environ.get('PLANTILLAS', os.path.join(BASE, 'test', 'plantillas'))
fs = {}  # ruta -> bytes
for f in os.listdir(D):
    if f.endswith('.json'):
        continue
    sub = '02_BASE_DONANTES' if f.startswith('Base') else '01_REPORTES/' + ('Niubiz' if 'reporteCargos' in f else 'AFRUS' if 'Transactions' in f else 'MC_Izipay')
    fs[f'{RAIZ}/{sub}/{f}'] = open(os.path.join(D, f), 'rb').read()
fs[f'{RAIZ}/05_HERRAMIENTA/plantilla_personas.pdf'] = open(os.path.join(PLANT, 'plantilla_personas.pdf'), 'rb').read()
fs[f'{RAIZ}/05_HERRAMIENTA/LEEME.txt'] = b'x'
filas = []      # registro Excel
correos = []    # sendMail
fallar = set()  # direcciones que fallan una vez
carpetas = set()
fallos = []


def check(c, m):
    print(('OK   ' if c else 'FALLA') + ' ' + m)
    if not c:
        fallos.append(m)


def hijos(ruta):
    pref = ruta.rstrip('/') + '/'
    vistos, out = set(), []
    for k in fs:
        if k.startswith(pref):
            resto = k[len(pref):]
            nombre = resto.split('/')[0]
            if nombre in vistos:
                continue
            vistos.add(nombre)
            if '/' in resto:
                out.append({'id': 'F:' + pref + nombre, 'name': nombre, 'folder': {}})
            else:
                out.append({'id': 'I:' + k, 'name': nombre, 'file': {}, 'size': len(fs[k]),
                            '@microsoft.graph.downloadUrl': 'https://fake.download/' + urllib.parse.quote(k)})
    return out


def manejar(route):
    req = route.request
    u = urllib.parse.urlparse(req.url)
    if u.netloc == 'fake.download':
        return route.fulfill(status=200, body=fs[urllib.parse.unquote(u.path[1:])], headers={'Access-Control-Allow-Origin': '*'})
    ruta = urllib.parse.unquote(u.path)
    j = lambda st, obj=None: route.fulfill(status=st, content_type='application/json', body=json.dumps(obj) if obj is not None else '',
                                           headers={'Access-Control-Allow-Origin': '*'})
    if ruta.endswith('/me'):
        return j(200, {'displayName': 'Operadora de Prueba', 'mail': 'operadora@creamas.org'})
    if ruta == '/v1.0/sites/SITE/drives':
        return j(200, {'value': [{'id': 'OTRA', 'name': 'Documentos', 'webUrl': 'https://x/sites/X/Shared%20Documents'}, {'id': 'DRV', 'name': 'inbox', 'webUrl': 'https://x/sites/X/inbox'}]})
    if ruta.startswith('/v1.0/sites/creamas.sharepoint.com:'):
        return j(200, {'id': 'SITE'})
    m = re.match(r'/v1\.0/drives/DRV/root(?::/(.+?):)?/children$', ruta)
    if m and req.method == 'POST':
        cuerpo = json.loads(req.post_data)
        p = ((m.group(1) + '/') if m.group(1) else '') + cuerpo['name']
        if p in carpetas or any(k.startswith(p + '/') for k in fs):
            return j(409, {'error': {'code': 'nameAlreadyExists', 'message': 'ya existe'}})
        carpetas.add(p)
        return j(201, {'id': 'F:' + p, 'name': cuerpo['name'], 'folder': {}})
    if m:
        h = hijos(m.group(1))
        return j(200, {'value': h}) if h else j(404, {'error': {'code': 'itemNotFound', 'message': 'no existe'}})
    m = re.match(r'/v1\.0/drives/DRV/root:/(.+?):/content$', ruta)
    if m and req.method == 'PUT':
        p = m.group(1)
        if p in fs and 'conflictBehavior=fail' in u.query:
            return j(409, {'error': {'code': 'nameAlreadyExists', 'message': 'ya existe'}})
        fs[p] = req.post_data_buffer
        return j(201, {'id': 'I:' + p, 'name': p.split('/')[-1]})
    m = re.match(r'/v1\.0/drives/DRV/root:/(.+)$', ruta)
    if m and req.method == 'GET':
        p = m.group(1)
        if p in carpetas or any(k.startswith(p + '/') for k in fs):
            return j(200, {'id': 'F:' + p, 'name': p.split('/')[-1], 'folder': {}, 'webUrl': 'https://creamas.sharepoint.com/sites/X/' + p})
        return j(200, {'id': 'I:' + p, 'name': p.split('/')[-1]}) if p in fs else j(404, {'error': {'code': 'itemNotFound', 'message': 'no existe'}})
    if '/workbook/tables/Envios/rows' in ruta:
        if req.method == 'GET':
            return j(200, {'value': [{'index': i, 'values': [v]} for i, v in enumerate(filas)]})
        if req.method == 'POST':
            filas.append(json.loads(req.post_data)['values'][0])
            return j(201, {'index': len(filas) - 1})
        mi = re.search(r'itemAt\(index=(\d+)\)', ruta)
        if req.method == 'PATCH' and mi:
            filas[int(mi.group(1))] = json.loads(req.post_data)['values'][0]
            return j(200, {})
    if ruta.endswith('/sendMail'):
        msg = json.loads(req.post_data)['message']
        para = msg['toRecipients'][0]['emailAddress']['address']
        if para in fallar:
            fallar.discard(para)
            return j(503 if False else 500, {'error': {'code': 'ErrorX', 'message': 'fallo simulado'}})
        correos.append({'para': para, 'asunto': msg['subject'], 'adjunto': msg['attachments'][0]['name'], 'bytes': len(msg['attachments'][0]['contentBytes'])})
        return j(202)
    return j(404, {'error': {'code': 'x', 'message': 'ruta no simulada ' + ruta}})


with sync_playwright() as p:
    nav = p.chromium.launch()
    ctx = nav.new_context()
    ctx.add_init_script("""sessionStorage.setItem('cd_token', JSON.stringify({acceso:'FAKE', renovar:'R', vence: Date.now()+3600e3}));
      localStorage.setItem('cd_m365', JSON.stringify({clientId:'cli', tenantId:'ten', pausaMs: 2000}));""")
    ctx.route('https://graph.microsoft.com/**', manejar)
    ctx.route('https://fake.download/**', manejar)
    pg = ctx.new_page()
    errores = []
    pg.on('pageerror', lambda e: errores.append(str(e)))
    pg.on('dialog', lambda d: d.accept())
    pg.goto('file://' + os.path.join(BASE, 'index.html'))
    pg.wait_for_selector('#btn-salir')
    check('Operadora de Prueba' in pg.inner_text('#sesion'), 'sesión mostrada')
    # Preparar carpeta: crea la estructura, sube la plantilla cargada y es repetible sin duplicar
    pg.evaluate("S.plantillas.empresas = new Uint8Array([37,80,68,70])")
    pg.evaluate("prepararCarpeta()"); pg.wait_for_function("document.querySelector('#enlace-carpeta').textContent.includes('Carpeta')", timeout=30000)
    existe = lambda r: r in carpetas or any(k.startswith(r + '/') for k in fs)
    check(all(existe(f'{RAIZ}/' + r) for r in ['01_REPORTES/Niubiz', '01_REPORTES/Manuales', '01_REPORTES/Exclusiones', '03_CONTROL', '04_LOTES', '06_MANUAL']), f'estructura creada ({len(carpetas)} carpetas nuevas)')
    check(f'{RAIZ}/05_HERRAMIENTA/plantilla_empresas.pdf' in fs, 'plantilla subida a 05_HERRAMIENTA')
    n = len(carpetas); pg.evaluate("prepararCarpeta()"); pg.wait_for_timeout(1500)
    check(len(carpetas) == n, 'repetir no crea duplicados')
    del fs[f'{RAIZ}/05_HERRAMIENTA/plantilla_empresas.pdf']; pg.evaluate("S.plantillas.empresas = null")
    pg.click('#btn-sp')
    pg.wait_for_function("document.querySelectorAll('#tabla-archivos tbody tr').length >= 5", timeout=60000)
    check(pg.evaluate('!!S.plantillas.personas'), 'plantilla leída desde SharePoint')
    pg.click('#btn-revisar')
    i = pg.evaluate("S.donantes.findIndex(d => d.doc === '77777777')")
    pg.select_option('#filtro', 'TODOS')
    pg.select_option(f'select.ed-correo-sel[data-i="{i}"]', 'gustavo7@ficticio.pe')
    for k in range(3):
        pg.evaluate(f'verMuestra({k})'); pg.wait_for_function("!document.querySelector('#modal').hidden"); pg.click('#cerrar-modal')
    pg.fill('#p-numero', '291'); pg.dispatch_event('#p-numero', 'change'); pg.check('#confirmo')
    pg.click('#btn-generar-sp')
    pg.wait_for_function('M.lotes.length === 1', timeout=120000)
    lote = [k for k in fs if '/04_LOTES/' in k]
    check(sum(1 for k in lote if k.endswith('.pdf')) == 6 and any(k.endswith('manifiesto_envio.csv') for k in lote), f'lote guardado en SharePoint ({len(lote)} archivos)')
    check(f'{RAIZ}/03_CONTROL/Registro certificados 2026.xlsx' in fs, 'registro Excel creado en 03_CONTROL')
    # 1) modo prueba
    pg.evaluate('M.cfg.pausaMs = 5')
    pg.click('.enviar-lote')
    pg.wait_for_function("document.querySelector('#envio-texto').textContent.startsWith('Terminado')", timeout=60000)
    check(len(correos) == 6 and all(c['para'] == 'operadora@creamas.org' and c['asunto'].startswith('[PRUEBA]') for c in correos), 'modo prueba: 6 correos solo a la operadora')
    check(pg.evaluate("pendientesDe(M.lotes[0]).pend.length") == 6, 'tras la prueba, siguen 6 pendientes reales')
    # 2) envío real con un fallo simulado
    correos.clear(); fallar.add('diego4@ficticio.pe')
    pg.uncheck('#modo-prueba')
    pg.evaluate("document.querySelector('#envio-texto').textContent=''")
    pg.click('.enviar-lote')
    pg.wait_for_function("document.querySelector('#envio-texto').textContent.startsWith('Terminado')", timeout=60000)
    check(len(correos) == 5, f'envío real: 5 enviados, 1 falló ({len(correos)})')
    check(all(c['adjunto'].endswith('.pdf') and c['bytes'] > 100000 for c in correos), 'cada correo lleva su PDF adjunto')
    check(pg.evaluate("pendientesDe(M.lotes[0]).pend.length") == 1, 'queda 1 pendiente (el que falló)')
    # 3) reanudar: solo se envía el pendiente
    pg.evaluate("document.querySelector('#envio-texto').textContent=''")
    pg.click('.enviar-lote')
    pg.wait_for_function("document.querySelector('#envio-texto').textContent.startsWith('Terminado')", timeout=60000)
    para = [c['para'] for c in correos]
    check(len(correos) == 6 and len(set(para)) == 6 and 'diego4@ficticio.pe' in para, 'reanudar envía solo el que faltaba; 6 donantes, sin duplicados')
    estados = [f[8] for f in filas if f[2]]
    check(estados.count('ENVIADO') == 6 and estados.count('PRUEBA_ENVIADA') == 6, f'registro: 6 ENVIADO y 6 PRUEBA ({estados.count("ENVIADO")}, {estados.count("PRUEBA_ENVIADA")})')
    check(pg.evaluate("pendientesDe(M.lotes[0]).pend.length") == 0 and 'completo' in pg.inner_text('#lista-lotes'), 'lote completo')
    # 4) nueva sesión: los mismos donantes quedan bloqueados y el número sigue
    pg.reload(); pg.wait_for_selector('#btn-salir')
    pg.click('#btn-sp')
    pg.wait_for_function("document.querySelectorAll('#tabla-archivos tbody tr').length >= 5", timeout=60000)
    pg.click('#btn-revisar')
    check(pg.evaluate("S.donantes.filter(d => d.incluir && d.estado === 'LISTO').length") == 0, 'nueva sesión: nadie queda para reemitir')
    check(pg.input_value('#p-numero') == '297', f"nueva sesión: siguiente número 297 ({pg.input_value('#p-numero')})")
    pg.screenshot(path=os.path.join(BASE, 'test', 'salida', 'pantalla_m365.png'), full_page=True)
    check(not errores, f'sin errores JS {errores}')
    nav.close()
print('\nRESULTADO:', 'TODO OK' if not fallos else f'{len(fallos)} FALLAS')
sys.exit(1 if fallos else 0)
