"""Prueba de punta a punta de la app con datos ficticios, en Chrome sin interfaz."""
import csv
import io
import json
import os
import sys
import zipfile
from playwright.sync_api import sync_playwright

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATOS = os.path.join(BASE, 'test', 'datos')
SALIDA = os.path.join(BASE, 'test', 'salida')
os.makedirs(SALIDA, exist_ok=True)
esperado = json.load(open(os.path.join(DATOS, 'esperado.json')))
PLANT = os.environ.get('PLANTILLAS', os.path.join(BASE, 'test', 'plantillas'))  # carpeta local con las plantillas PDF (no se suben)
archivos = [os.path.join(DATOS, f) for f in sorted(os.listdir(DATOS)) if not f.endswith('.json')]
archivos += [os.path.join(PLANT, 'plantilla_personas.pdf'), os.path.join(PLANT, 'plantilla_empresas.pdf')]
fallos = []


def check(cond, msg):
    print(('OK   ' if cond else 'FALLA') + ' ' + msg)
    if not cond:
        fallos.append(msg)


with sync_playwright() as p:
    nav = p.chromium.launch()
    pag = nav.new_page(accept_downloads=True)
    errores = []
    pag.on('pageerror', lambda e: errores.append(str(e)))
    pag.goto('file://' + os.path.join(BASE, 'index.html'))
    pag.set_input_files('#input-archivos', archivos)
    pag.wait_for_function("document.querySelectorAll('#tabla-archivos tbody tr').length >= 5", timeout=60000)
    tipos = pag.eval_on_selector_all('#tabla-archivos tbody tr', 'rs => rs.map(r => [r.cells[0].innerText, r.cells[1].innerText, r.cells[3].innerText])')
    for t in tipos:
        print('   archivo:', t)
    check(all(t[1] != 'No reconocido' for t in tipos), 'todos los archivos reconocidos')
    print('  ', pag.inner_text('#resumen-archivos'))
    pag.click('#btn-revisar')
    pag.select_option('#filtro', 'TODOS')
    don = pag.evaluate("S.donantes.map(d => ({doc:d.doc, total:d.total, estado:d.estado, lineas:d.lineas.map(l=>[l.cantidad,l.valor]), obs:d.obs, nombre:d.nombre, correo:d.correo}))")
    pordoc = {d['doc']: d for d in don}
    for doc, e in esperado.items():
        if doc == 'sin_doc':
            continue
        d = pordoc.get(doc)
        check(d is not None, f'{doc} existe')
        if not d:
            continue
        check(abs(d['total'] - e['total']) < 0.001, f"{doc} total {d['total']} = {e['total']}")
        check(d['estado'] == e['estado'], f"{doc} estado {d['estado']} = {e['estado']} {d['obs']}")
        if 'lineas' in e:
            check(d['lineas'] == e['lineas'], f"{doc} líneas {d['lineas']}")
    check(set(pordoc) == {k for k in esperado if k != 'sin_doc'}, f'sin donantes de más: {sorted(pordoc)}')
    check(pag.evaluate('S.sinDoc.length') == esperado['sin_doc'], 'pagos sin DNI = 1')
    check(pag.evaluate('S.duplicados') > 0, f"duplicados descartados: {pag.evaluate('S.duplicados')}")
    # Corregir el donante con dos correos: elegir uno
    i = pag.evaluate("S.donantes.findIndex(d => d.doc === '77777777')")
    pag.select_option(f'select.ed-correo-sel[data-i="{i}"]', 'gustavo7@ficticio.pe')
    check(pag.evaluate("S.donantes.find(d => d.doc === '77777777').estado") == 'LISTO', '77777777 pasa a LISTO tras elegir correo')
    # Ver 3 muestras
    for doc in ('11111111', '02222222', '77777777'):
        i = pag.evaluate(f"S.donantes.findIndex(d => d.doc === '{doc}')")
        pag.evaluate(f"verMuestra({i})")
        pag.wait_for_function("!document.querySelector('#modal').hidden")
        pag.click('#cerrar-modal')
    check(pag.is_disabled('#btn-generar'), 'botón generar bloqueado sin número ni confirmación')
    pag.fill('#p-numero', '291'); pag.dispatch_event('#p-numero', 'change')
    pag.check('#confirmo')
    check(not pag.is_disabled('#btn-generar'), 'botón generar habilitado')
    print('  ', pag.inner_text('#rango-numeros'))
    with pag.expect_download(timeout=120000) as dl:
        pag.click('#btn-generar')
    ruta = os.path.join(SALIDA, dl.value.suggested_filename)
    dl.value.save_as(ruta)
    pag.wait_for_timeout(500)
    check(pag.is_disabled('#btn-generar'), 'tras generar no se puede volver a generar el mismo lote')
    check(pag.evaluate("S.donantes.filter(d => d.incluir && d.estado === 'LISTO').length") == 0, 'donantes generados quedan como ya emitidos')
    pag.screenshot(path=os.path.join(SALIDA, 'pantalla.png'), full_page=True)
    check(not errores, f'sin errores de JavaScript {errores}')
    nav.close()

z = zipfile.ZipFile(ruta)
nombres = z.namelist()
pdfs = [n for n in nombres if n.endswith('.pdf')]
check(len(pdfs) == 6, f'{len(pdfs)} PDF generados (esperado 6)')
man = list(csv.DictReader(io.StringIO(z.read('manifiesto_envio.csv').decode('utf-8-sig')), delimiter=';'))
for m in man:
    print('   manifiesto:', m['numero_certificado'], m['documento'], m['correo'], m['monto_total'], m['archivo'])
check([m['numero_certificado'] for m in man] == [f'002-000021,{n}' for n in range(291, 297)], 'numeración correlativa 291-296')
check(all(abs(float(m['monto_total']) - esperado[m['documento']]['total']) < 0.001 for m in man), 'montos del manifiesto = esperados')
sus = list(csv.DictReader(io.StringIO(z.read('sustento_pagos.csv').decode('utf-8-sig')), delimiter=';'))
check(len(sus) == len({s['id_pago'] for s in sus}), 'ningún pago en dos certificados')
rep = z.read('filas_para_REPORTE.txt').decode('utf-8-sig').split('\r\n')
check(all(len(r.split('\t')) == 34 for r in rep), 'filas para REPORTE con 34 columnas (B a AI)')
for n in pdfs:
    open(os.path.join(SALIDA, os.path.basename(n)), 'wb').write(z.read(n))
print('\nRESULTADO:', 'TODO OK' if not fallos else f'{len(fallos)} FALLAS')
sys.exit(1 if fallos else 0)
