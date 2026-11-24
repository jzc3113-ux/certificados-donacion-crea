"""Construye index.html (un solo archivo) a partir de src/ y vendor/.
Uso: python3 build.py   (opcional: config.json con {"clientId": "...", "tenantId": "..."})"""
import base64
import json
import os

B = os.path.dirname(os.path.abspath(__file__))
html = open(os.path.join(B, 'src/index.html'), encoding='utf-8').read()
partes = {'/*__PDFLIB__*/': 'vendor/pdf-lib.min.js', '/*__JSZIP__*/': 'vendor/jszip.min.js',
          '/*__APP__*/': 'src/app.js', '/*__M365__*/': 'src/m365.js', '/*__UX__*/': 'src/ux.js'}
for marca, ruta in partes.items():
    code = open(os.path.join(B, ruta), encoding='utf-8').read().replace('</script', '<\\/script').replace('//# sourceMappingURL=', '// ')
    assert html.count(marca) == 1, marca
    a, b = html.split(marca)
    html = a + code + b
cfg_path = os.path.join(B, 'config.json')
cfg = json.load(open(cfg_path)) if os.path.exists(cfg_path) else {}
html = html.replace('/*__CFG__*/{}', json.dumps(cfg))
xlsx = base64.b64encode(open(os.path.join(B, 'src/registro_plantilla.xlsx'), 'rb').read()).decode()
html = html.replace("'/*__REGISTRO_B64__*/'", "'" + xlsx + "'")
emi = base64.b64encode(open(os.path.join(B, 'src/emitidos_plantilla.xlsx'), 'rb').read()).decode()
html = html.replace("'/*__EMITIDOS_B64__*/'", "'" + emi + "'")
for marca_img, ruta in {'__IMG_LOGO_BLANCO__': 'src/marca/logo_blanco.png', '__IMG_CREATIN__': 'src/marca/creatin_globo.png'}.items():
    html = html.replace(marca_img, 'data:image/png;base64,' + base64.b64encode(open(os.path.join(B, ruta), 'rb').read()).decode())
open(os.path.join(B, 'index.html'), 'w', encoding='utf-8').write(html)
print('index.html generado,', len(html), 'caracteres')
