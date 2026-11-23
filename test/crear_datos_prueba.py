"""Crea reportes FICTICIOS con el mismo formato que Niubiz, AFRUS y MC Procesos, y una base de donantes.
Ningún dato es real. Devuelve también los totales esperados para comparar con la app."""
import csv
import json
import os
from openpyxl import Workbook

OUT = os.path.join(os.path.dirname(__file__), "datos")
os.makedirs(OUT, exist_ok=True)

NIUBIZ_COLS = ['Id.Transaccion', 'Núm.Tarjeta', 'Cod.Autorización', 'Nom. Tarjeta Habiente', 'Núm. Iden. Beneficiario',
               'Nom. Beneficiario', 'Emisor', 'Producto', 'Fecha Afiliación', 'Canal Desafiliación', 'Fecha Desafiliación',
               'Canal Suspensión', 'Fecha Suspensión', 'Fecha Estimada Pago', 'Fecha Proceso Pago', 'Importe Cargo',
               'Importe Máximo', 'Estado Cargo', 'Cód. Acción', 'Motivo Rechazo', 'Nro. Lote', 'Nro. Reintento',
               'Núm. Token', 'Fecha de Vencimiento Token']
seq = [900000000000000]


def niubiz(dni, nombre, fecha, monto, ok, reintento=1):
    seq[0] += 1
    return {
        'Id.Transaccion': str(seq[0]) if ok or reintento % 2 == 0 else '', 'Núm.Tarjeta': '411111******1111',
        'Cod.Autorización': f'{seq[0] % 1000000:06d}' if ok else '', 'Nom. Tarjeta Habiente': nombre,
        'Núm. Iden. Beneficiario': "'" + dni, 'Nom. Beneficiario': nombre, 'Emisor': 'BANCO FICTICIO', 'Producto': 'DEBITO',
        'Fecha Afiliación': '01/01/2024', 'Fecha Estimada Pago': fecha, 'Fecha Proceso Pago': fecha,
        'Importe Cargo': f'{monto:.2f}', 'Importe Máximo': '500.00', 'Estado Cargo': 'AUTORIZADO' if ok else 'DENEGADO',
        'Cód. Acción': 'AB123456789', 'Motivo Rechazo': '' if ok else 'SALDO INSUFICIENTE', 'Nro. Lote': '123456789',
        'Nro. Reintento': str(reintento),
    }


filas = []
# D1: 50 mensual ene-oct, marzo rechazado (con 2 reintentos fallidos)
for m in range(1, 11):
    ok = m != 3
    filas.append(niubiz('11111111', 'ANA PRUEBA UNO', f'05/{m:02d}/2026', 50, ok))
    if not ok:
        filas.append(niubiz('11111111', 'ANA PRUEBA UNO', f'08/{m:02d}/2026', 50, False, 2))
# D2: DNI de 7 dígitos; 30 ene-jun y 40 jul-oct
for m in range(1, 11):
    filas.append(niubiz('2222222', 'BETO PRUEBA DOS', f'10/{m:02d}/2026', 30 if m <= 6 else 40, True))
# D3: ene-feb (sin correo en la base)
for m in (1, 2):
    filas.append(niubiz('33333333', 'CARLA PRUEBA TRES', f'12/{m:02d}/2026', 50, True))
# D4: dic-2025 (fuera de período) y ene-2026
filas.append(niubiz('44444444', 'DIEGO PRUEBA CUATRO', '15/12/2025', 50, True))
filas.append(niubiz('44444444', 'DIEGO PRUEBA CUATRO', '15/01/2026', 50, True))
# muchos rechazos de relleno
for k in range(300):
    filas.append(niubiz(f'9{k:07d}', 'RECHAZADO FICTICIO', '20/04/2026', 20, False))

# Dos archivos con solapamiento (las filas de enero-febrero aparecen en ambos)
a = [f for f in filas if f['Fecha Proceso Pago'][3:5] in ('01', '02', '12')] + [f for f in filas if f['Fecha Proceso Pago'][3:5] in ('03', '04', '05')]
b = [f for f in filas if f['Fecha Proceso Pago'][3:5] in ('01', '02')] + [f for f in filas if f['Fecha Proceso Pago'][3:5] not in ('01', '02', '12', '03', '04', '05')]
for nombre, data in (('_repositorio_descarga_async_111_reporteCargos.csv', a), ('_repositorio_descarga_async_222_reporteCargos.csv', b)):
    with open(os.path.join(OUT, nombre), 'w', newline='', encoding='utf-8') as fh:
        w = csv.DictWriter(fh, fieldnames=NIUBIZ_COLS, quoting=csv.QUOTE_MINIMAL)
        w.writeheader()
        for r in data:
            w.writerow({k: r.get(k, '') for k in NIUBIZ_COLS})

# AFRUS (xlsx, 56 columnas reales)
AFRUS_COLS = ['id', 'type', 'amount', 'conversion_amount', 'year', 'month', 'day', 'created_at_utc_0_servidor', 'created_at_utc_x_usuario',
              'updated_at_utc_0_servidor', 'updated_at_utc_x_usuario', 'status', 'payment_method', 'card', 'brand', 'transaction_gateway_code',
              'subscription_id', 'is_first_charge', 'appeal', 'fund', 'url_landing', 'utm_source', 'utm_campaign', 'utm_content', 'utm_term',
              'utm_medium', 'lead_id', 'first_name', 'last_name', 'email', 'phone', 'identification', 'zip_code', 'country', 'province', 'city',
              'address', 'full_address', 'organization', 'campaign_id', 'campaign', 'form_id', 'widget', 'form_id_legacy', 'form_legacy',
              'gateway', 'currency', 'organization_currency', 'ticket_ids', 'cyc_applied', 'cyc_amount_collected', 'cyc_total_rate',
              'donation_type', 'incentive', 'coupon', 'third_party_lead_id']
af = []
aid = [5000000]


def afrus(dni, nom, ape, email, fecha, monto, status):
    aid[0] += 1
    y, m, d = fecha.split('-')
    r = dict.fromkeys(AFRUS_COLS, '')
    r.update({'id': aid[0], 'type': 'Charge', 'amount': monto, 'conversion_amount': monto, 'year': int(y), 'month': int(m), 'day': int(d),
              'created_at_utc_0_servidor': fecha + ' 15:00:00', 'created_at_utc_x_usuario': fecha + ' 10:00:00', 'status': status,
              'payment_method': 'CARD', 'first_name': nom, 'last_name': ape, 'email': email, 'identification': dni,
              'campaign': 'Nieto Crea+', 'currency': 'Peruvian Nuevo Sol', 'donation_type': 'Estándar'})
    return r


for m in (3, 4, 5):
    af.append(afrus('55555555', 'Elena', 'Prueba Cinco', 'elena5@ficticio.pe', f'2026-{m:02d}-03', 20, 'SUCCEEDED'))
af.append(afrus('55555555', 'Elena', 'Prueba Cinco', 'elena5@ficticio.pe', '2026-06-03', 20, 'FAILED'))
af.append(afrus('55555555', 'Elena', 'Prueba Cinco', 'elena5@ficticio.pe', '2026-07-03', 20, 'REFUNDED'))
af.append(afrus('11111111', 'Ana', 'Prueba Uno', 'ana1@ficticio.pe', '2026-11-02', 100, 'SUCCEEDED'))
af.append(afrus('', 'Fabio', 'Prueba Seis', 'fabio6@ficticio.pe', '2026-08-09', 25, 'SUCCEEDED'))
af.append(afrus('', 'Sin', 'Documento', 'nadie@ficticio.pe', '2026-08-10', 15, 'SUCCEEDED'))
wb = Workbook(); ws = wb.active; ws.title = 'Transactions'; ws.append(AFRUS_COLS)
for r in af:
    ws.append([r[c] for c in AFRUS_COLS])
wb.save(os.path.join(OUT, '20-10-2026-Transactions-FICTICIO.xlsx'))

# MC Procesos (xlsx con 2 hojas, valores con espacio no separable al final como el real)
MC_COLS = ['Tarjeta', 'Tarjetahabiente', 'Producto', 'Código Servicio', 'Usuario', 'Fecha Afiliación', 'Importe Máximo', 'Importe Cargo',
           'Número Recibo', 'Fecha Act Tarjeta', 'Fecha Autorización', 'Fecha Depósito', 'Código Respuesta', 'Código Autorización',
           'Mensaje Respuesta', 'Acción', 'Estado Transacción']
NB = ' '
wb = Workbook(); s1 = wb.active; s1.title = '5008271_D0634'
s1.append(['PROCESOS DE MEDIOS DE PAGO S.A.']); s1.append(['Reporte de cargos', '2026'])
s2 = wb.create_sheet('MCproceso marzo a octubre'); s2.append(MC_COLS)
rec = [10000000000000000000]


def mc(cs, nombre, fecha, monto, aprobado):
    rec[0] += 1
    return ['545454******5454', nombre, 'Crea+', cs, nombre, '01/01/2025' + NB, 200, monto, str(rec[0]), '01/01/2025' + NB,
            fecha + NB if not aprobado else fecha + NB, fecha + NB if aprobado else '', '00' if aprobado else '96', '123456' if aprobado else '',
            'Aprobado' if aprobado else 'EXC.RECHZ.VISA.A', 'Aprobado' if aprobado else 'Denegado', ('Depositado' if aprobado else 'Autorizado') + NB]


for m in (3, 4):
    s2.append(mc('77777777', 'PRUEBA SIETE, GUSTAVO', f'07/{m:02d}/2026', 35, True))
for m in (5, 6, 7):
    s2.append(mc('77777777', 'PRUEBA SIETE, GUSTAVO', f'07/{m:02d}/2026', 35, False))
for m in (3, 4, 5):
    s2.append(mc('78787878', 'PRUEBA OCHO, HILDA', f'07/{m:02d}/2026', 35, False))
wb.save(os.path.join(OUT, 'marzo a octubre FICTICIO.xlsx'))

# Base de donantes
wb = Workbook(); ws = wb.active; ws.title = 'Base'
ws.append(['DNI', 'Nombre completo', 'Correo', 'Celular'])
for r in [('11111111', 'Ana Prueba Uno', 'ana1@ficticio.pe'), ('02222222', 'José Ñandú Prueba Dos', 'beto2@ficticio.pe'),
          ('33333333', 'Carla Prueba Tres', ''), ('44444444', 'Diego Prueba Cuatro', 'diego4@ficticio.pe'),
          ('55555555', 'Elena Prueba Cinco', 'elena5@ficticio.pe'), ('66666666', 'Fabio Prueba Seis', 'fabio6@ficticio.pe'),
          ('77777777', 'Gustavo Prueba Siete', 'gustavo7@ficticio.pe'), ('77777777', 'Gustavo Prueba Siete', 'gus.siete@ficticio.pe')]:
    ws.append(list(r) + ['999999999'])
wb.save(os.path.join(OUT, 'Base_donantes_FICTICIA.xlsx'))

esperado = {
    '11111111': {'total': 550.0, 'estado': 'LISTO', 'lineas': [[9, 50.0], [1, 100.0]]},
    '02222222': {'total': 340.0, 'estado': 'LISTO', 'lineas': [[6, 30.0], [4, 40.0]]},
    '33333333': {'total': 100.0, 'estado': 'REVISAR'},
    '44444444': {'total': 50.0, 'estado': 'LISTO'},
    '55555555': {'total': 60.0, 'estado': 'LISTO'},
    '66666666': {'total': 25.0, 'estado': 'LISTO'},
    '77777777': {'total': 70.0, 'estado': 'REVISAR'},
    'sin_doc': 1,
}
json.dump(esperado, open(os.path.join(OUT, 'esperado.json'), 'w'), indent=1)
print('ok', os.listdir(OUT))
