# Certificados de donación · Crea+

Herramienta interna para emitir y enviar los certificados de donación de Crea+ a partir de los reportes de
Niubiz, AFRUS y MC Procesos / Izipay.

**Abrir la herramienta:** https://jzc3113-ux.github.io/certificados-donacion-crea/

## Qué hace

1. Lee los reportes de las plataformas, la base de donantes y el registro (desde la carpeta de SharePoint o arrastrando archivos).
2. Cuenta solo cobros exitosos del período y descarta rechazos, devoluciones y duplicados.
3. Muestra la lista por donante para revisar y corregir.
4. Genera los PDF sobre la plantilla oficial, con numeración correlativa, y los guarda en SharePoint.
5. Envía cada certificado desde el buzón institucional, con modo prueba, registro en Excel y reanudación sin duplicados.

## Seguridad y privacidad

- Este repositorio contiene **solo código**. No contiene datos de donantes, reportes, registros ni las plantillas con firma.
- `.gitignore` bloquea archivos .csv, .xlsx, .pdf y .zip. No lo quites.
- Los datos se procesan en el navegador de quien usa la herramienta y solo se guardan en el SharePoint de Crea+.
- Solo cuentas de Crea+ pueden iniciar sesión (registro de aplicación de un solo inquilino en Microsoft Entra ID).
- El Id. de aplicación y el Id. de directorio no son secretos. No hay claves ni contraseñas en el código.

## Estructura

- `index.html`: la herramienta completa (se genera con `python3 build.py`).
- `src/`: código fuente (`app.js` lógica y PDF, `m365.js` inicio de sesión, SharePoint y correo, `index.html` interfaz).
- `vendor/`: pdf-lib (MIT) y JSZip (MIT).
- `test/`: pruebas automáticas con datos ficticios (requieren las plantillas PDF, que no están en el repositorio).

## Configuración (TI)

Crear `config.json` (no se sube si contiene algo sensible; los Id. no lo son) con:

```json
{ "clientId": "Id. de aplicación", "tenantId": "Id. de directorio" }
```

y ejecutar `python3 build.py`. También se pueden ingresar en la sección **Configuración (solo TI)** de la herramienta.

Registro de aplicación en Entra ID: tipo SPA, URI de redirección `https://jzc3113-ux.github.io/certificados-donacion-crea/`,
permisos delegados de Microsoft Graph `User.Read`, `Sites.ReadWrite.All` y `Mail.Send.Shared`, con consentimiento de administrador.
La cuenta que envía necesita el permiso "Enviar como" sobre el buzón institucional.
