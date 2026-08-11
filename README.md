# Validación simultánea de contenedores

Aplicación móvil para validar cajas y artículos por contenedor, equipo y tarima. El código puede mantenerse en GitHub; la URL operativa se publica como aplicación web de Google Apps Script para tener acceso directo y seguro a Google Sheets y Drive.

## Funciones incluidas

- Varios contenedores y equipos activos simultáneamente.
- Tarimas únicas asignadas a un solo equipo.
- Varias personas del mismo equipo ven la misma tarima actualizada cada 4 segundos.
- Una captura por caja, con consolidación automática por artículo.
- Referencia interna o EAN: basta con uno, pero acepta ambos.
- Fotos obligatorias del artículo y la caja máster.
- Edición mediante lápiz mientras la tarima está en proceso.
- Bloqueo automático al finalizar la tarima.
- Administración posterior directamente desde Google Sheets.
- Carpetas de Drive independientes por contenedor.

## Instalación

1. Crea un Google Sheet vacío.
2. Abre **Extensiones → Apps Script**.
3. Copia `Code.gs`, `Index.html` y `appsscript.json` en el proyecto.
4. Ejecuta manualmente `setupSystem` una sola vez y autoriza Sheets/Drive.
5. En Apps Script elige **Implementar → Nueva implementación → Aplicación web**.
6. Ejecutar como: **tú**. Acceso: **cualquier persona con el enlace** (o solo tu organización, si corresponde).
7. Comparte la URL `/exec` con los teléfonos.

## Actualizaciones desde GitHub

Instala `clasp`, inicia sesión y vincula el proyecto con `clasp clone ID_DEL_SCRIPT`. Después podrás usar `clasp push` para publicar cambios desde el repositorio. Nunca subas credenciales ni archivos de autenticación al repositorio.

## Regla administrativa

Una tarima finalizada ya no muestra el lápiz y el backend rechaza cambios desde teléfonos. El responsable puede corregir la hoja `CAPTURAS` directamente y después ejecutar `adminRefreshReport` desde Apps Script para regenerar el reporte y las vistas por contenedor.

## Hojas generadas

- `CONFIGURACION`
- `CONTENEDORES`
- `EQUIPOS`
- `TARIMAS`
- `CAPTURAS`
- `REPORTE_GENERAL`
- `VISTA_<CONTENEDOR>` (una vista automática por cada contenedor)

No cambies los encabezados técnicos. Puedes ocultar las hojas de detalle para los usuarios y trabajar principalmente en `REPORTE_GENERAL`.
