# Correo de salud de ATMs – Auditoría Continua

Plantilla HTML compatible con Outlook de escritorio (tablas, CSS inline, 640 px, sin JS/imágenes base64)
para el nodo **Send Email** de KNIME. Solo lleva información agregada.

## Archivos

| Archivo | Para qué |
|---|---|
| `alerta_atm_template.html` | Plantilla que se pega en Send Email (diseño v2) |
| `alerta_atm_template_v1.html` | Respaldo del diseño anterior |
| `knime_expresiones.md` | Expresiones Column Expressions / Rule Engine / GroupBy |
| `preview.py` + `sample_vars.json` | Vista previa local con 3 escenarios (OK, atención, crítico) |
| `check_outlook.py` | Validación de compatibilidad con Outlook |

## Variables (todas String: `$${Snombre}$$`)

| Variable | Ejemplo | De dónde sale |
|---|---|---|
| `fecha_corte` | `20/10/2026` | Fecha del corte procesado (String) |
| `fecha_generacion` | `21/10/2026 06:05` | Fecha/hora de ejecución (Create Date&Time / Date&Time to String) |
| `estado_texto` | `ATENCIÓN` | Rule Engine 1 (sección 3) |
| `estado_color` | `#D97706` | Rule Engine 2 (sección 3) |
| `resumen` | `Conciliación del 98.4% · 3 casos…` | Column Expressions B (sección 4), HTML con entidades |
| `total_atm` | `412` | Conteo de ATMs del universo |
| `match_trx` / `sin_match_trx` | `409` / `3` | ATMs con/sin match Journal vs Switch |
| `pct_trx` | `99.3` | Column Expressions A |
| `match_cont` / `sin_match_cont` | `405` / `7` | ATMs con/sin match Switch vs T24 |
| `pct_cont` | `98.3` | Column Expressions A |
| `pct_conciliacion` | `99.9` | Column Expressions A: % de casos que no pasan de 48 h |
| `total_casos` | `15275` | Column Expressions A: suma de las 5 categorías |
| `n_exacto` | `14890` | Conteo: match exacto |
| `n_reverso` | `198` | Conteo: match con reverso dentro del ATM |
| `n_pend_t24` | `120` | Conteo: pendiente en T24 por corte |
| `n_falt_switch` | `64` | Conteo: faltante en Switch por corte |
| `n_investigar` | `3` | Conteo: sin conciliar más de 48 h |
| `p_exacto` … `p_investigar` (5) | `96`, `1`, `1`, `1`, `1` | Column Expressions A (mínimo 1 %, suman 100) |
| `barra_html` | `<td width="96%" …>` | Column Expressions B (sección 2) |
| `atm_en_linea` | `398` | Monitoreo |
| `atm_fuera_linea` | `14` | Monitoreo |
| `advertencias` | `9` | Monitoreo |
| `criticos` | `0` | Monitoreo |
| `eventos_html` | `<tr><td>ATM-TGU-031</td>…</tr>` | Column Expressions + GroupBy (sección 5) |
| `url_dashboard` | `https://…/atm` | URL interna del dashboard (botón final). Fija, puede ir como String Configuration |

Son **31 variables**. Parámetro aparte (no va en la plantilla): `umbral_investigar` (Integer, N = 5) para el Rule Engine.

## Configurar el nodo Send Email

1. Antes de **Table Row to Variable**, pasa todas las columnas numéricas por **Number To String**. Los nombres de columna deben coincidir con la tabla de arriba.
2. Conecta el puerto de variables de **Table Row to Variable** a **Send Email**.
3. En Send Email:
   - *Text format*: **HTML**.
   - *Text*: pega todo el contenido de `alerta_atm_template.html`.
   - *Subject* (sugerido): `[$${Sestado_texto}$$] Salud de ATMs – corte $${Sfecha_corte}$$`
4. Si las tildes de las variables salen mal en Outlook, cámbialas por entidades HTML (`&oacute;`, `&Iacute;`…).
   El texto fijo de la plantilla ya usa entidades.

### Logo (opcional)
La plantilla funciona sin logo: el encabezado usa texto. El logo público de Ficohsa es **SVG**, y
**Outlook de escritorio no muestra SVG**, así que esa URL no sirve. Para agregar logo hace falta una versión **PNG** publicada en una URL pública oficial de Grupo Ficohsa.
Pon esa URL en el bloque comentado `LOGO OPCIONAL` del encabezado y descoméntalo.
No uses SharePoint, rutas internas ni base64. Outlook bloquea imágenes externas por defecto, así que el `alt` debe bastar.

## Vista previa y validación (local, no envía correos)

```bash
python preview.py                      # escenario "atencion" -> preview.html
python preview.py --escenario critico  # ok | atencion | critico
python preview.py --todos              # preview_ok.html, preview_atencion.html, preview_critico.html
python preview.py --test               # casos límite de la regla de porcentajes

python check_outlook.py                                     # valida la plantilla
python check_outlook.py --estricto preview_critico.html     # valida un preview ya armado
```

`preview.py` avisa de variables que están en la plantilla y no en el JSON, y al revés.
También marca como error cualquier valor con secuencias de 12 a 19 dígitos (posible tarjeta o cuenta).
Abre los `preview*.html` en el navegador. La validación final en Outlook real queda a criterio del equipo, según su procedimiento de pruebas.

## Seguridad (ISO 27001)
- Solo agregados e identificadores de ATM: sin tarjetas, cuentas ni nombres de clientes.
- Sin imágenes externas activas, sin JavaScript, sin base64.
- Los datos de `sample_vars.json` son ficticios.
