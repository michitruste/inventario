# Inventario: matching, app y subida a la web

Prototipo listo para usar con **datos de ejemplo inventados**. Para adaptarlo a tus datos reales, normalmente solo tienes que editar `config.py` (nombres de columnas) y `web_config.json` (URLs y selectores de la página).

## 1. Instalar (una sola vez)

Abre la consola (cmd) en esta carpeta y corre:

```
python -m pip install -r requirements.txt
python -m playwright install chromium
```

Opcional, para abrir la app con doble clic: crea con el Bloc de notas un archivo `iniciar_app.bat` en esta carpeta con estas dos líneas:

```
cd /d "%~dp0"
python -m streamlit run app.py
```

## 2. Probar con los ejemplos

```
python ejemplo\generar_ejemplo.py
python cli.py importar-catalogo ejemplo\catalogo_medidas.xlsx
python cli.py importar-seguimiento ejemplo\seguimiento.xlsx
python cli.py cargar-hoja 1003 ejemplo\hoja_1003.txt
python cli.py resumen
```

Luego abre la app con `python -m streamlit run app.py` (o con doble clic en `iniciar_app.bat` si lo creaste).

Si quieres empezar de cero, borra la carpeta `datos`.

## 3. Pasar a tus datos reales

1. Abre `config.py` y, en `CATALOGO_COLUMNAS` y `SEGUIMIENTO_COLUMNAS`, agrega los encabezados tal como aparecen en tus Excel. No importan mayúsculas ni acentos. Si tienes fotos sueltas en la PC, agrega esa carpeta en `CARPETAS_FOTOS_EXTRA`.
2. Borra la carpeta `datos` que dejaron los ejemplos.
3. Importa tus Excel desde la app (⚙️ Importar / Exportar) o desde la consola:
   ```
   python cli.py importar-catalogo "C:\...\medidas.xlsx"
   python cli.py importar-seguimiento "C:\...\seguimiento.xlsx"
   ```
   Las fotos se extraen solas, ya sea que estén **flotantes**, **dentro de la celda** (Excel 365) o escritas como **nombre de archivo**. Todas quedan en `datos/fotos`.

> Trabaja siempre sobre una **copia** de tus Excel. El sistema solo los lee, nunca los modifica.

## 4. Uso diario

| Antes | Ahora |
|---|---|
| Copiar la tabla de la web al Excel y limpiarla | **📥 Cargar hoja**: pegas la tabla y se limpia sola. O sin copiar nada: `python web_bot.py extraer 1003` |
| Ctrl+F por cada ítem | El match es automático. En **🔎 Revisar ítems** solo confirmas los 🟡/🔴 |
| Buscar la foto en la PC | Cada registro del catálogo ya tiene su foto |
| Subir a mano a la web | `python web_bot.py subir 1003`: llena todo y tú das Guardar |
| Actualizar status | Cambia solo (missing → to upload → completed) |
| Armar estadísticas | **📊 Dashboard** |

### Cómo funciona el match (confianza)

| Nivel | Ejemplo | Puntaje |
|---|---|---|
| Ya confirmado por una persona | misma descripción exacta | 100 ✅ |
| Prefijo + código | `ABCD-K7TR` | 100 🟢 |
| Código | `K7TR` | 97 🟢 |
| Nombre | `adaptador ethernet` | 95 🟢 |
| Código con 0/O o 1/I cambiados | `Q9Z0` ↔ `Q9ZO` | 88 🟡 |
| Solo categoría | `HVAC` genérico | 80 🟡 |
| Parecido (errores de dedo) | `Fuente de poder 500 W` ↔ `500W` | ≤85 🟡/🔴 |

- Con confianza **🟢 alta** y un registro que tenga 4 medidas + foto, el ítem pasa solo a **to upload**.
- Lo **🟡/🔴** espera a que lo confirmes con **"Usar este"**. Al confirmarlo, el sistema **aprende**: todos los ítems con ese mismo código o nombre, de esta hoja y de las que lleguen después, se resuelven solos.
- Los umbrales se ajustan en `config.py`.

### Status

- `missing`: no hay registro, falta confirmar el match, o el registro no tiene medidas o foto completas. El dashboard dice cuál de estas es la causa.
- `to upload`: listo para subir.
- `completed`, `correction` y `regresado al cliente`: **solo los cambia una persona** (o el bot, después de confirmar en la web). El matching nunca los toca.
- Si cambias la foto de un registro en 📚 Catálogo, puedes marcar como `correction` todos los ítems `completed` que la usan.

### 🎯 "Qué registrar primero"

En el dashboard (y en `python cli.py resumen`) aparecen los códigos en `missing` ordenados por cuántos ítems y cuánto costo desbloquean. Medir y fotografiar uno solo libera todos los ítems de esa fila.

## 5. Bot de la web (Playwright)

1. Edita `web_config.json` con las URLs y los selectores de tu página. Para encontrar los selectores: corre `playwright codegen https://TU-SITIO`, haz clic en los campos y copia lo que genera.
2. `python web_bot.py login`: se abre el navegador, **tú** inicias sesión y se guarda la sesión en `datos/sesion.json`. No compartas ese archivo: da acceso a tu cuenta.
3. `python web_bot.py extraer 1003`: lee la hoja directo de la web. Si un ítem ya tiene medidas en la web, lo marca `completed`.
4. `python web_bot.py subir 1003 --limite 3`: llena medidas y foto. Tú revisas en el navegador, das Guardar y presionas Enter en la consola.
5. Cuando ya confíes en él: `python web_bot.py subir 1003 --auto`. Da Guardar solo, espera la confirmación de la web y marca `completed`. Si algo falla, toma una captura en `datos/` y deja el ítem en `to upload`.

Para probarlo con la página simulada: `python web_bot.py extraer 1003 --config ejemplo\web_config_ejemplo.json`.

> Antes de automatizar la subida, confirma que tu empresa o el cliente lo permiten.

## 6. Varias personas

- **Opción A (recomendada):** una PC corre `streamlit run app.py --server.address 0.0.0.0`, y los demás abren `http://IP-DE-ESA-PC:8501` en su navegador. Hay una sola base de datos y no hay conflictos.
- **Opción B:** quien prefiera Excel usa **⚙️ Exportar** (sale con desplegable de status, colores y miniaturas), edita ahí y luego lo devuelve con **"Traer status editados en Excel"**.
- No pongas `inventario.db` en una carpeta compartida de red para que dos personas la abran a la vez. SQLite no se lleva bien con eso.
- **Respaldo:** copia `datos/` (la base de datos más las fotos).

## Archivos

| Archivo | Qué hace |
|---|---|
| `config.py` | **Lo único que editas**: columnas, rutas, umbrales |
| `app.py` | App web (Streamlit) |
| `cli.py` | Los mismos comandos, desde la consola |
| `parsing.py` | Limpia la tabla pegada, separa `HVAC \| xxxx-COD-xxx`, valida series |
| `matching.py` | Motor de match y aprendizaje |
| `excel_import.py` | Lee tus Excel y extrae las fotos |
| `exportar.py` | Excel de salida e importación de status editados |
| `reportes.py` | Cálculos del dashboard |
| `web_bot.py`, `web_config.json` | Playwright |
| `db.py` | Base de datos SQLite (`datos/inventario.db`; se puede abrir con *DB Browser for SQLite*) |
