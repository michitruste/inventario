"""
Configuración general. Es el ÚNICO archivo que deberías tener que editar
para adaptar el sistema a tus Excel reales.
"""
from pathlib import Path

BASE = Path(__file__).resolve().parent

# --- Rutas -------------------------------------------------------------------
DB_PATH = BASE / "datos" / "inventario.db"          # base de datos (un archivo)
FOTOS_DIR = BASE / "datos" / "fotos"                # aquí se guardan todas las fotos
# Carpeta(s) de la PC donde ya tienes fotos sueltas (se buscan por nombre de clave)
CARPETAS_FOTOS_EXTRA = [
    # Path(r"C:\Users\TU_USUARIO\Pictures\inventario"),
]

# --- Status ------------------------------------------------------------------
STATUS = ["missing", "to upload", "completed", "correction", "regresado al cliente"]
# Status que el sistema NUNCA cambia automáticamente (solo una persona)
STATUS_MANUALES = {"completed", "correction", "regresado al cliente"}

# --- Excel de catálogo (medidas + fotos) ----------------------------------------
# Nombre de la columna en TU Excel -> campo interno.
# Escribe los encabezados tal como aparecen (no importan mayúsculas ni acentos).
CATALOGO_COLUMNAS = {
    "descripcion": ["descripcion", "description", "item", "nombre"],
    "clave":       ["clave", "codigo", "code"],     # opcional: si existe, se usa directo
    "height":      ["height", "alto", "altura"],
    "length":      ["length", "lenght", "largo"],
    "breadth":     ["breadth", "ancho", "width"],
    "weight":      ["weight", "peso"],
    "imagen":      ["imagen", "image", "foto", "picture"],
}

# --- Excel de seguimiento (hojas con status) --------------------------------------
SEGUIMIENTO_COLUMNAS = {
    "hoja":        ["hoja", "sheet", "no. hoja", "numero hoja"],
    "serial":      ["serial", "no. de serie", "serie", "numero de serie"],
    "descripcion": ["descripcion", "description"],
    "costo":       ["costo", "cost", "precio"],
    "status":      ["status", "estado"],
}

# --- Tabla pegada desde la web ------------------------------------------------------
# Si la tabla copiada trae encabezados, se detectan con estas palabras.
# Si no los trae, se usa el orden de WEB_ORDEN_COLUMNAS.
WEB_COLUMNAS = {
    "serial":      ["serial", "serie", "no. de serie"],
    "descripcion": ["description", "descripcion"],
    "costo":       ["cost", "costo"],
    "height":      ["height"],
    "length":      ["length", "lenght"],
    "breadth":     ["breadth"],
    "weight":      ["weight"],
}
WEB_ORDEN_COLUMNAS = ["serial", "descripcion", "costo", "height", "length", "breadth", "weight"]
# Si la web ya trae medidas en una fila, ¿marcarla como completed al cargarla?
WEB_MEDIDAS_IMPLICA_COMPLETED = True

# --- Matching ------------------------------------------------------------------------
UMBRAL_ALTA = 90     # >= esto: confianza alta  (se pasa a "to upload" solo)
UMBRAL_MEDIA = 70    # >= esto: confianza media (sugerencia, revisar)
UMBRAL_MINIMO = 55   # por debajo de esto no se sugiere nada
AUTO_ASIGNAR_DESDE = UMBRAL_ALTA  # desde qué score se asigna sin confirmar
