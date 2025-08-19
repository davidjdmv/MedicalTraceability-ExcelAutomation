import os
from dataclasses import dataclass
from dotenv import load_dotenv

load_dotenv()

def _get_bool(name: str, default: bool) -> bool:
    val = os.getenv(name)
    if val is None: return default
    return str(val).strip().lower() in ("1","true","yes","y")

def _get_int(name: str, default: int | None):
    val = os.getenv(name)
    if val is None or str(val).strip() == "": return default
    try:
        return int(val)
    except ValueError:
        return default

@dataclass
class Settings:
    excel_visible: bool = _get_bool("EXCEL_VISIBLE", False)

    archivo_original: str = os.getenv("ARCHIVO_ORIGINAL", r"C:\ruta\a\TRAZAB25.xlsm")
    hoja_orden: str = os.getenv("HOJA_ORDEN", "ORDENPROD")
    hoja_calidad: str = os.getenv("HOJA_CALIDAD", "Calidad Elast")
    hoja_tabla: str = os.getenv("HOJA_TABLA", "TABLA_GENERAL")
    nombre_tabla: str = os.getenv("NOMBRE_TABLA", "Tabla8")

    ruta_pdf_certificados: str = os.getenv("RUTA_PDF_CERTIFICADOS", r"C:\ruta\a\certificados")
    ruta_excel_salida: str = os.getenv("RUTA_EXCEL_SALIDA", r"C:\ruta\a\ordenes_excel")
    ruta_pdf_salida: str = os.getenv("RUTA_PDF_SALIDA", r"C:\ruta\a\ordenes_pdf")
    ruta_colores: str = os.getenv("RUTA_COLORES", r"C:\ruta\a\colores")

    fila_inicio: int = _get_int("FILA_INICIO", 1) or 1
    fila_fin: int | None = _get_int("FILA_FIN", 5)

settings = Settings()
