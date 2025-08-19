# Medical Device Traceability — Excel Office Automation (Win32 COM)

**Logro profesional:** Este proyecto fue diseñado y construido **para reducir hasta 2 semanas por ciclo de trabajo**, logrando **automatizar trabajo operativo de oficina en Excel** para un flujo de **trazabilidad de dispositivos médicos**. Centraliza los datos en una **tabla maestra** y desde allí **genera formatos, registra información, y emite informes y certificados de calidad** en Excel y PDF, reduciendo tiempo de ejecución y errores manuales.

## 🎯 Propósito
- **Trazabilidad de dispositivos médicos**: registro y actualización de información crítica a lo largo del ciclo de producción.
- **Estandarización de formatos**: creación de órdenes de producción y hojas de calidad desde plantillas.
- **Informes de calidad**: generación de **certificados de calidad en PDF** y actualización de libros auxiliares (por ejemplo, matrices de colores).
- **Automatización basada en tabla**: el sistema itera sobre filas de una **tabla fundamental** (listobject) en Excel y dispara toda la lógica.
- **Resultado**: proceso reproducible, rápido y con control de versiones .

## 🧠 Arquitectura técnica
- **Automatización Excel**: `pywin32` / `win32com` para controlar Excel (abrir libros, copiar hojas, escribir celdas, exportar a PDF).
- **Capa de utilidades**: helpers para insertar filas, copiar fórmulas y obtener valores nombrados.
- **Tareas**: módulo de **procesamiento por lotes** que ejecuta el pipeline completo (órdenes, calidad, etiquetas, trazabilidad, certificados).
- **Configuración**: centralizada en `src/config.py` (rutas, nombres de hojas/tablas, rango de filas).

> **Requisito**: Windows + Microsoft Excel instalado (COM).

## 🗂️ Estructura del repositorio
```
MedicalTraceability-ExcelAutomation/
├─ README.md
├─ requirements.txt
├─ .gitignore
├─ .env.example
├─ scripts/
│  └─ run_windows.bat
└─ src/
   ├─ app.py
   ├─ config.py
   ├─ excel/
   │  ├─ interop.py
   │  └─ utils.py
   ├─ tasks/
   │  └─ process_lotes.py
   └─ utils/
      └─ logger.py
```

## ⚙️ Configuración
Edita `src/config.py` o usa un `.env` (ver `.env.example`) para tus rutas y nombres de hojas/tablas.

- **Archivo maestro Excel** con macros y hojas base.
- **Tabla (ListObject)** `nombre_tabla` dentro de `hoja_tabla` de donde se leen los lotes y parámetros.
- **Rutas de salida** para Excel y PDF.
- **Rango de filas** a procesar (`fila_inicio`, `fila_fin`).

## ▶️ Ejecución
1. Crea y activa un entorno virtual, instala dependencias:
   ```bash
   pip install -r requirements.txt
   ```
2. Ajusta `src/config.py` o `.env`.
3. Ejecuta en Windows:
   ```bat
   scripts\run_windows.bat
   ```

## 🚀 Subir a GitHub (Desktop)
1. **File → New repository…** (elige carpeta del proyecto).
2. Copia todo el contenido del ZIP.
3. Commit inicial: `feat: initial office automation (medical traceability)`.
4. Publish repository (privado o público).

## 🧪 Detalles y consideraciones
- Uso de `ListObjects` (tablas) para índices y consistencia.
- Exportación a PDF con `ExportAsFixedFormat`.
- Copia de hojas a nuevos libros (`Worksheet.Copy`) con `FileFormat=52` (xlsm).
- Cálculo de **consecutivos** en formato `04-xxxx` a partir de la última fila.
- Integración con **libro de colores** para volcar matrices en posiciones específicas.
- Gestión de errores y cierre seguro de Excel.

## 🏅 Impacto
- Redujo drásticamente el trabajo manual repetitivo.
- Aumentó la trazabilidad y la **confiabilidad del registro**.
- Proyecto **que reduce hasta en 2 semanas por ciclo de trabajo**, con adopción inmediata por el equipo de calidad/producción.

## 📜 Licencia
MIT


