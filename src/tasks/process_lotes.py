import os
from typing import Optional
from src.config import Settings
from src.excel.interop import ensure_excel, open_workbook, close_workbook, quit_excel
from src.excel.utils import obtener_valor, agregar_fila, copiar_formula, obtener_siguiente_consecutivo

xlUp = -4162

def _int_or_zero(v):
    try:
        return int(v or 0)
    except Exception:
        return 0

def run_process(cfg: Settings) -> None:
    excel = ensure_excel(cfg.excel_visible)
    try:
        wb = open_workbook(excel, cfg.archivo_original)
        tabla_ws = wb.Worksheets(cfg.hoja_tabla)
        tabla = tabla_ws.ListObjects(cfg.nombre_tabla)

        # Determinar rango de filas
        total = tabla.ListRows.Count
        start = cfg.fila_inicio
        end = cfg.fila_fin if cfg.fila_fin is not None else total
        end = min(end, total)

        # Referencias constantes a hojas
        orden_ws = wb.Worksheets(cfg.hoja_orden)
        calidad_ws = wb.Worksheets(cfg.hoja_calidad)
        registro_ws = wb.Worksheets('Registro')
        cert_ws = wb.Worksheets('CERTCALIDAD')
        trazab_ws = wb.Worksheets('TRAZAB')

        for i in range(start, end + 1):
            fila = tabla.ListRows(i).Range
            lote = fila.Cells(1, 1).Value
            cantidad = fila.Cells(1, 2).Value
            inyeccion_txt = fila.Cells(1, 3).Value
            corte_txt = fila.Cells(1, 4).Value
            acondic_txt = fila.Cells(1, 5).Value

            wb.Names('dato3').RefersToRange.Value = lote
            orden_ws.Range('AD14').Value = cantidad

            if inyeccion_txt is not None and str(inyeccion_txt).strip() != "":
                orden_ws.Range('E15').Value = inyeccion_txt
            if corte_txt is not None and str(corte_txt).strip() != "":
                orden_ws.Range('R15').Value = corte_txt
            if acondic_txt is not None and str(acondic_txt).strip() != "":
                orden_ws.Range('AE15').Value = acondic_txt

            # === Libro de colores ===
            color = _int_or_zero(orden_ws.Range('AS29').Value)
            peso = orden_ws.Range('AS26').Value
            mes = _int_or_zero(orden_ws.Range('T10').Value)
            ano_short = _int_or_zero(orden_ws.Range('V10').Value)
            ano_full = 2000 + ano_short

            prefijo = f"{chr(ord('E') + (ano_short - 23))}COLORES V3"
            mes_str = f"{mes:02d}"
            nombre_color_file = f"{prefijo} - {mes_str} {ano_full}.xlsx"
            ruta_color_file = os.path.join(cfg.ruta_colores, nombre_color_file)

            wb_col = excel.Workbooks.Open(ruta_color_file)
            ws_col = wb_col.Worksheets(1)

            fila_peso = 3 + color
            ws_col.Range(f"AH{fila_peso}").Value = peso

            col1_num = 57 + (color - 1) * 8
            col2_num = col1_num + 5

            rng1 = ws_col.Range(ws_col.Cells(3, col1_num), ws_col.Cells(8, col1_num))
            rng2 = ws_col.Range(ws_col.Cells(3, col2_num), ws_col.Cells(8, col2_num))

            dest1 = orden_ws.Range('L54:L59')
            dest2 = orden_ws.Range('Q54:Q59')
            dest1.Value = rng1.Value
            dest2.Value = rng2.Value
            wb_col.Close(SaveChanges=False)

            # === REGISTRO ===
            consecutivo = obtener_siguiente_consecutivo(registro_ws)
            fila_registro = [
                consecutivo,
                obtener_valor(wb, 'dato1'),
                obtener_valor(wb, 'dato2'),
                obtener_valor(wb, 'dato3'),
                orden_ws.Range('AR22').Value,
                None,
                orden_ws.Range('AE16').Value,
                orden_ws.Range('AE16').Value,
                (orden_ws.Range('AD14').Value or 0) * 42,
            ] + [None]*10 + [orden_ws.Range('AE16').Value]

            while len(fila_registro) < 25:
                fila_registro.append(None)

            last_row_prev = registro_ws.Cells(registro_ws.Rows.Count, 1).End(xlUp).Row
            agregar_fila(registro_ws, fila_registro)
            last_row_new = last_row_prev + 1

            cols_formula = [6, 10, 16, 17, 22, 23]
            for col in cols_formula:
                copiar_formula(registro_ws, col, last_row_prev, last_row_new)

            # === CERTCALIDAD ===
            last_row = registro_ws.Cells(registro_ws.Rows.Count, 1).End(xlUp).Row
            cert_no = registro_ws.Cells(last_row, 1).Value
            referencia = registro_ws.Cells(last_row, 3).Value
            lote_val = registro_ws.Cells(last_row, 4).Value
            fecha_fab = registro_ws.Cells(last_row, 5).Value
            fecha_venc = registro_ws.Cells(last_row, 6).Value
            fecha_inspec = registro_ws.Cells(last_row, 7).Value
            fecha_exp_cert = registro_ws.Cells(last_row, 8).Value
            nac_1_0 = registro_ws.Cells(last_row, 16).Value
            pasa_no_pasa = registro_ws.Cells(last_row, 17).Value

            cert_ws.Range('B13').Value = cert_no
            cert_ws.Range('E13').Value = referencia
            cert_ws.Range('H13').Value = lote_val
            cert_ws.Range('B15').Value = fecha_fab
            cert_ws.Range('C15').Value = fecha_venc
            cert_ws.Range('E15').Value = fecha_inspec
            cert_ws.Range('H15').Value = fecha_exp_cert
            cert_ws.Range('F22:F26').Value = tuple([nac_1_0]*5)
            cert_ws.Range('H22:H26').Value = tuple([pasa_no_pasa]*5)

            cert_no = cert_ws.Range('B13').Value
            lote_val = cert_ws.Range('H13').Value
            nombre_pdf_cert = f"{cert_no} Certificado de Calidad {lote_val}.pdf"
            ruta_pdf_cert = os.path.join(cfg.ruta_pdf_certificados, nombre_pdf_cert)
            cert_ws.ExportAsFixedFormat(
                Type=0, Filename=ruta_pdf_cert, Quality=0,
                IncludeDocProperties=True, IgnorePrintAreas=False,
                OpenAfterPublish=False
            )

            # === ENTRADAS EN OTRAS HOJAS ===
            from_rows = [
                ('Inyeccion', [
                    orden_ws.Range('E16').Value,
                    orden_ws.Range('F13').Value,
                    orden_ws.Range('D14').Value,
                    orden_ws.Range('AQ16').Value,
                    obtener_valor(wb, 'dato3'),
                    None,
                    orden_ws.Range('E15').Value
                ]),
                ('Corte', [
                    orden_ws.Range('R16').Value,
                    orden_ws.Range('S13').Value,
                    orden_ws.Range('Q14').Value,
                    orden_ws.Range('AQ16').Value,
                    obtener_valor(wb, 'dato3'),
                    None,
                    orden_ws.Range('R15').Value
                ]),
                ('Acondic', [
                    orden_ws.Range('AE16').Value,
                    orden_ws.Range('AF13').Value,
                    orden_ws.Range('AD14').Value,
                    orden_ws.Range('AQ16').Value,
                    obtener_valor(wb, 'dato3'),
                    None,
                    orden_ws.Range('AE15').Value
                ]),
                ('Etiquetas', [
                    orden_ws.Range('AE16').Value,
                    obtener_valor(wb, 'dato1'),
                    (orden_ws.Range('AD14').Value or 0) + 2,
                    obtener_valor(wb, 'dato3')
                ]),
            ]

            for sheet_name, values in from_rows:
                agregar_fila(wb.Worksheets(sheet_name), values)

            # === TRAZAB ===
            fila_trazab = [
                orden_ws.Range('AR22').Value,
                obtener_valor(wb, 'dato3'),
                orden_ws.Range('F11').Value,
                orden_ws.Range('AD14').Value
            ]
            agregar_fila(trazab_ws, fila_trazab)

            last_row_trazab = trazab_ws.Cells(trazab_ws.Rows.Count, 1).End(xlUp).Row

            vals_l = [orden_ws.Range(f'L{54+i}').Value for i in range(6)]
            vals_q = [orden_ws.Range(f'Q{54+i}').Value for i in range(6)]
            intercalado = []
            for l, q in zip(vals_l, vals_q):
                intercalado.append(l)
                intercalado.append(q)

            for idx, val in enumerate(intercalado):
                trazab_ws.Cells(last_row_trazab, 5 + idx).Value = val

            trazab_ws.Cells(last_row_trazab, 17).Value = orden_ws.Range('E15').Value
            trazab_ws.Cells(last_row_trazab, 18).Value = orden_ws.Range('R15').Value
            trazab_ws.Cells(last_row_trazab, 19).Value = orden_ws.Range('AE15').Value

            # === GUARDAR Y EXPORTAR ===
            nombre_archivo = orden_ws.Range('H10').Value
            if not nombre_archivo:
                raise ValueError("La celda H10 está vacía. No se puede continuar.")

            nombre_salida = f"Orden Produccion LT {nombre_archivo}"
            nombre_salida_calidad = f"Calidad Durante Producción LT {nombre_archivo}"

            nuevo_excel_path = os.path.join(cfg.ruta_excel_salida, f"{nombre_salida}.xlsm")
            nuevo_excel_path_calidad = os.path.join(cfg.ruta_excel_salida, f"{nombre_salida_calidad}.xlsm")

            ruta_pdf_orden = os.path.join(cfg.ruta_pdf_salida, f"{nombre_salida}.pdf")
            ruta_pdf_calidad = os.path.join(cfg.ruta_pdf_salida, f"{nombre_salida} CALIDAD.pdf")

            if os.path.exists(nuevo_excel_path):
                os.remove(nuevo_excel_path)
            orden_ws.Copy()
            temp_wb = excel.ActiveWorkbook
            temp_wb.SaveAs(nuevo_excel_path, FileFormat=52)
            temp_wb.Close(SaveChanges=False)

            if os.path.exists(nuevo_excel_path_calidad):
                os.remove(nuevo_excel_path_calidad)
            calidad_ws.Copy()
            temp_wb = excel.ActiveWorkbook
            temp_wb.SaveAs(nuevo_excel_path_calidad, FileFormat=52)
            temp_wb.Close(SaveChanges=False)

            orden_ws.ExportAsFixedFormat(
                Type=0, Filename=ruta_pdf_orden, Quality=0,
                IncludeDocProperties=True, IgnorePrintAreas=False,
                OpenAfterPublish=False
            )

            calidad_ws.ExportAsFixedFormat(
                Type=0, Filename=ruta_pdf_calidad, Quality=0,
                IncludeDocProperties=True, IgnorePrintAreas=False,
                OpenAfterPublish=False
            )

            wb.Save()

        close_workbook(wb, save=False)
    finally:
        try:
            quit_excel(excel)
        except Exception:
            pass
