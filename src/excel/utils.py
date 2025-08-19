from typing import Any

xlUp = -4162

def obtener_valor(wb, nombre_definido: str):
    try:
        return wb.Names[nombre_definido].RefersToRange.Value
    except Exception:
        return None

def agregar_fila(ws, valores: list[Any]):
    last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    for j, val in enumerate(valores, start=1):
        ws.Cells(last_row, j).Value = val

def copiar_formula(ws, col: int, last_row: int, target_row: int):
    cell = ws.Cells(last_row, col)
    target = ws.Cells[target_row, col] if hasattr(ws, '__getitem__') else ws.Cells(target_row, col)  # safety
    target = ws.Cells(target_row, col)
    if cell.HasFormula:
        target.Formula = cell.Formula
    else:
        target.Value = cell.Value

def obtener_siguiente_consecutivo(ws) -> str:
    last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    if last_row <= 1:
        return "04-1099"
    else:
        last_value = ws.Cells(last_row, 1).Value
        if not last_value or not isinstance(last_value, str) or '-' not in last_value:
            return "04-1099"
        prefix, num = last_value.split('-')
        nuevo = int(num) + 1
        return f"{prefix}-{nuevo:04d}"
