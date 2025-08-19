import win32com.client as win32

xlUp = -4162

def ensure_excel(visible: bool = False):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = bool(visible)
    excel.DisplayAlerts = False
    return excel

def open_workbook(excel, path: str):
    return excel.Workbooks.Open(path)

def close_workbook(wb, save: bool = False):
    wb.Close(SaveChanges=bool(save))

def quit_excel(excel):
    excel.Quit()
