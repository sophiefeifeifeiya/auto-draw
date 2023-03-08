import win32com.client as win32
# close the excel
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Application.Quit()
