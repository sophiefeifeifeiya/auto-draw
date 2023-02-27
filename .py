import win32com.client as win32
excel =  win32.DispatchEx('Excel.Application')
excel.Application.Quit() 