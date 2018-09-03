# import win32com.client as win32
import pyexcel as p

def xls_to_xlsx_converter(filename) :
#     fname = "C:\Users\Kiot\Documents\python_gui\input.xls"
#     excel = win32.gencache.EnsureDispatch('Excel.Application')
#     wb = excel.Workbooks.Open(fname)

#     wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
#     wb.Close()                               #FileFormat = 56 is for .xls extension
#     excel.Application.Quit()


    p.save_book_as(file_name=filename,dest_file_name=filename+'x')