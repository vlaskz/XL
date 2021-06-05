#man stands for manipulation

#converts xls to xlsx using win32com library and excludes the xls files after successful conversion
def convert_xls_xlsx(xls):
    import win32com.client as win32 
    import os  
    fname = xls
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)
    wb.SaveAs(fname+'x', FileFormat=51) #51 for xlsx/56 for xlsx
    wb.Close()
    excel.Application.Quit()
    os.system('cmd /c del '+fname)
    print(fname + ' has been converted to '+fname+'x!')


#fills cells with no data.
#wb is workbook, sh is sheet, r is row, and c is column

def purge(sh, r_ini, r_fin, c_ini, c_fin):
    for r in range (r_ini,r_fin):
        for c in range (c_ini, c_fin):
            sh.cell(column=c, row=r).value = ""
    print(sh , ' has been purged!')   