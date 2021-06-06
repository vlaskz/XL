#Disclaimer: man stands for manipulation, and not, I'll not argue with feminists..

#converts xls to xlsx using win32com library and excludes the xls files after successful conversion
def convert_xls_xlsx(xls, erase_source_files):
    import win32com.client as win32 
    import os  
    file = xls
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(file)
    wb.SaveAs(file+'x', FileFormat=51) #51 for xlsx/56 for xls
    wb.Close()
    excel.Application.Quit()
    if(erase_source_files):
        os.system('cmd /c del '+file)
        print('Source file has not been deleted.')
    print(file + ' has been converted to '+file+'x!')


#fills cells with no data.
#wb is workbook, sh is sheet, r is row, and c is column
def purge_data(sh, r_ini, r_fin, c_ini, c_fin):
    for r in range (r_ini,r_fin):
        for c in range (c_ini, c_fin):
            sh.cell(column=c, row=r).value = ""
    print(sh , ' has been purged!') 

#it does what it seems to do: fetch data.
def fetch_data(source, destination, r_ini, r_fin, c_ini, c_fin):
    for r in range(r_ini, r_fin):
        for c in range(c_ini, c_fin):
            destination.cell(row=r, column=c).value = source.cell(row=r, column=c).value
    print(destination.title + ' has been fed by ' + source.title)

#it cleans the already paid bills
def clean_data(ws, r_ini, r_fin):
    for r in range (r_ini, r_fin):
        s = ws.cell(row=r, column=1).value
        if(s.find('PAGO')!=-1):
            ws.delete_rows(r,1)
            print(ws.title, '-> Row ',r,' has been erased. Already paid.')
      