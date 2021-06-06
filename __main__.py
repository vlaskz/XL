from openpyxl import load_workbook
import time
import Manipulation as man, Const_Bradesco as brd #some people certainly will brag about this.

#converts xls worksheets to xlsx - ok
#Set erase_source_files as False to prevent it. Useful in dev stage. Copy the
#same files over and over again isn't something I appreciate.
print(time.ctime(),'[main]: starting data workbook conversion')
man.convert_xls_xlsx('c:\\users\\coder\\Desktop\\XL\\VE09.xls',False)
man.convert_xls_xlsx('c:\\users\\coder\\Desktop\\XL\\AV09.xls',False)
man.convert_xls_xlsx('c:\\users\\coder\\Desktop\\XL\\VE28.xls',False)
man.convert_xls_xlsx('c:\\users\\coder\\Desktop\\XL\\AV28.xls',False)
print(time.ctime(),'[main]: finishing workbook conversion')

#loads the new datafiles into memory - ok
print(time.ctime(),'[main]: starting workbook data loading')
VE09 = load_workbook('VE09.xlsx')
AV09 = load_workbook('AV09.xlsx')
VE28 = load_workbook('VE28.xlsx')
AV28 = load_workbook('AV28.xlsx')
RECB = load_workbook('receber.xlsx', read_only=False)
print(time.ctime(),'[main]: finishing workbook data loading')

#clears the old data from the worksheets.
print(time.ctime(),'[main]: starting cleaning data from main workbook')
man.purge_data(RECB['VE28BRAD'],brd.ROW_INI,VE09.active.max_row,brd.COL_INI,brd.COL_FIN)
man.purge_data(RECB['AV09BRAD'],brd.ROW_INI,AV09.active.max_row,brd.COL_INI,brd.COL_FIN)
man.purge_data(RECB['VE09BRAD'],brd.ROW_INI,VE28.active.max_row,brd.COL_INI,brd.COL_FIN)
man.purge_data(RECB['AV28BRAD'],brd.ROW_INI,AV28.active.max_row,brd.COL_INI,brd.COL_FIN)
print(time.ctime(),'[main]: finishing cleaning data from main workbook')

#copy data from datafiles to main workbook
print(time.ctime(),'[main]: finishing copying data from data workbooks to main workbook')
man.fetch_data(VE09.active,RECB['VE09BRAD'],brd.ROW_INI,VE09.active.max_row,brd.COL_INI,brd.COL_FIN)
man.fetch_data(AV09.active,RECB['AV09BRAD'],brd.ROW_INI,AV09.active.max_row,brd.COL_INI,brd.COL_FIN)
man.fetch_data(VE28.active,RECB['VE28BRAD'],brd.ROW_INI,VE28.active.max_row,brd.COL_INI,brd.COL_FIN)
man.fetch_data(AV28.active,RECB['AV28BRAD'],brd.ROW_INI,AV28.active.max_row,brd.COL_INI,brd.COL_FIN)
print(time.ctime(),'[main]: finishing copying data from data workbooks to main workbook')

#part two: let's unbloat and unjunk the data (yes I'm a neological person)
print(time.ctime(),'[main]: starting unbloating data')
man.clean_data(RECB['VE09BRAD'], brd.ROW_INI, RECB['VE09BRAD'].max_row)
man.clean_data(RECB['AV09BRAD'], brd.ROW_INI, RECB['AV09BRAD'].max_row)
man.clean_data(RECB['VE28BRAD'], brd.ROW_INI, RECB['VE28BRAD'].max_row)
man.clean_data(RECB['AV28BRAD'], brd.ROW_INI, RECB['AV28BRAD'].max_row)
print(time.ctime(),'[main]: finishing unbloating data')

#format data in the "O" column as float and set the field as currency (at the moment is BRL(R$))
print(time.ctime(),'[main]: starting currency data formatting in main workbook')
man.format_currency_data(RECB['VE09BRAD'],brd.CUR_COL,brd.ROW_INI,RECB['VE09BRAD'].max_row)
man.format_currency_data(RECB['AV09BRAD'],brd.CUR_COL,brd.ROW_INI,RECB['AV09BRAD'].max_row)
man.format_currency_data(RECB['VE28BRAD'],brd.CUR_COL,brd.ROW_INI,RECB['VE28BRAD'].max_row)
man.format_currency_data(RECB['AV28BRAD'],brd.CUR_COL,brd.ROW_INI,RECB['AV28BRAD'].max_row)
print(time.ctime(),'[main]: finishing currency data formatting in main workbook')

print(time.ctime(),'[main]: saving main workbook')
RECB.save(filename='receber.xlsx')

print(time.ctime(),'[main]: opening main workbook')
man.open_in_excel(filepath='c:\\users\\coder\\Desktop\\XL\\receber.xlsx')

#Part 3: Send receber.xlsx to email.

#sends emails to mail_listing
print(time.ctime(),'[main]: starting sending main workbook to mailing list')
#man.send_mail()