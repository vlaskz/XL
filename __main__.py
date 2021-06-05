from openpyxl import Workbook, load_workbook
import man, brad #some people certainly will brag about this.



#converts xls worksheets to xlsx - ok
man.convert_xls_xlsx('c:\\users\\coder\\Desktop\\XL\\VE09.xls',False)
man.convert_xls_xlsx('c:\\users\\coder\\Desktop\\XL\\AV09.xls',False)
man.convert_xls_xlsx('c:\\users\\coder\\Desktop\\XL\\VE28.xls',False)
man.convert_xls_xlsx('c:\\users\\coder\\Desktop\\XL\\AV28.xls',False)

#loads the new datafiles into memory - ok
VE09 = load_workbook('VE09.xlsx')
AV09 = load_workbook('AV09.xlsx')
VE28 = load_workbook('VE28.xlsx')
AV28 = load_workbook('AV28.xlsx')
RECB = load_workbook('receber.xlsx', read_only=False)

#clears the old data from the worksheets.
man.purge_data(RECB['VE09BRAD'],
                brad.INITIAL_ROW,
                brad.FINAL_ROW,
                brad.INITIAL_COLUMN,
                brad.FINAL_COLUMN)
man.purge_data(RECB['AV09BRAD'],
                brad.INITIAL_ROW,
                brad.FINAL_ROW,
                brad.INITIAL_COLUMN,
                brad.FINAL_COLUMN)
man.purge_data(RECB['VE28BRAD'],
                brad.INITIAL_ROW,
                brad.FINAL_ROW,
                brad.INITIAL_COLUMN,
                brad.FINAL_COLUMN)
man.purge_data(RECB['AV28BRAD'],
                brad.INITIAL_ROW,
                brad.FINAL_ROW,
                brad.INITIAL_COLUMN,
                brad.FINAL_COLUMN)
#copy data from datafiles to main workbook
man.fetch_data(VE09.active,
                RECB['VE09BRAD'],
                brad.INITIAL_ROW,
                brad.FINAL_ROW,
                brad.INITIAL_COLUMN,
                brad.FINAL_COLUMN)
man.fetch_data(AV09.active,
                RECB['AV09BRAD'],
                brad.INITIAL_ROW,
                brad.FINAL_ROW,
                brad.INITIAL_COLUMN,
                brad.FINAL_COLUMN)
man.fetch_data(VE28.active,
                RECB['VE28BRAD'],
                brad.INITIAL_ROW,
                brad.FINAL_ROW,
                brad.INITIAL_COLUMN,
                brad.FINAL_COLUMN)
man.fetch_data(AV28.active,
                RECB['AV28BRAD'],
                brad.INITIAL_ROW,
                brad.FINAL_ROW,
                brad.INITIAL_COLUMN,
                brad.FINAL_COLUMN)
RECB.save(filename='receber.xlsx')


