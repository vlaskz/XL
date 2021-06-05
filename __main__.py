from openpyxl import Workbook, load_workbook
import man, brad #some people certainly will brag about this.
#open the current working book.
receber = load_workbook('receber.xlsx', read_only=False)

#clears the old data from the workbook.
man.purge_data({receber['VE09BRAD'],receber['AV09BRAD'],receber['VE28BRAD'],receber['AV28BRAD']},brad.INITIAL_ROW, brad.FINAL_ROW, brad.INITIAL_COLUMN, brad.FINAL_COLUMN)

receber.save(filename='receber.xlsx')
#converts xls worksheets to xlsx

man.convert_xls_xlsx('c:\\users\\coder\\Desktop\\XL\\VE09.xls')
man.convert_xls_xlsx('c:\\users\\coder\\Desktop\\XL\\AV09.xls')
man.convert_xls_xlsx('c:\\users\\coder\\Desktop\\XL\\VE28.xls')
man.convert_xls_xlsx('c:\\users\\coder\\Desktop\\XL\\AV28.xls')

#loads the new datafiles into memory
ve09 = load_workbook('VE09.xlsx')
av09 = load_workbook('AV09.xlsx')
ve28 = load_workbook('VE28.xlsx')
av28 = load_workbook('AV28.xlsx')




