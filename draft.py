import openpyxl

bill_wc_form = openpyxl.load_workbook("resourses/bill_wc_form.xls")
bill_wc_form_edit = bill_wc_form.active

# Рахунок на оплату № NUMBER_OF_BILL від DATE_MAKE р.

bill_wc_form_edit["C12"] = fr"Рахунок на оплату № {numberofbill} від {datemake} р."
bill_wc_form_edit[]