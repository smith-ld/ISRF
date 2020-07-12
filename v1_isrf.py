import ISRFExcel
import os

current = ISRFExcel.ISRFExcel()
current.load_excel_file("SPAN ISRF 1-83.xlsx")
folder_name = 'ESOL Spanish ISRFs'
current.organize_form_responses(24)
# 2 is default start row, to ignore header columns; otherwise start on last time ran + 1 to not recreate old forms.
current.make_forms(folder_name)
