import ISRFExcel

current = ISRFExcel.ISRFExcel()
current.load_excel_file("NYS ISRF Responses.xlsx")
current.organize_form_responses()
current.make_forms('here')





