import ISRFExcel
import os
import sys

if __name__ == "__main__":
    args = sys.argv
    current = ISRFExcel.ISRFExcel()
    # name = "ESOL French responses round 1.xlsx"
    current.load_excel_file(args[1])

    # folder_name = 'ESOL French ISRFs'
    folder_name = args[2]
    # current.organize_form_responses(2, 11, 'FRENCH')
    current.organize_form_responses(int(args[3]), int(args[4]), args[5])

    # excelFile FolderName StartRow EndRow Language
    # 2 is default start row, to ignore header columns; otherwise start on last time ran + 1 to not recreate old forms.
    current.make_forms(folder_name)
