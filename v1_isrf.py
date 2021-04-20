import ISRFExcel
import os
import sys
import time

if __name__ == "__main__":
    print("... Program Starting", end='')
    time.sleep(2)

    # if len(sys.argv) == 1:
    #     print("")
    #     print("If at any time you make an error, just type 'restart' to restart the questionnaire.")
    #     verbage = []
    #     placement = 0
    #     while True:
    #         intake = input(verbage[placement]).strip()
    #         if intake == 'restart':
    #             placement = 0
    #             continue
    #         elif placement == 0:
    #
    #
    #
    #
    #     time.sleep(3)
    # else:
    args = sys.argv
    current = ISRFExcel.ISRFExcel(None)
    # name = "ESOL French responses round 1.xlsx"
    current.load_excel_file(args[1])

    # folder_name = 'ESOL French ISRFs'
    folder_name = args[2]
    # current.organize_form_responses(2, 11, 'FRENCH')
    current.organize_form_responses(int(args[3]), int(args[4]), args[5])

    # excelFile FolderName StartRow EndRow Language TranslationsFile
    # 2 is default start row, to ignore header columns; otherwise start on last time ran + 1 to not recreate old forms.
    current.make_forms(folder_name)
