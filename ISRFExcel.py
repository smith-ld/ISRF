import openpyxl
import datetime
import PersonObject as p
import pdfrw as pdf
import os


NY_CITIES = []
datetimelocations = {3, 4} # indexing for excel. For ISRF form is -1 these locations.

class ISRFExcel:

    def __init__(self):
        self._workbook = None
        self._current_worksheet = None
        self._responses = p.SingletonPersons()



    def load_excel_file(self, excel_filename):
        self._workbook = openpyxl.load_workbook(filename=excel_filename)
        self._current_worksheet = self._workbook.worksheets[0]



    def organize_form_responses(self):
        for row in self._current_worksheet.iter_rows(min_row=2):
            person = p.PersonObject()
            person.update_name(row[1].value, row[2].value)
            person.update_dates(row[3].value, row[4].value)
            c = [x.value.title() for x in row[5:7]]
            c.append(str(int(row[8].value)))
            person.update_entire_address(c)
            mob = self.clean_phone_numbers(row[9].value)
            home = self.clean_phone_numbers(row[10].value)
            emer = self.clean_phone_numbers(row[13].value)
            person.update_phone_numbers([mob, home, emer])
            person.update_email(row[11].value)
            person.update_emergency_contact(row[12].value)
            person.update_gender(row[14].value)
            person.update_latino(row[15].value)
            person.update_ethnicity(row[16].value)
            person.update_employment(row[17].value)
            c = [x.value for x in row[19:22]]
            person.update_us_studies(row[18].value, c)
            person.update_oconus_studies(row[22].value, row[23].value, row[24].value)
            person.update_dependents(row[25].value, row[26].value,row[27].value,row[28].value,row[29].value,row[30].value)
            person.update_learning_barriers(row[31].value)
            self._responses.add_person(person)
            # for cell in range(len(row)):
            #     #print(row[cell].value, str(cell))
            #
            #     # if cell in datetimelocations:
            #     #     date = row[cell].value
            #     #     print(date)
            # print()

    def clean_phone_numbers(self, phone):
        # print(phone)
        t = type(phone)
        phone_nums = []
        try:
            if t == str:
                #print(phone)
                if phone == '' or len(phone) == 1 or ord(phone[0]) > 57:
                    return [None]
                phone = phone.replace("-", "")
            else:
                phone = str(int(phone))

            # print(phone)
            phone_nums.append(phone[0:3])
            phone_nums.append(phone[3:6])
            phone_nums.append(phone[6:])

            return phone_nums

        except:
            return [None]


    def make_forms(self, output_location):
        output_location = './'
        for person in self._responses.get_person_list():
            self.make_isrf(person)



    def make_isrf(self, person):
        persons_pdf = pdf.PdfReader("ISRF_V1 (10).pdf")
        Annots = persons_pdf.Root.AcroForm.Fields
        dates = []
        date = person.get_birthdate()
        date = str(date)
        info = '  {}  {}  {}  {}  {}  {}   {}  {}'.format(date[5], date[6], date[8], date[9], date[0], date[1], date[2], date[3])
        Annots[3].update(pdf.PdfDict(V = info, MaxLen=40)) #bday
        Annots[0].update(pdf.PdfDict(V = person.get_fullname()[0])) #fname
        Annots[2].update(pdf.PdfDict(V = person.get_fullname()[1])) #lname
        address = person.get_address()
        #print(address)
        Annots[5].update(pdf.PdfDict(V = address[0])) #add
        Annots[6].update(pdf.PdfDict(V = address[1])) #city
        Annots[7].update(pdf.PdfDict(V=' N  Y', MaxLen=8))  #state
        info = " " + "  ".join(address[2])
        info = info[:10] + " " + info[10:]
        Annots[8].update(pdf.PdfDict(V = info, MaxLen=30)) #zipcode
        #Annots[3].update(pdf.PdfDict(V = '  6  7  8 ', MaxLen=20))
        Annots[15].update(pdf.PdfDict(V = person.get_email()))
        phones = person.get_phone_numbers()
        # writing mobile
        Annots[12].update(pdf.PdfDict(V="  " + "  ".join(phones[0][0]), MaxLen=15))
        Annots[13].update(pdf.PdfDict(V="  " + "  ".join(phones[0][1]), MaxLen=15))
        Annots[14].update(pdf.PdfDict(V="  " + "  ".join(phones[0][2]), MaxLen=15))
        if phones[2] is not None:
            Annots[16].update(pdf.PdfDict(V="  " + "  ".join(phones[2][0]), MaxLen=15))
            Annots[17].update(pdf.PdfDict(V="  " + "  ".join(phones[2][1]), MaxLen=15))
            Annots[18].update(pdf.PdfDict(V="  " + "  ".join(phones[2][2]), MaxLen=15))
        if person.get_em_contact() is not None:
                Annots[19].update(pdf.PdfDict(V=person.get_em_contact()))
        if person.get_gender() == "Male":
            Annots[75].update(pdf.PdfDict(AS= pdf.PdfName("Yes")))
        elif person.get_gender() == 'Female':
            Annots[76].update(pdf.PdfDict(AS= pdf.PdfName("Yes")))
        else:
            Annots[77].update(pdf.PdfDict(AS= pdf.PdfName("Yes")))
        if person.get_latinoa():
            Annots[30].update(pdf.PdfDict(AS = pdf.PdfName("On")))
        else:
            Annots[31].update(pdf.PdfDict(AS = pdf.PdfName("On")))
        ethnicities = person.get_ethnicities()
        print(ethnicities)
        if ethnicities.__contains__('Native Hawaiian'):
            Annots[32].update(pdf.PdfDict(AS = pdf.PdfName("On")))
        if ethnicities.__contains__('Native American'):
            Annots[33].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        if ethnicities.__contains__('Alaskan Native'):
            Annots[34].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        if ethnicities.__contains__('Asian'):
            Annots[35].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        if ethnicities.__contains__('Pacific Islander'):
            Annots[36].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        if ethnicities.__contains__('African American'):
            Annots[37].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        if ethnicities.__contains__('Afro-Caribbean'):
            Annots[38].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        if ethnicities.__contains__('African'):
            Annots[39].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        if ethnicities.__contains__('Latinoa'):
            Annots[40].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        if ethnicities.__contains__('White [not Latinoa]'):
            Annots[41].update(pdf.PdfDict(AS=pdf.PdfName("On")))

        work_declaration = person.get_working_declaration()
        if work_declaration == 'FT':
            Annots[23].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        elif work_declaration == 'PT':
            Annots[24].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        elif work_declaration == 'NT':
            Annots[25].update(pdf.PdfDict(AS=pdf.PdfName("On"), ))
        elif work_declaration == 'UE':
            Annots[27].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        elif work_declaration == 'UA':
            Annots[28].update(pdf.PdfDict(AS=pdf.PdfName("On")))

        if person.get_studied_in_us():
            Annots[46].update(pdf.PdfDict(AS=pdf.PdfName("On")))
            Annots[43].update(pdf.PdfDict(V=str(person.get_highest_us_grade())))
            ny_grade = person.get_highest_ny_grade()
            try:
                ny_grade = int(ny_grade)
                Annots[44].update(pdf.PdfDict(V = str(ny_grade)))
                ny_school = person.get_ny_school()
                if ny_school is not None:
                    Annots[45].update(pdf.PdfDict(V = str(ny_school)))
            except:
                print('Error on student\'s NY grade: could not parse: {}' \
                      'for {}'.format(str(ny_grade), person.get_fullname()))

        else:
            Annots[47].update(pdf.PdfDict(AS=pdf.PdfName("On")))
            country_hse = person.get_finished_hs()
            country_uni = person.get_finished_uni()
            if country_hse:
                Annots[50].update(pdf.PdfDict(AS=pdf.PdfName("On")))
            if country_uni:
                Annots[51].update(pdf.PdfDict(AS=pdf.PdfName("On")))
            country_years = person.get_country_years()
            if country_years is not None:
                Annots[52].update(pdf.PdfDict(V = str(country_years)))

        hasDependents = person.get_dependent_status()

        if not hasDependents:
            Annots[96].update(pdf.PdfDict(AS=pdf.PdfName("On")))
            Annots[98].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        else:
            Annots[95].update(pdf.PdfDict(AS=pdf.PdfName("On")))
            if person.single_parent():
                Annots[97].update(pdf.PdfDict(AS=pdf.PdfName("On")))
            else:
                Annots[98].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        dependents = person.get_dependents()
        for i in range(len(dependents)):
            if dependents[i] is not None:
                    Annots[53+i].update(pdf.PdfDict(V = str(dependents[i])))


        learning_barriers = person.get_learning_barriers()
        #print(learning_barriers)
        if learning_barriers.__contains__('Homeless or living in a shelter'):
            Annots[78].update(pdf.PdfDict(AS=pdf.PdfName("On"), V = 'On'))
        else:
            Annots[57].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        if learning_barriers.__contains__('You used to take care of the home or children'):
            Annots[80].update(pdf.PdfDict(AS=pdf.PdfName("On"), V = 'Yes'))
        else:
            Annots[59].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        if learning_barriers.__contains__('Disabled'):
            Annots[81].update(pdf.PdfDict(AS=pdf.PdfName("On"), V = 'On'))
        else:
            Annots[60].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))

        if learning_barriers.__contains__('Low Income'):
            Annots[82].update(pdf.PdfDict(AS=pdf.PdfName("On"), V = 'On'))
        else:
            Annots[61].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        if learning_barriers.__contains__('You only work during certain seasons.'):
            Annots[83].update(pdf.PdfDict(AS=pdf.PdfName("On"), V = 'On'))
        else:
            Annots[62].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        if learning_barriers.__contains__('You have a learning disability.'):
            Annots[90].update(pdf.PdfDict(AS=pdf.PdfName("On"), V = 'On'))
        else:
            Annots[63].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))

        if learning_barriers.__contains__('You ran away from your home when you were a child or teenager.'):
            Annots[91].update(pdf.PdfDict(AS=pdf.PdfName("On"), V = 'On'))
        else:
            Annots[64].update(pdf.PdfDict(AS=pdf.PdfName("On"), V = 'On'))

        if learning_barriers.__contains__('English is NOT your first language.'):
            Annots[89].update(pdf.PdfDict(AS=pdf.PdfName("On"), V = 'On'))
        else:
            Annots[88].update(pdf.PdfDict(AS=pdf.PdfName("On"), V = 'On'))

        if learning_barriers.__contains__('You spent time in prison.'):
            Annots[93].update(pdf.PdfDict(AS=pdf.PdfName("On"), V = 'On'))
        else:
            Annots[66].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))

        if learning_barriers.__contains__('You used to be in foster care.'):
            Annots[94].update(pdf.PdfDict(AS=pdf.PdfName("On"), V = 'On'))
        else:
            Annots[67].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))


        if learning_barriers.__contains__('The educational system in your country was very different'):
            Annots[84].update(pdf.PdfDict(AS=pdf.PdfName("On"), V = 'On'))
        else:
            Annots[68].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))

        if learning_barriers.__contains__('You have been unemployed for many years.'):
            Annots[85].update(pdf.PdfDict(AS=pdf.PdfName("On"), V = 'On'))
        else:
            Annots[69].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))

        if learning_barriers.__contains__('Your TANF (Temporary Assistance for Needy Families) will end within the next two years.'):
            Annots[86].update(pdf.PdfDict(AS=pdf.PdfName("On"), V = 'On'))
        else:
            Annots[70].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))

        if learning_barriers.__contains__('Single Parent'):
            Annots[87].update(pdf.PdfDict(AS=pdf.PdfName("On"), V = 'On'))
        else:
            Annots[71].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))

        Annots[58].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        Annots[65].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))

        # Annots[72].update(pdf.PdfDict(V='Electronically Completed by CI worker: ' + 'LS'))
        name = person.get_fullname()
        hel = name[0] + name[1]+ '.pdf'
        pdf.PdfWriter().write(hel, persons_pdf)
