import openpyxl
import datetime
import PersonObject as p
import pdfrw as pdf
import os

ANNOTATIONKEYS = {
    "M": 75,
    "F": 76,
    'NONBINARY': 77,
    "LATIONA": 30,
    "NONLATINOA": 31,
    'NH': 32, 'NA': 33, 'AN': 34, 'AS': 35, 'PI': 36, 'AA': 37, 'AC': 38, 'AF': 39,
    'L': 40, 'W': 41, 'FT': 23, 'PT': 24, 'NT': 25, 'UE': 27, 'UA': 28, 'STUDIEDINUS': 46,
    'NONUSSTUDY': 47, 'COUNTRYHSE': 50, 'COUNTRYUNI': 51, 'NODEP': 96, 'NOTSP': 98,
    'HASDEP': 95, 'SINGLEPARENT': 97, 'HOME': 78, 'HM': 80, 'D': 81
}
datetimelocations = {3, 4}  # indexing for excel. For ISRF form is -1 these locations.

GENDERMALE = ['MALE', 'MASCULINO', 'MÂLE']
GENDERFEMALE = ['FEMALE', 'FEMENINO', 'FEMELLE']
CITIES = {'nueva york': 'New York'}


class ISRFExcel:

    def __init__(self):
        self._workbook = None
        self._current_worksheet = None
        self._responses = p.SingletonPersons()

    def load_excel_file(self, excel_filename):
        self._workbook = openpyxl.load_workbook(filename=excel_filename)
        self._current_worksheet = self._workbook.worksheets[0]

    def english_employment_scrub(self, employment):
        if employment == 'I work 20 hours a week or more.':
            working_declaration = 'FT'
        elif employment == 'I work fewer than 20 hours a week.':
            working_declaration = 'PT'
        elif employment == 'I\'m working at the moment, but my job will end soon.':
            working_declaration = 'NT'
        elif employment == 'I am not working, but I\'m looking for a job and want to ' \
                           'start working as soon as possible.':
            working_declaration = 'UE'
        else:
            working_declaration = 'UA'

        return working_declaration

    def clean_city(self, city):
        if city.lower() in CITIES:
            return CITIES[city.lower()]
        return city

    def spanish_employment_scrub(self, employment):
        if employment == 'Trabajo 20 horas o más por semana.':
            working_declaration = 'FT'
        elif employment == 'Trabajo menos de 20 horas a la semana.':
            working_declaration = 'PT'
        elif employment == 'Estoy trabajando por ahora, pero mi empleo terminará pronto (quedaré desempleado).':
            working_declaration = 'NT'
        elif employment == 'Actualmente no trabajo, pero estoy en búsqueda de empleo y me gustaría ' \
                           'encontrar uno tan pronto como sea posible.':
            working_declaration = 'UE'
        else:
            working_declaration = 'UA'

        return working_declaration

    def spanish_ethnicity_scrub(self, ethnicity_list):
        # TODO - FIX SPANISH ETHNICITIES
        ethnicity_list = ethnicity_list.replace('Latino(a)', 'Latinoa').split(",")
        ethnicity_list = list(map(lambda x: x.strip(), ethnicity_list))
        temp_list = {'Native Hawaiian': 'NH',
                     'Nativo(a) Americano(a)': 'NA',
                     'Alaskan Native': 'AN',
                     'Asiático(a)': 'AF',
                     'Pacific Islander': 'PI',
                     'African American': 'AA',
                     'Afro-Caribbean': 'AC',
                     'African': 'AS',
                     'Latinoa': 'L', 'Latino(a)': 'L',
                     'Blanco [No latino(a)]': 'W'
                     }

        # Latino(a)
        # Asiático(a),  Blanco [No latino(a)]
        # Nativo(a) Americano(a)
        ethnicities = []

        for eth in ethnicity_list:
            if eth in temp_list.keys():
                ethnicities.append(temp_list[eth])
            else:
                print("\nPossible code values: " + str(temp_list.values()))
                code = input("Not recognized Spanish ethnicity: {} please enter the code manually: ".format(eth))
                ethnicities.append(code)

                with open('Adjustments.txt', 'a') as f:
                    adjustments = "Adjustment for SPANISH: "
                    st = "{}:'{}'".format(eth, code)
                    f.writelines(adjustments)
                    f.writelines(st)
        return ethnicities

    def spanish_learning_barriers_scrub(self, learning_barriers_list):
        barriers = learning_barriers_list.strip().split(",")
        # print(barriers)
        barriers = [x.strip() for x in barriers]
        items = []
        d = {
            'Sin hogar o viviendo en un refugio.': 'HOME',  # TODO
            'Solía hacerse cargo del hogar o de sus hijos, pero ahora debe encontrar un trabajo.': 'HM',
            'Posee alguna discapacidad.': 'D',
            'Bajos ingresos.': 'LI',
            'Sólo trabaja durante algunas temporadas.': 'MIG',
            'Posee alguna discapacidad de aprendizaje.': 'LD',
            'Dejó su hogar cuando era niño(a) o adolescente.': 'RA',
            'El inglés no es su idioma nativo.': 'ESL',
            'You spent time in prison.': 'EO',  # TODO
            'You used to be in foster care.': 'FC',  # TODO
            'El sistema educacional en su país es muy diferente o nunca estudió en su país.': 'CB',
            'or you never studied in your country.': 'CB',
            'Ha estado desempleado(a) por varios años.': 'UE',  # TODO
            'but now you must find a job.': 'TANF',  # TODO
            'Your TANF (Temporary Assistance for Needy Families) will end within the next two years.': 'TANF',  # TODO
            'Es padre soltero o madre soltera.': 'SP',
            'Solía hacerse cargo del hogar o de sus hijos': 'HM', 'pero ahora debe encontrar un trabajo.': 'HM',

        }
        for barrier in barriers:
            if barrier in d.keys():
                items.append(d[barrier])
            else:
                print("Possible code values: " + str(d.values()))
                code = input(
                    "Not recognized Spanish learning barrier: {} please enter the code manually: ".format(barrier))
                items.append(code)
                with open('Adjustments.txt', 'a') as f:
                    adjustments = "Adjustment for SPANISH: "
                    st = "{}:'{}'".format(barrier, code)
                    f.writelines(adjustments)
                    f.writelines(st)
        return items

    def english_ethnicity_scrub(self, ethnicity_list):
        ethnicity_list = ethnicity_list.replace('Latino(a)', 'Latinoa') \
            .replace('White [not Latino(a)]', 'White not Latinoa').split(",")
        temp_list = {'Native Hawaiian': 'NH',
                     'Native American': 'NA',
                     'Alaskan Native': 'AN',
                     'Asian': 'AF',
                     'Pacific Islander': 'PI',
                     'African American': 'AA',
                     'Afro-Caribbean': 'AC',
                     'African': 'AS',
                     'Latinoa': 'L',
                     'White [not Latinoa]': 'W'
                     }
        ethnicities = []
        for k, v in temp_list.items():
            if ethnicity_list.__contains__(k):
                ethnicities.append(temp_list[k])
        return ethnicities

    def english_learning_barriers_scrub(self, learning_barriers_list):
        barriers = learning_barriers_list.strip().split(",")
        # print(barriers)
        barriers = [x.strip() for x in barriers]
        items = []
        d = {
            'Homeless or living in a shelter': 'HOME',
            'You used to take care of the home or children': 'HM',
            'Disabled': 'D',
            'Low Income': 'LI',
            'You only work during certain seasons.': 'MIG',
            'You have a learning disability.': 'LD',
            'You ran away from your home when you were a child or teenager.': 'RA',
            'English is NOT your first language.': 'ESL',
            'You spent time in prison.': 'EO',
            'You used to be in foster care.': 'FC',
            'The educational system in your country was very different': 'CB',
            'or you never studied in your country.': 'CB',
            'You have been unemployed for many years.': 'UE',
            'but now you must find a job.': 'HM',
            'Your TANF (Temporary Assistance for Needy Families) will end within the next two years.': 'TANF',
            'Single Parent': 'SP'
        }
        for barrier in barriers:
            items.append(d[barrier])
        return items

    def english_gender_scrub(self, gender):
        if gender in GENDERMALE:
            return 'MALE'
        elif gender in GENDERFEMALE:
            return 'FEMALE'
        else:
            return 'NONBINARY'

    def organize_form_responses(self, start_row):
        for row in self._current_worksheet.iter_rows(min_row=2, max_row=83):

            person = p.PersonObject()
            person.update_name(row[1].value, row[2].value)
            person.update_dates(row[3].value, row[4].value)
            print(person.get_fullname())
            c = [x.value.title() for x in row[5:7]]
            try:
                c.append(str(int(row[8].value)))
            except TypeError:
                if row[8].value is None:
                    print(str(person.get_fullname()) + " had an error in the address. Recheck their input. "
                                                       "Every value must be filled. ")
                    continue
            person.update_entire_address(c)
            mob = self.clean_phone_numbers(row[9].value)
            home = self.clean_phone_numbers(row[10].value)
            emer = self.clean_phone_numbers(row[13].value)
            person.update_phone_numbers([mob, home, emer])
            person.update_email(row[11].value)
            person.update_emergency_contact(row[12].value)

            person.update_gender(self.english_gender_scrub(row[14].value.upper()))  # TODO EG
            person.update_latino(row[15].value)
            person.update_ethnicity(self.spanish_ethnicity_scrub(row[16].value))  # TODO ELB
            person.update_employment(self.spanish_employment_scrub(row[17].value))
            c = [x.value for x in row[19:22]]
            # print(c)
            person.update_us_studies(row[18].value, c)
            person.update_oconus_studies(row[22].value, row[23].value, row[24].value)
            person.update_dependents(row[25].value, row[26].value, row[27].value, row[28].value, row[29].value,
                                     row[30].value)
            print(row[31].value)
            person.update_learning_barriers(self.spanish_learning_barriers_scrub(row[31].value))  # TODO ELB
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
                # print(phone)
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

        for person in self._responses.get_person_list():
            self.make_isrf(person, output_location)

    def make_isrf(self, person, folder_name):
        persons_pdf = pdf.PdfReader("ISRF_V1 (10).pdf")
        Annots = persons_pdf.Root.AcroForm.Fields
        dates = []
        date = person.get_birthdate()
        # print(date)
        date = str(date)
        indeces = [5, 6, 8, 9, 0, 1, 2, 3]
        info = "  "
        length = len(indeces)
        for i in range(length):
            info += date[indeces[i]]
            if i < length:
                info += "  "
        print(date)
        # info = '  {}  {}  {}  {}  {}  {}   {}  {}'.format(date[5], date[6], date[8], date[9], date[0], date[1], date[2],
        #                                                   date[3])
        Annots[3].update(pdf.PdfDict(V=info, MaxLen=40))  # bday
        Annots[0].update(pdf.PdfDict(V=person.get_fullname()[0]))  # fname
        Annots[2].update(pdf.PdfDict(V=person.get_fullname()[1]))  # lname
        address = person.get_address()
        # print(address)
        Annots[5].update(pdf.PdfDict(V=address[0]))  # add
        Annots[6].update(pdf.PdfDict(V=self.clean_city(address[1])))  # city
        Annots[7].update(pdf.PdfDict(V=' N  Y', MaxLen=8))  # state
        info = " " + "  ".join(address[2])
        info = info[:10] + " " + info[10:]
        Annots[8].update(pdf.PdfDict(V=info, MaxLen=30))  # zipcode
        # Annots[3].update(pdf.PdfDict(V = '  6  7  8 ', MaxLen=20))
        Annots[15].update(pdf.PdfDict(V=person.get_email()))
        phones = person.get_phone_numbers()
        # writing mobile
        Annots[12].update(pdf.PdfDict(V="  " + "  ".join(phones[0][0]), MaxLen=15))
        Annots[13].update(pdf.PdfDict(V="  " + "  ".join(phones[0][1]), MaxLen=15))
        Annots[14].update(pdf.PdfDict(V="  " + "  ".join(phones[0][2]), MaxLen=15))
        if phones[2] is not None and len(phones[2]) > 0:
            try:
                Annots[16].update(pdf.PdfDict(V="  " + "  ".join(phones[2][0]), MaxLen=15))
                Annots[17].update(pdf.PdfDict(V="  " + "  ".join(phones[2][1]), MaxLen=15))
                Annots[18].update(pdf.PdfDict(V="  " + "  ".join(phones[2][2]), MaxLen=15))
            except TypeError:
                pass
        if person.get_em_contact() is not None:
            Annots[19].update(pdf.PdfDict(V=person.get_em_contact()))
        if person.get_gender() == "MALE":
            Annots[75].update(pdf.PdfDict(AS=pdf.PdfName("Yes")))
        elif person.get_gender() == 'FEMALE':
            Annots[76].update(pdf.PdfDict(AS=pdf.PdfName("Yes")))
        else:
            Annots[77].update(pdf.PdfDict(AS=pdf.PdfName("Yes")))
        if person.get_latinoa():
            Annots[30].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        else:
            Annots[31].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        ethnicities = person.get_ethnicities()
        # print(ethnicities)

        if ethnicities.__contains__('NH'):
            Annots[32].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        if ethnicities.__contains__('NA'):
            Annots[33].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        if ethnicities.__contains__('AN'):
            Annots[34].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        if ethnicities.__contains__('AS'):
            Annots[35].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        if ethnicities.__contains__('PI'):
            Annots[36].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        if ethnicities.__contains__('AA'):
            Annots[37].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        if ethnicities.__contains__('AC'):
            Annots[38].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        if ethnicities.__contains__('AF'):
            Annots[39].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        if ethnicities.__contains__('L'):
            Annots[40].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        if ethnicities.__contains__('W'):
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
            CONUS = True
            #
            Annots[43].update(pdf.PdfDict(V=str(person.get_highest_us_grade())))
            ny_grade = person.get_highest_ny_grade()
            try:
                ny_grade = int(ny_grade)
                Annots[44].update(pdf.PdfDict(V=str(ny_grade)))
                ny_school = person.get_ny_school()
                if ny_school is not None:
                    Annots[45].update(pdf.PdfDict(V=str(ny_school)))
            except:
                print('Error on student\'s NY grade: could not parse: {}' \
                      ' for {}'.format(str(ny_grade), person.get_fullname()))

        else:
            CONUS = False

        country_hse = person.get_finished_hs()
        country_uni = person.get_finished_uni()
        if country_hse or country_uni:
            if CONUS:
                Annots[46].update(pdf.PdfDict(AS=pdf.PdfName("On")))
            else:
                Annots[47].update(pdf.PdfDict(AS=pdf.PdfName("On")))
            if country_uni:
                Annots[51].update(pdf.PdfDict(AS=pdf.PdfName("On")))
            elif country_hse:
                Annots[50].update(pdf.PdfDict(AS=pdf.PdfName("On")))

        country_years = person.get_country_years()
        if country_years is not None:
            Annots[52].update(pdf.PdfDict(V=str(country_years)))

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
                Annots[53 + i].update(pdf.PdfDict(V=str(dependents[i])))

        learning_barriers = person.get_learning_barriers()
        # print(learning_barriers)
        if learning_barriers.__contains__('HOME'):
            Annots[78].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        else:
            Annots[57].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        if learning_barriers.__contains__('HM'):
            Annots[80].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='Yes'))
        else:
            Annots[59].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        if learning_barriers.__contains__('D'):
            Annots[81].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        else:
            Annots[60].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))

        if learning_barriers.__contains__('LI'):
            Annots[82].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        else:
            Annots[61].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        if learning_barriers.__contains__('MIG'):
            Annots[83].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        else:
            Annots[62].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        if learning_barriers.__contains__('LD'):
            Annots[90].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        else:
            Annots[63].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))

        if learning_barriers.__contains__('RA'):
            Annots[91].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        else:
            Annots[64].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))

        if learning_barriers.__contains__('ESL'):
            Annots[89].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        else:
            Annots[88].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))

        if learning_barriers.__contains__('EO'):
            Annots[93].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        else:
            Annots[66].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))

        if learning_barriers.__contains__('FC'):
            Annots[94].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        else:
            Annots[67].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))

        if learning_barriers.__contains__('CB'):
            Annots[84].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        else:
            Annots[68].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))

        if learning_barriers.__contains__('UE'):
            Annots[85].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        else:
            Annots[69].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))

        if learning_barriers.__contains__('TANF'):
            Annots[86].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        else:
            Annots[70].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))

        if learning_barriers.__contains__('SP') and person.single_parent():
            # print('is a single parent' + str(person.get_fullname()))
            Annots[87].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        else:
            Annots[71].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))

        Annots[58].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))
        Annots[65].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))

        # Annots[72].update(pdf.PdfDict(V='Electronically Completed by CI worker: ' + 'LS'))
        name = person.get_fullname()
        old_path = os.getcwd()
        path = os.path.join(os.getcwd(), folder_name)
        os.chdir(path)

        # print(os.getcwd())
        location = name[0] + " " + name[1] + ' ISRF.pdf'

        try:
            # person.Root.AcroForm.update(pdf.PdfDict(NeedAppearances=pdf.PdfObject('true')))
            persons_pdf.Root.AcroForm.update(pdf.PdfDict(NeedAppearances=pdf.PdfObject('true')))
        except AttributeError:
            print(str(name) + "  did not have the Root attribute")
        print(location)
        pdf.PdfWriter().write(location, persons_pdf)
        os.chdir(old_path)
