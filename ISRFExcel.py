import openpyxl
import datetime
import PersonObject as p
import pdfrw as pdf
import os
import regex
import traceback

ANNOTATIONKEYS = {
    'FT': 23, 'PT': 24, 'NT': 25, 'UE': 27, 'UA': 28, 'NH': 32, 'NA': 33, 'AN': 34,
    'AS': 35, 'PI': 36, 'AA': 37, 'AC': 38, 'AF': 39, 'L': 40, 'W': 41
}
datetimelocations = {3, 4}  # indexing for excel. For ISRF form is -1 these locations.

GENDERMALE = ['MALE', 'MASCULINO', 'MÂLE']
GENDERFEMALE = ['FEMALE', 'FEMENINO', 'FEMELLE']
CITIES = {'nueva york': 'New York'}
NUM_WRITTEN = 0


class ISRFExcel:

    def __init__(self, translation_file):
        self._workbook = None
        self._current_worksheet = None
        self._responses = p.SingletonPersons()
        if translation_file is not None:
            self._translations = self.parse_translation_file(translation_file)
        else:
            self._translations = translation_file

    def parse_translation_file(self, t_file):
        with open(t_file, 'r') as file:
            lines = file.readlines()
            languages = lines[0].split()
            languages = [x.strip() for x in languages]
            # print(languages)
            i = 1
            lang_dict = {}
            current_lang_dict = {}
            current_lang = ""
            while i < len(lines):
                line = lines[i].strip()

                if line in languages:
                    current_lang = line
                    # print(current_lang)
                elif line == "":
                    # print("HERE")
                    # print(current_lang_dict)
                    # print(current_lang)
                    if current_lang != "":
                        lang_dict[current_lang] = current_lang_dict
                else:
                    # print(str(i)  + " " + line)
                    current_lang_dict[line] = lines[i + 1].strip()
                    i += 2
                    continue
                i += 1
            lang_dict[current_lang] = current_lang_dict
            return lang_dict

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

    def french_employment_scrub(self, employment):
        if employment == 'Je travaille 20 heures par semaine ou plus.':
            working_declaration = 'FT'
        elif employment == 'Je travaille moins de 20 heures par semaine.':
            working_declaration = 'PT'
        elif employment == 'Estoy trabajando por ahora, pero mi empleo terminará pronto (quedaré desempleado).':
            working_declaration = 'NT'
        elif employment == 'Je ne travaille pas, mais je suis à la recherche d\'' \
                           'un emploi et je veux commencer à travailler dès que possible.':
            working_declaration = 'UE'
        elif employment == 'Je ne travaille pas. Je ne veux pas trouver de travail en ce moment. ' \
                           'Il se peut que je cherche un emploi l\'année prochaine.':
            working_declaration = 'UA'
        else:
            print("Possible employment values are: {}, and they said {}. Please pick which one the student meant.".
                  format(employment,
                         str(["FT", 'PT', 'NT', 'UE', 'UA'])))

            working_declaration = input("Which one did they mean? ")
            with open('Adjustments.txt', 'a+') as f:
                f.writelines("French Employment needs adjustment: {} and code {}\n".format(employment,
                                                                                           working_declaration))
        return working_declaration

    def spanish_ethnicity_scrub(self, ethnicity_list):
        # TODO - FIX SPANISH ETHNICITIES
        ethnicity_list = ethnicity_list.replace('Latino(a)', 'Latinoa').split(",")
        ethnicity_list = list(map(lambda x: x.strip(), ethnicity_list))
        temp_list = {'Native Hawaiian': 'NH',
                     'Nativo(a) Americano(a)': 'NA',
                     'Alaskan Native': 'AN',
                     'Asiático(a)': 'AF',
                     'Isleño(a) del Pacífico (de Oceanía)': 'PI',
                     'Afro-Caribeño(a)': 'AC',
                     'Afroamericano(a)': 'AA',
                     'African': 'AF',
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
                    st = "{}:'{}'\n".format(eth, code)
                    f.writelines(adjustments)
                    f.writelines(st)
        return ethnicities

    def french_ethnicity_scrub(self, ethnicity_list):
        # TODO - FIX FRENCH ETHNICITIES
        ethnicity_list = ethnicity_list.strip().split(",")
        ethnicity_list = list(map(lambda x: x.strip(), ethnicity_list))
        temp_list = {'Native Hawaiian': 'NH',
                     'Nativo(a) Americano(a)': 'NA',
                     'Alaskan Native': 'AN',
                     'Africain': 'AF',
                     'Pacific Islander': 'PI',
                     'Afro-caribéen': 'AC',
                     'Afro-Américain': 'AA',
                     'Asian': 'AS',
                     'Latinoa': 'L', 'Latino(a)': 'L',
                     'Blanc [pas Latino(a)]': 'W'
                     }

        ethnicities = []

        for eth in ethnicity_list:
            if eth in temp_list.keys():
                ethnicities.append(temp_list[eth])
            else:
                print("\nPossible code values: " + str(temp_list.values()))
                code = input("Not recognized French ethnicity: {} please enter the code manually: ".format(eth))
                ethnicities.append(code)

                with open('Adjustments.txt', 'a') as f:
                    adjustments = "Adjustment for FRENCH Ethnicity: "
                    st = "{}:'{}'\n".format(eth, code)
                    f.writelines(adjustments)
                    f.writelines(st)
        return ethnicities

    def french_learning_barriers_scrub(self, learning_barriers_list):
        barriers = learning_barriers_list.strip().split(",")
        # print(barriers)
        barriers = [x.strip() for x in barriers]
        items = []
        d = {
            'Sans abri ou vivant dans un refuge': 'HOME',
            'Solía hacerse cargo del hogar o de sus hijos, pero ahora debe encontrar un trabajo.': 'HM',  # TODO
            'désactivé': 'D',
            'Personne à faibles revenus': 'LI',
            'Cela ne fonctionne que pendant quelques saisons.': 'MIG',
            'Vous avez un trouble d\'apprentissage.': 'LD',
            'Vous vous êtes enfui de chez vous lorsque vous étiez enfant ou adolescent.': 'RA',
            'L\'anglais n\'est PAS votre langue maternelle.': 'ESL',
            'Vous avez passé du temps en prison.': 'EO',
            'Vous étiez en famille d\'accueil.': 'FC',
            'Le système éducatif de votre pays était très différent': 'CB',
            'ou vous n\'avez jamais étudié dans votre pays.': 'CB',
            'Vous êtes au chômage depuis de nombreuses années.': 'UE',
            'but now you must find a job.': 'TANF',  # TODO
            'Votre TANF (Assistance temporaire pour les familles nécessiteuses) prendra fin dans les deux '
            'prochaines années.': 'TANF',
            'Parent célibataire.': 'SP',
            'Parent célibataire': 'SP',
            'Solía hacerse cargo del hogar o de sus hijos': 'HM', 'pero ahora debe encontrar un trabajo.': 'HM',  # TODO

        }
        for barrier in barriers:
            if barrier in d.keys():
                items.append(d[barrier])
            else:
                print("Possible code values: " + str(d.values()))
                code = input(
                    "Not recognized French learning barrier: {} please enter the code manually: ".format(barrier))
                items.append(code)
                with open('Adjustments.txt', 'a+') as f:
                    adjustments = "Adjustment for French Learning Barriers: "
                    st = "{}:'{}'\n".format(barrier, code)
                    f.writelines(adjustments)
                    f.writelines(st)
        return items

    def spanish_learning_barriers_scrub(self, learning_barriers_list):
        barriers = learning_barriers_list.strip().split(",")
        # print(barriers)
        barriers = [x.strip() for x in barriers]
        items = []
        d = {
            'Sin hogar o viviendo en un refugio.': 'HOME',
            'Solía hacerse cargo del hogar o de sus hijos, pero ahora debe encontrar un trabajo.': 'HM',
            'Posee alguna discapacidad.': 'D',
            'Bajos ingresos.': 'LI',
            'Sólo trabaja durante algunas temporadas.': 'MIG',
            'Posee alguna discapacidad de aprendizaje.': 'LD',
            'Dejó su hogar cuando era niño(a) o adolescente.': 'RA',
            'El inglés no es su idioma nativo.': 'ESL',
            'Pasaste tiempo en prisión.': 'EO',
            'Solías estar en un hogar de acogida.': 'FC',
            'El sistema educacional en su país es muy diferente o nunca estudió en su país.': 'CB',
            'or you never studied in your country.': 'CB',
            'Ha estado desempleado(a) por varios años.': 'UE',
            'but now you must find a job.': 'TANF',  # TODO
            'Su TANF (Asistencia Temporal para Familias Necesitadas terminará dentro de los próximos dos años.': 'TANF',
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
                with open('Adjustments.txt', 'a+') as f:
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
                     'African': 'AF',
                     'Latinoa': 'L',
                     'White [not Latinoa]': 'W'
                     }
        ethnicities = []
        for k, v in temp_list.items():
            if ethnicity_list.__contains__(k):
                ethnicities.append(temp_list[k])
        return ethnicities

    def english_learning_barriers_scrub(self, learning_barriers_list):
        if learning_barriers_list is None:
            return []
        barriers = learning_barriers_list.strip().split(",")
        # print(barriers)
        items = []

        barriers = [x.strip() for x in barriers]

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

    def gender_scrub(self, gender):
        if gender in GENDERMALE:
            return 'MALE'
        elif gender in GENDERFEMALE:
            return 'FEMALE'
        else:
            return 'NONBINARY'

    # def employment_scrub(self, language, personal_item):
    #
    #     item = "{}_EMPLOYMENT".format(language.upper())
    #     if item not in self._translations:
    #         print("{} was not found in the translation file, exiting program.")
    #         sys.exit(0)
    #     language_specific_dict = self._translations[item]
    #     return language_specific_dict[personal_item]

    def organize_form_responses(self, start_row, max_row, response_type):
        for row in self._current_worksheet.iter_rows(min_row=start_row, max_row=max_row):

            person = p.PersonObject()
            start_date = row[0].value
            person.set_program_startdate(start_date)
            person.update_name(row[1].value, row[2].value)
            person.update_dates(row[3].value, row[4].value)
            # print(row[5].value)
            # print(row[6].value)
            # print(r.value for r in row[3:5])
            try:
                c = [x.value.title() for x in row[5:7]]

                c.append(str(int(row[8].value)))

            except Exception as e:
                print("[Error]: {}".format(e.args))
                name = person.get_fullname()
                print("[Error]: Please double check ~{}'s~ address.".format(" ".join(name)))
                inp = input("Press Enter to acknowledge... the program will continue to run on everything else.")
                continue

            person.update_entire_address(c)
            mob = self.clean_phone_numbers(row[9].value)
            home = self.clean_phone_numbers(row[10].value)
            emer = self.clean_phone_numbers(row[13].value)
            person.update_phone_numbers([mob, home, emer])
            person.update_email(row[11].value)
            person.update_emergency_contact(row[12].value)

            person.update_gender(self.gender_scrub(row[14].value.upper()))
            person.update_latino(row[15].value)

            response_type = response_type.upper()
            if response_type == "ENGLISH":
                person.update_learning_barriers(self.english_learning_barriers_scrub(row[31].value))
                person.update_ethnicity(self.english_ethnicity_scrub(row[16].value))
                person.update_employment(self.english_employment_scrub(row[17].value))
            elif response_type == "SPANISH":
                person.update_learning_barriers(self.spanish_learning_barriers_scrub(row[31].value))
                person.update_ethnicity(self.spanish_ethnicity_scrub(row[16].value))  # TODO ELB
                person.update_employment(self.spanish_employment_scrub(row[17].value))
            elif response_type == "FRENCH":
                person.update_learning_barriers(self.french_learning_barriers_scrub(row[31].value))
                person.update_ethnicity(self.french_ethnicity_scrub(row[16].value))
                person.update_employment(self.french_employment_scrub(row[17].value))
            else:
                pass
                # TODO - another language, fix with directions
            c = [x.value for x in row[19:22]]
            # print(c)
            person.update_us_studies(row[18].value, c)
            person.update_oconus_studies(row[22].value, row[23].value, row[24].value)
            person.update_dependents(row[25].value, row[26].value, row[27].value, row[28].value, row[29].value,
                                     row[30].value)
            # print(row[31].value)
            # TODO ELB
            self._responses.add_person(person)

    def clean_phone_numbers(self, phone):
        # print(phone)
        # t = type(phone)
        phone_nums = []
        try:
            phone = regex.sub("\D", "", str(phone))
            if phone == '' or len(phone) == 1 or ord(phone[0]) > 57:
                # print(phone)
                return [None]

            # print(phone)
            phone_nums.append(phone[0:3])
            phone_nums.append(phone[3:6])
            phone_nums.append(phone[6:])

            return phone_nums

        except:
            #  print(phone, "\there we are")

            # traceback.print_stack()
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

        Annots[3].update(pdf.PdfDict(V=info, MaxLen=40))  # bday
        Annots[0].update(pdf.PdfDict(V=person.get_fullname()[0]))  # fname
        Annots[2].update(pdf.PdfDict(V=person.get_fullname()[1]))  # lname
        address = person.get_address()
        # print(address)
        Annots[4].update(pdf.PdfDict(V="  " + "  ".join(person.get_program_startdate()), MaxLen=40))
        Annots[5].update(pdf.PdfDict(V=address[0]))  # add
        Annots[6].update(pdf.PdfDict(V=self.clean_city(address[1])))  # city
        Annots[7].update(pdf.PdfDict(V=' N  Y', MaxLen=8))  # state
        info = " " + "  ".join(address[2])
        info = info[:10] + " " + info[10:]
        Annots[8].update(pdf.PdfDict(V=info, MaxLen=30))  # zipcode
        Annots[15].update(pdf.PdfDict(V=person.get_email()))
        phones = person.get_phone_numbers()
        # writing mobile

        # print(phones)
        if phones[0] is not None and phones[0][0] is not None:
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

        tto = []

        if person.get_latinoa():
            tto.append(30)
        else:
            tto.append(31)
        ethnicities = person.get_ethnicities()
        # print(ethnicities)
        for x in ethnicities:
            Annots[ANNOTATIONKEYS[x]].update(pdf.PdfDict(AS=pdf.PdfName("On")))

        work_declaration = person.get_working_declaration()
        if work_declaration == 'FT':
            tto.append(23)
        elif work_declaration == 'PT':
            tto.append(24)
        elif work_declaration == 'NT':
            tto.append(25)
        elif work_declaration == 'UE':
            tto.append(27)
        elif work_declaration == 'UA':
            tto.append(28)

        if person.get_studied_in_us():
            CONUS = True
            #
            Annots[43].update(pdf.PdfDict(V=str(person.get_highest_us_grade())))
            ny_grade = person.get_highest_ny_grade()
            try:
                ny_grade = int(regex.sub("\D", "", str(ny_grade)))
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
                tto.append(46)
            else:
                tto.append(47)
            if country_uni:
                tto.append(51)
            elif country_hse:
                tto.append(50)

        country_years = person.get_country_years()
        if country_years is not None:
            Annots[52].update(pdf.PdfDict(V=str(country_years)))

        hasDependents = person.get_dependent_status()

        if not hasDependents:
            tto.append(96)
            tto.append(98)
        else:
            tto.append(95)
            if person.single_parent():
                tto.append(97)
            else:
                tto.append(98)
        dependents = person.get_dependents()
        for i in range(len(dependents)):
            if dependents[i] is not None:
                Annots[53 + i].update(pdf.PdfDict(V=str(dependents[i])))

        learning_barriers = person.get_learning_barriers()
        tto_y = []

        if learning_barriers.__contains__('HOME'):
            tto_y.append(78)
        else:
            tto_y.append(57)
        if learning_barriers.__contains__('HM'):
            tto_y.append(80)
        else:
            tto_y.append(59)
        if learning_barriers.__contains__('D'):
            tto_y.append(81)
        else:
            tto_y.append(60)

        if learning_barriers.__contains__('LI'):
            tto_y.append(82)
        else:
            tto_y.append(61)
        if learning_barriers.__contains__('MIG'):
            tto_y.append(83)
        else:
            tto_y.append(62)
        if learning_barriers.__contains__('LD'):
            tto_y.append(90)
        else:
            tto_y.append(63)

        if learning_barriers.__contains__('RA'):
            tto_y.append(91)
        else:
            tto_y.append(64)

        if learning_barriers.__contains__('ESL'):
            tto_y.append(89)
        else:
            tto_y.append(88)

        if learning_barriers.__contains__('EO'):
            tto_y.append(93)
        else:
            tto_y.append(66)

        if learning_barriers.__contains__('FC'):
            tto_y.append(94)
        else:
            tto_y.append(67)

        if learning_barriers.__contains__('CB'):
            tto_y.append(84)
        else:
            tto_y.append(68)

        if learning_barriers.__contains__('UE'):
            tto_y.append(85)
        else:
            tto_y.append(69)

        if learning_barriers.__contains__('TANF'):
            tto_y.append(86)
        else:
            tto_y.append(70)

        if learning_barriers.__contains__('SP') and person.single_parent():
            # print('is a single parent' + str(person.get_fullname()))
            tto_y.append(87)
        else:
            tto_y.append(71)

        tto_y.extend([58, 65])

        # Annots[72].update(pdf.PdfDict(V='Electronically Completed by CI worker: ' + 'LS'))
        name = person.get_fullname()
        old_path = os.getcwd()
        path = os.path.join(os.getcwd(), folder_name)
        os.chdir(path)

        # print(os.getcwd())
        location = name[1] + ", " + name[0] + ' ISRF.pdf'
        for i in tto:
            Annots[i].update(pdf.PdfDict(AS=pdf.PdfName("On")))
        for i in tto_y:
            Annots[i].update(pdf.PdfDict(AS=pdf.PdfName("On"), V='On'))

        try:
            # person.Root.AcroForm.update(pdf.PdfDict(NeedAppearances=pdf.PdfObject('true')))
            persons_pdf.Root.AcroForm.update(pdf.PdfDict(NeedAppearances=pdf.PdfObject('true')))
        except AttributeError:
            print(str(name) + "  did not have the Root attribute")
        global NUM_WRITTEN

        print(str(NUM_WRITTEN) + " " + location)
        NUM_WRITTEN += 1
        pdf.PdfWriter().write(location, persons_pdf)
        os.chdir(old_path)
