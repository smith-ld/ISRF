import datetime

# multiple language support
YESANSWERS = ['Yes', 'Sí']


class PersonObject:

    def __init__(self):

        self._name = []
        self._birthdate = None
        self._todaysdate = None
        self._address = []
        self._phones = []
        self._email = None
        self._em_contact = None
        self._gender = None
        self._latino = False
        self._ethnicity_list = []
        self._working_declaration = None
        self._studied_in_US = False
        self._highest_us_grade = None
        self._highest_ny_grade = None
        self._last_ny_schoolname = None
        self._finished_hs = False
        self._finished_uni = False
        self._country_years = None
        self._dependents = []
        self._single_parent = False
        self._has_dependents = False
        self._learning_barriers = []

    def get_birthdate(self):
        return self._birthdate

    def get_address(self):
        return self._address

    def get_fullname(self):
        return self._name

    def update_name(self, fname, lname) :
        self._name.extend([fname, lname])

    def update_dates(self, birthdate, todaysdate):
        self._birthdate = birthdate
        self._todaysdate = todaysdate

    def update_entire_address(self, address):
        #print(self._address)
        self._address = address
        # print(self._address)

    def update_phone_numbers(self, phones):
        self._phones = phones
        #print(self._phones)

    def get_phone_numbers(self):
        return self._phones

    def update_email(self, email):
        self._email = email

    def get_email(self):
        return self._email

    def update_emergency_contact(self, em_con):
        if em_con is not None:
            self._em_contact = em_con

    def update_gender(self, gender):
        self._gender = gender

    def update_latino(self, declaration):
        if declaration in YESANSWERS:
            self._latino = True
        else:
            self._latino = False

    def update_ethnicity(self, ethnicity_list):
        self._ethnicity_list = ethnicity_list

    def update_employment(self, employment):
        self._working_declaration = employment

    def update_us_studies(self, studies, us_study_list):
        if studies in YESANSWERS:
            self._studied_in_US = True
            self._highest_us_grade = us_study_list[0]
            self._highest_ny_grade = us_study_list[1]
            if us_study_list[-1] != 'None':
                self._last_ny_schoolname = us_study_list[-1]

    def update_oconus_studies(self, finish_hs_oconus, finish_uni_oconus, years):
        if finish_hs_oconus in YESANSWERS:
            self._finished_hs = True
        else:
            self.update_oconus_study_years(years)
        if finish_uni_oconus in YESANSWERS:
            self._finished_uni = True

    def update_oconus_study_years(self, years_studied):
        self._country_years = years_studied

    def update_dependents(self, has_deps, single_par, zero_to_four ,five_to_ten, ele_to_thir, four_to_eigh):
        if has_deps in YESANSWERS:
            self._has_dependents = True
            self._dependents = [zero_to_four, five_to_ten, ele_to_thir, four_to_eigh]
        if single_par in YESANSWERS:
            self._single_parent = True

    def update_learning_barriers(self, learning_barriers):
        self._learning_barriers = learning_barriers

    def get_learning_barriers(self):
        return self._learning_barriers

    def __str__(self):
        return " ".join(self._name[0:2])

    def get_em_contact(self):
        return self._em_contact

    def get_gender(self):
        return self._gender

    def get_latinoa(self):
        return self._latino

    def get_ethnicities(self):
        return self._ethnicity_list

    def get_working_declaration(self):
        return self._working_declaration

    def get_studied_in_us(self):
        return self._studied_in_US

    def get_highest_us_grade(self):
        return self._highest_us_grade

    def get_highest_ny_grade(self):
        return self._highest_ny_grade

    def get_ny_school(self):
        return self._last_ny_schoolname

    def get_country_years(self):
        return self._country_years

    def get_dependent_status(self):
        return self._has_dependents

    def get_finished_hs(self):
        return self._finished_hs

    def get_finished_uni(self):
        return self._finished_uni

    def single_parent(self):
        return self._single_parent

    def get_dependents(self):
        # print(self._dependents)
        return self._dependents

"""    self._studied_in_US = False
        self._highest_us_grade = None
        self._highest_ny_grade = None
        self._last_ny_schoolname = None
        self._finished_hs = False
        self._finished_uni = False
        self._country_years = None
        self._dependents = []
        self._single_parent = False
        self._has_dependents = False
        self._learning_barriers = []"""


class SingletonPersons:
    def __init__(self):
        self._persons = []

    def get_person_list(self):
        return self._persons

    def add_person(self, person):
        self._persons.append(person)
