"""Function to store class required"""


class InputSchedule:
    """Class for extracting data from input"""

    def __init__(self, lead_1, lead_2, lead_3, session_slot):
        """Constructor"""
        self.lead_1 = lead_1
        self.lead_2 = lead_2
        self.lead_3 = lead_3
        self.session_slot = session_slot

    def get_lead_1(self):
        """Return first lead/faculty"""
        return self.lead_1

    def get_lead_2(self):
        """Return second lead/faculty"""
        return self.lead_2

    def get_lead_3(self):
        """Return third lead/faculty"""
        return self.lead_3

    def get_session_slot(self):
        """Return session slot"""
        return self.session_slot


class FacultySlots:
    """Class for representing faculty"""

    def __init__(self, fac_name, m_slots=0, a_slots=0):
        """Constructor for class"""
        self.fac_name = fac_name
        self.m_slots = m_slots
        self.a_slots = a_slots

    def get_faculty_name(self):
        """Return Faculty Name"""
        return self.fac_name

    def get_m_slots(self):
        """Return number of morning slots"""
        return self.m_slots

    def get_a_slots(self):
        """return number of afternoon slots"""
        return self.a_slots

    def update_m_slots(self):
        """Increment number of morning slots by 1"""
        self.m_slots += 1

    def update_a_slots(self):
        """Increment number of afternoon slots by 1"""
        self.a_slots += 1

    def prt(self):
        """Print all the information"""
        print(self.fac_name, self.m_slots, self.a_slots)
