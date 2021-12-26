"""
test cases for Faculty Calendar
"""
from faculty_calendar import lengthofsession, check
import ifls
from faculty_class import FacultySlots


# test functions
def test_lengthofsession():
    """
    test function of lengthofsession()
    """
    temp1 = lengthofsession("ChrisMcCandless")
    temp2 = lengthofsession("Naruto")
    temp3 = lengthofsession("Dhoni")
    temp4 = lengthofsession(None)
    assert temp1 == 15
    assert temp2 == 6
    assert temp3 == 5
    assert temp4 == 0


def test_check():
    """
    test function for check()
    """
    temp1 = check("Eminem")
    temp2 = check("Itachi")
    temp3 = check("Sukuna")
    temp4 = check(None)
    assert temp1 == "Eminem"
    assert temp2 == "Itachi"
    assert temp3 == "Sukuna"
    assert temp4 == "empty"


"""Test Code for Integrated Faculty Load Sheet"""


def test_update_faculty():
    """Test function for update_faculty function"""
    faculty_1 = FacultySlots('XYZ')
    faculty_list = []
    faculty_list.append(faculty_1)

    temp_1 = ifls.update_faculty('XYZ', 'M', faculty_list)
    temp_2 = ifls.update_faculty('XYZ', 'M&A', faculty_list)
    temp_3 = ifls.update_faculty('XYZ', None, faculty_list)
    temp_4 = ifls.update_faculty('XYZ', 'A', faculty_list)
    temp_5 = ifls.update_faculty(None, None, faculty_list)

    assert temp_1 == 'Morning'
    assert temp_2 == 'Morning and Afternoon'
    assert temp_3 == 'No Session'
    assert temp_4 == 'Afternoon'
    assert temp_5 == 'No Name'


def test_free_slots():
    """Test code for free_slots function"""
    temp1 = ifls.free_slots(6)
    assert temp1 == 12
