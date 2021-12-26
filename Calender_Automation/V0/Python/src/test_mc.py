"""
test cases for Master Calendar
"""
from calendar_func import create_master_cal


def test_master_calendar():
    """
    Checking whether the function performs
    all operations and sets flag
    MANUAL TEST DUE TO CHANGING PATH
    """
    path = "PUT INPUT FILE PATH HERE"
    temp = create_master_cal(path)
    assert temp == 1
