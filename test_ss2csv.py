#!/usr/bin/env python

#
# Tests
#

import imp
import sys
import pytest
from ss2csv import *

test_file = 'test_ss.xlsx'


def test_xlrd_installation():
    """
    Make sure the xlrd module is installed and can be loaded.
    """
    try:
        imp.find_module('xlrd')
        assert True
    except ImportError:
        sys.exit(1)


def test_is_spreadsheet():
    """
    Test that the given test file is actually a spreadsheet
    """
    assert is_spreadsheet(test_file)


def test_workboot_data():
    """
    Test to verify a workbook obj can be created from the given
    test file
    """
    wb = WorkbookData(test_file)
    assert isinstance(wb, WorkbookData)

    # Verify the given workbook has at least 1 worksheet.  If not,
    # something went really wrong
    assert wb.worksheet_count >= 1


def test_save_to_csv():
    """
    Test to validate a csv file can be saved from the xlsx.
    """
    wb = WorkbookData(test_file)
    s = SaveToCSV(wb)

    # Make sure that the built instance is what it should be
    assert isinstance(s, threading.Thread)
    assert isinstance(s, SaveToCSV)
