#!/usr/bin/env python

#
# Tests
#

import sys
import pytest
from ss2csv import *

try:
    import xlrd
except ImportError, xlrd_error:
    print xlrd_error.message
    sys.exit(1)

test_file = 'test_ss.xlsx'


def test_is_spreadsheet():
    assert is_spreadsheet(test_file)


def test_workboot_data():
    wb = WorkbookData(test_file)
    assert isinstance(wb, WorkbookData)
    assert wb.worksheet_count >= 1


def test_save_to_csv():
    wb = WorkbookData(test_file)
    s = SaveToCSV(wb)
    assert isinstance(s, threading.Thread)
    assert isinstance(s, SaveToCSV)
