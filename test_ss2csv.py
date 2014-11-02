#!/usr/bin/env python

import sys
from ss2csv import *

#
# Tests
#

try:
    import xlrd
except ImportError, xlrd_error:
    print xlrd_error.message
    sys.exit(1)

def test_is_spreadsheet(file):
    assert is_spreadsheet(file)
