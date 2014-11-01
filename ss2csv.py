#!/usr/bin/env python
# coding=utf-8

#
# Base functionality to convert XLS(x) files to CSV.
# Options to consider:
#   1. -f <filename>
#   2. -d <directory with files>
#       A. Implement a file search for
#

import sys
import csv
import threading
import mimetypes
import glob
import os

# If xlrd isn't installed, stop.
try:
    import xlrd
except ImportError, xlrd_error:
    print xlrd_error.message
    sys.exit(1)


def find_excel_files_by_mimetype():
    """
   Return a list of Excel files by matching MIMETYPE.
   :return: list
   """
    return [f for f in glob.glob("*.*")
        if mimetypes.guess_type(f)[0] == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']


def find_excel_files_by_extension(extension, directory):
    """
   Return a list of Excel files by extension.
   :param extension: str
   :param directory: str
   :return: list
   """
    os.chdir(directory)
    return [f for f in glob.glob("*.{0}".format(extension))]


class WorkbookData(object):
    """
   Primary class to handle the provided workbook.
   """
    def __init__(self, spreadsheet):
        self._ss = xlrd.open_workbook(spreadsheet)
        self.sheets = self._ss.sheets()
        self.worksheet_count = self._ss.nsheets

        if self.worksheet_count == 1:
            self.worksheet_stats = {'columns': self._ss.sheets()[0].ncols,
                                    'rows': self._ss.sheets()[0].nrows}
        else:
            self.worksheet_stats = [{'name': sheet.name,
                                     'columns': sheet.ncols,
                                     'rows': sheet.nrows} for sheet in self.sheets]


class SaveToCSV(threading.Thread):
    """
   Save each sheet in the workbook to it's own CSV file.  Each
   worksheet is saved to CSV by it's own thread.
   """
    def __init__(self, workbook):
        threading.Thread.__init__(self)
        self.workbook = workbook

    def run(self):
        """
       Main worker.
       :return: None
       """
        for sheet in self.workbook.sheets:
            with open("{0}.csv".format(sheet.name), 'wb') as csv_file:
                cw = csv.writer(csv_file, quoting=csv.QUOTE_ALL)
                [cw.writerow(sheet.row_values(row_num))
                 for row_num in xrange(len(self.workbook.worksheet_stats))]

if __name__ == '__main__':
    wb = WorkbookData('test_file_1.xlsx')
    threads = []
    for num in xrange(wb.worksheet_count):
        thread = SaveToCSV(wb)
        thread.start()
        threads.append(thread)

    [t.join() for t in threads]

    if wb.worksheet_count == 1:
        print """Converted workbooks : {0}
             Rows      : {1}
             Colums    : {2}""".format(wb.worksheet_count,
                                        wb.worksheet_stats['rows'],
                                        wb.worksheet_stats['columns'])
    else:
        print "Converted workbooks : {0}".format(wb.worksheet_count)
        for sheet_data in wb.worksheet_stats:
            print "          Sheet     : {0}".format(sheet_data['name'])
            print """          Rows      : {0}
         Columns   : {2}\n""".format(wb.worksheet_count,
                                      sheet_data['rows'],
                                      sheet_data['columns'])
