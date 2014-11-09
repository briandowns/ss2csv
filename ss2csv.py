#!/usr/bin/env python
# coding=utf-8

#
# Convert spreadsheets to CSV.  Each sheet in the workbook will be
# saved to its own csv file named with the sheet name.
#

#
# TODO: Create better output for conversion stats.
#

import os
import csv
import sys
import glob
import argparse
import threading
import mimetypes

# If xlrd isn't installed, tuck and roll.
try:
    import xlrd
except ImportError, xlrd_error:
    print xlrd_error.message
    sys.exit(1)


# Remnant from a previous iteration of the script.
def find_excel_files_by_mimetype(directory):
    """
    Return a list of Excel files by matching MIMETYPE.
    @return: list of found files
    """
    return [f for f in glob.glob("{0}/*.*".format(os.chdir(directory))) if is_spreadsheet(f)]


def is_spreadsheet(file):
    """
    Make sure the file passed in is a spreadsheet.
    @param file: str for file to test
    @return: bool
    """
    if 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' == \
            mimetypes.guess_type(file)[0]:
        return True
    return False


# Remnant from a previous iteration of the script.
def find_excel_files_by_extension(extension, directory):
    """
    Return a list of Excel files by extension.
    @param extension: str
    @param directory: str
    @return: list of found files
    """
    return [f for f in glob.glob("{0}/*.{1}".format(os.chdir(directory),
                                                    extension))]


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
                                     'rows': sheet.nrows}
                                    for sheet in self.sheets]


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
        """
        for sheet in self.workbook.sheets:
            with open("{0}.csv".format(sheet.name), 'wb') as csv_file:
                cw = csv.writer(csv_file, quoting=csv.QUOTE_ALL)
                [cw.writerow(sheet.row_values(row_num))
                 for row_num in xrange(len(self.workbook.worksheet_stats))]

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Save spreadsheet in CSV.')
    parser.add_argument('-c', '--convert',
                        required=True, help='Convert spreadsheet')
    args = vars(parser.parse_args())

    if args['convert']:
        if is_spreadsheet(args['convert']):
            wb = WorkbookData(args['convert'])
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
