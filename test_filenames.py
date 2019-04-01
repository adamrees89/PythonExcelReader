import unittest
import templateReader
import os
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

class templateReaderFunctionTest(unittest.TestCase):
    """ Class doc """

    def test_Excelfile(self):
        """ Class initialiser """
        f = 'Sample Excel Files/testExcel.xlsx'
        try:
            os.remove("template.db")
        except OSError:
            pass

        try:
            wb = openpyxl.load_workbook(filename=f)
            for ws in wb.worksheets:
                templateReader.templateSheet(ws)
        except Exception:
            self.fail("templateSheet() threw an exception".format(f))

    def test_Wordfile(self):
        """ Class initialiser """
        f = 'Sample Excel Files/sample.docx'
        try:
            wb = openpyxl.load_workbook(filename=f)
            for ws in wb.worksheets:
                templateReader.templateSheet(ws)
        except openpyxl.utils.exceptions.InvalidFileException:
            pass
        except Exception as e:
            self.fail('Unexpected exception raised:', e)
        else:
            self.fail('Exception not raised')


def main():
    unittest.main()


if __name__ == '__main__':
    main()
