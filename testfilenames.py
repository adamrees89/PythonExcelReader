import unittest
import templateReader
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

class templateReaderFunctionTest(unittest.TestCase):
	""" Class doc """
	
	def test_Excelfile(self):
		""" Class initialiser """
		f='Sample Excel Files/testExcel.xlsx'
		try:
			wb = openpyxl.load_workbook(filename=f)
			for ws in wb.worksheets:
				templateReader.templateSheet(ws)
		except Exception:
			self.fail("templateReader.templateSheet() threw an exception with the sample .xlsx file:".format(f))

	def test_Wordfile(self):
		""" Class initialiser """
		f='Sample Excel Files/sample.docx'
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

