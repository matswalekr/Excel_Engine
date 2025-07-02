import sys
import os

# Add the project root directory to the Python path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from excel_code import Excel, Excel_read, Excel_modify ,Excel_write, open_excel
from openpyxl import load_workbook
import unittest
from typing import List
import warnings
import pandas as pd

class Test_Excel_Cell(unittest.TestCase):

    def setUp(self)->None:
        self.test_file = r"tests/test_workbook.xlsx"
        self.excel = load_workbook(filename = self.test_file)

    def test_cell_basic(self)->None:

        cell = Excel.Cell(
            name = "Hello",
            value_ = 5,
            col = "A",
            row = 3,
            sheet_name = "Sheet1",
            workbook=self.excel
        )
        # test basic cell operations
        self.assertEqual(cell.pos, "A3")
        self.assertEqual(cell.value_, 5)

    def test_comparison(self)->None:
        cell1 = Excel.Cell(
            name = "Hello",
            value_ = 5,
            col = "A",
            row = 3,
            sheet_name = "Sheet1",
            workbook=self.excel
        ) 

        cell_same = Excel.Cell(
            name = "New Name",
            value_ = 7, #This tests if the name is the same, but the value_s are different, the result should still be the same
            col = "A",
            row = 3,
            sheet_name = "Sheet1",
            workbook=self.excel
        )

        self.assertTrue(cell_same == cell1)

        cell_different_sheet = Excel.Cell(
            name = "New Name",
            value_ = 7, #This tests if the name is the same, but the value_s are different, the result should still be the same
            col = "A",
            row = 3,
            sheet_name = "Sheet2",
            workbook=self.excel
        )
        self.assertTrue(cell_different_sheet != cell1)

        cell_different_row = cell_different_sheet
        cell_different_row.sheet_name = "Sheet1"
        cell_different_row.row = 4
        self.assertTrue(cell_different_row != cell1)
        
        cell_different_column = cell_different_row
        cell_different_column.row = 3
        cell_different_column.col = "B"
        self.assertTrue(cell_different_column != cell1)

    def test_shifts(self)->None:
        cell = Excel.Cell(
            name = "Hello",
            value_ = 5,
            col = "A",
            row = 5,
            sheet_name = "Sheet1",
            workbook=self.excel
        ) 

        cell.shift_column(shift = 2)#Shift the column by 2
        self.assertEqual(cell.pos, "C5")

        cell.shift_column(shift = -1) #shift the column back by 1
        self.assertEqual(cell.pos, "B5")

        cell.shift_row(shift = 2) #shift the row by 2
        self.assertEqual(cell.pos, "B7")

        cell.shift_row(shift = -1) #shift the row back by 1
        self.assertEqual(cell.pos, "B6")

    def test_operations(self)->None:
        cell = Excel.Cell(
            name = "Hello",
            value_ = 5,
            col = "A",
            row = 5,
            sheet_name = "Sheet1",
            workbook=self.excel
        )

        cell *= 2 # 5*2 = 10
        self.assertEqual(cell.value_, 10)

        cell /= 2 #10/2 = 5
        self.assertEqual(cell.value_, 5)

        cell += 3 #5 + 3 = 8
        self.assertEqual(cell.value_, 8)

        cell -= 2 # 8-2= 6
        self.assertEqual(cell.value_, 6)

    def test_out_of_bounds_error(self)->None:
        with self.assertRaises(ValueError):
            cell = Excel.Cell(
                name = "Hello",
                value_ = 5,
                col = "A",
                row = 3,
                sheet_name = "Sheet1",
                workbook=self.excel
            )
            cell.shift_column(-1)

        with self.assertRaises(ValueError):
            cell = Excel.Cell(
                name = "Hello",
                value_ = 5,
                col = "A",
                row = 3,
                sheet_name = "Sheet1",
                workbook=self.excel
            )
            cell.shift_row(-8)

    def test_other_errors(self)->None:
        with self.assertRaises(TypeError):
            cell = Excel.Cell(
                name = "Hello",
                value_ = 5,
                col = "A",
                row = 3,
                sheet_name = "Sheet1",
                workbook=self.excel
            )
            cell.shift_column("B")      

    def test_formula_evaluation(self)->None:
        cell = Excel.Cell(
                name = "Hello",
                value_ = "=2*2", #Note that this has not been loaded, but has manually been inputted into the document
                col = "A",
                row = 3,
                sheet_name = "Test_cell",
                filename = r"tests/test_workbook.xlsx",
                workbook=self.excel
        )

        self.assertEqual(cell.value, 4)
        
        cell2 = Excel.Cell(
                name = "Hello",
                value_ = "=SUM(C3:C4)", #Note that this has not been loaded, but has manually been inputted into the document
                col = "C",
                row = 5,
                sheet_name = "Test_cell",
                workbook=self.excel,
                filename = r"tests/test_workbook.xlsx"
        )
        # Note that C3 and C4 ahve been initialised as 2 and 3

        self.assertEqual(cell2.value, 5)


class Test_Excel_Base_Class(unittest.TestCase):

    def setUp(self) -> None:
        self.test_file = r"tests/test_workbook.xlsx"
        self.excel = Excel(path = self.test_file, open_path=True)
        # Suppress DeprecationWarnings
        warnings.filterwarnings("ignore", category=UserWarning)

    def test_open(self)->None:
        with Excel(path = self.test_file) as doc:
            self.assertIsInstance(doc, Excel)

    def test_sheet_names(self)->None:
        sheet_names:List[str] = self.excel.sheets

        self.assertIn("Test_Name_Found", sheet_names)
        self.assertIn("Test_cell", sheet_names)
        
    def test_defined_names(self)->None:
        defined_names = self.excel.defined_names
        self.assertIn("Test_defined_name", defined_names.keys())


class Test_Excel_Read(unittest.TestCase):

    def setUp(self) -> None:
        self.test_file = r"tests/test_workbook_read.xlsx"
        self.excel = Excel_read(path = self.test_file, open_path=True)
        # Suppress DeprecationWarnings
        warnings.filterwarnings("ignore", category=UserWarning)

    def test_get_items(self):
        cellA2= self.excel["A2"] #Initialised to be hello (manually)

        self.assertEqual(cellA2.value, "Hello")
        
        cellA3 = self.excel["A3"] #Initialised to be 13 (manually)

        self.assertEqual(cellA3.value, 13)

    def test_contains(self):
        self.assertTrue("Sheet2" in self.excel)

        self.assertFalse("RandomName" in self.excel)

        cell_contained = Excel.Cell(col = "A",
                                   row = 6,
                                   workbook = self.excel,
                                   sheet_name = "Sheet1") #Value also set manually
        
        self.assertTrue(cell_contained in self.excel)


class Test_Excel_Modify(unittest.TestCase):
    def setUp(self) -> None:
        self.test_file = r"tests/test_workbook_modify.xlsx"
        self.excel = Excel_modify(path = self.test_file)
        # Suppress DeprecationWarnings
        warnings.filterwarnings("ignore", category=UserWarning)

    def test_sheet_handling(self)->None:
        sheet_name = "Sheet2"

        if sheet_name in self.excel.sheets:
            self.excel.remove_sheet(sheet_name)
        
        # Assert the cell doesnt exist yet
        self.assertFalse(sheet_name in self.excel.sheets)

        #Add the sheet and assert it exists
        self.excel.new_sheet(sheet_name)
        self.assertTrue(sheet_name in self.excel.sheets)

        #Remove the sheet again and assert it doesnt exist
        self.excel.remove_sheet(sheet_name)
        self.assertFalse(sheet_name in self.excel.sheets)

    def test_assign_df(self)->None:
        data = {
            'A': [1, 2, 3],
            'B': [4, 5, 6],
            'C': [7, 8, 9]
        }
        df = pd.DataFrame(data)

        sheet_name = "Df_test_sheet"

        self.excel.set_cells_pandas(start_cell="A1", df = df, sheet_name=sheet_name)

        worksheet = self.excel[sheet_name]

        self.assertTrue(worksheet["A1"].value == 1)
        self.assertTrue(worksheet["A2"].value == 2)
        self.assertTrue(worksheet["A3"].value == 3)

        self.assertTrue(worksheet["B1"].value == 4)
        self.assertTrue(worksheet["B2"].value == 5)
        self.assertTrue(worksheet["B3"].value == 6)

        self.assertTrue(worksheet["C1"].value == 7)
        self.assertTrue(worksheet["C2"].value == 8)
        self.assertTrue(worksheet["C3"].value == 9)

    def test_delete_content_sheet(self)->None:
        self.excel.new_sheet("NewSheet")

        self.excel["NewSheet"]["A2"].value = 3

        self.assertTrue(self.excel["NewSheet"]["A2"].value == 3)

        self.excel.clear_sheet("NewSheet")

        self.assertTrue(self.excel["NewSheet"]["A2"].value is None)

        self.excel.remove_sheet("NewSheet")

    def test_set_range(self)->None:
        sheet_name:str = "Range_test_sheet"

        if sheet_name not in self.excel.sheets:
            self.excel.new_sheet(sheet_name)

        #Set the active sheet to the test sheet for this test
        self.excel.workbook.active = self.excel[sheet_name]

        self.excel["A2:B3"] = 4

        self.assertTrue(self.excel[sheet_name]["A2"].value == 4)
        self.assertTrue(self.excel[sheet_name]["A3"].value == 4)
        self.assertTrue(self.excel[sheet_name]["B2"].value == 4)
        self.assertTrue(self.excel[sheet_name]["B3"].value == 4)

    def test_delete_item(self)->None:

        self.excel["D32"] = 3
        self.assertTrue(self.excel["D32"].value == 3)

        del self.excel["D32"]
        self.assertTrue(self.excel["D32"].value == 0)

        self.excel["D32:D33"] = 5
        self.assertTrue(self.excel["D32"].value == 5)
        self.assertTrue(self.excel["D33"].value == 5)

        del self.excel["D32:D33"]
        self.assertTrue(self.excel["D33"].value == 0)
        self.assertTrue(self.excel["D33"].value == 0)


class Test_Excel_Write(unittest.TestCase):
    def setUp(self) -> None:
        self.test_file = r"tests/test_excel_write.xlsx"
        # Suppress DeprecationWarnings
        warnings.filterwarnings("ignore", category=UserWarning)


    def test_saving(self)->None:

        with Excel_write(path = self.test_file) as excel:
            excel["A2"] = 3
            self.assertTrue(excel["A2"].value == 3)
        # This should automatically save the excel sheet
            

        with Excel_read(path = self.test_file) as excel_read:
            self.assertTrue(excel_read["A2"].value == 3)

        #Enter the file again in read mode to set the value back to 0 for further testing
        with Excel_write(path = self.test_file) as excel:
            del excel["A2"]
            self.assertTrue(excel["A2"].value == 0)
        # This should automatically save the excel sheet
        


class Test_enter_excel(unittest.TestCase):

    def setUp(self) -> None:
        self.test_path = r"tests/test_workbook_modify.xlsx"

    def test_enter_read(self)->None:
        test_excel = open_excel(path = self.test_path, mode = "r")

        self.assertTrue(isinstance(test_excel, Excel_read))
        
    def test_enter_modify(self)->None:
        test_excel = open_excel(path = self.test_path, mode = "m")

        self.assertTrue(isinstance(test_excel, Excel_modify))

    def test_enter_write(self)->None:
        test_excel = open_excel(path = self.test_path, mode = "w")

        self.assertTrue(isinstance(test_excel, Excel_write))


# To run use:  PYTHONPATH=/Users/matswalker/DCF/Excel_Engine pytest tests/test_excel_engine.py