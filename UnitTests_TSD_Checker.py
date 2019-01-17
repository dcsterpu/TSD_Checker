import unittest
import TSD_Checker
import xlrd
import openpyxl

class Test_02043_18_04939_STRUCT_0000(unittest.TestCase):


    def setUp(self):

        self.app = TSD_Checker.QApplication(TSD_Checker.sys.argv)
        self.TSD_checker = TSD_Checker.Test()

    def test_ok_xls(self):
        self.workbook = xlrd.open_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0000/doc4_ok_xls.xls", formatting_info=True)
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0000_XLS(self.workbook), 1)

    def test_ok_uppercase_xls(self):
        self.workbook = xlrd.open_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0000/doc4_ok_uppercase_xls.xls", formatting_info=True)
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0000_XLS(self.workbook), 1)

    def test_ok_french_xls(self):
        self.workbook = xlrd.open_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0000/doc4_ok_french_xls.xls", formatting_info=True)
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0000_XLS(self.workbook), 1)

    def test_notok_xls(self):
        self.workbook = xlrd.open_workbook(
            "C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0000/doc4_notok_xls.xls",
            formatting_info=True)
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0000_XLS(self.workbook), False)

    def test_notok_french_xls(self):
        self.workbook = xlrd.open_workbook(
            "C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0000/doc4_notok_french_xls.xls",
            formatting_info=True)
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0000_XLS(self.workbook), False)




    def test_ok_xlsx(self):
        self.workbook = openpyxl.load_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0000/doc4_ok_xlsx.xlsx")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0000_XLSX_XLSM(self.workbook), 1)

    def test_ok_uppercase_xlsx(self):
        self.workbook = openpyxl.load_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0000/doc4_ok_uppercase_xlsx.xlsx")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0000_XLSX_XLSM(self.workbook), 1)

    def test_ok_french_xlsx(self):
        self.workbook = openpyxl.load_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0000/doc4_ok_french_xlsx.xlsx")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0000_XLSX_XLSM(self.workbook), 1)

    def test_notok_xlsx(self):
        self.workbook = openpyxl.load_workbook(
            "C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0000/doc4_notok_xlsx.xlsx")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0000_XLSX_XLSM(self.workbook), False)

    def test_notok_french_xlsx(self):
        self.workbook = openpyxl.load_workbook(
            "C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0000/doc4_notok_french_xlsx.xlsx")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0000_XLSX_XLSM(self.workbook), False)



    def test_ok_xlsm(self):
        self.workbook = openpyxl.load_workbook(
            "C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0000/doc4_ok_xlsm.xlsm")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0000_XLSX_XLSM(self.workbook), 1)

    def test_ok_uppercase_xlsm(self):
        self.workbook = openpyxl.load_workbook(
            "C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0000/doc4_ok_uppercase_xlsm.xlsm")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0000_XLSX_XLSM(self.workbook), 1)

    def test_ok_french_xlsm(self):
        self.workbook = openpyxl.load_workbook(
            "C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0000/doc4_ok_french_xlsm.xlsm")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0000_XLSX_XLSM(self.workbook), 1)

    def test_notok_xlsm(self):
        self.workbook = openpyxl.load_workbook(
            "C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0000/doc4_notok_xlsm.xlsm")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0000_XLSX_XLSM(self.workbook), False)

    def test_notok_french_xlsm(self):
        self.workbook = openpyxl.load_workbook(
            "C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0000/doc4_notok_french_xlsm.xlsm")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0000_XLSX_XLSM(self.workbook), False)



class Test_02043_18_04939_STRUCT_0005(unittest.TestCase):
    def setUp(self):
        self.app = TSD_Checker.QApplication(TSD_Checker.sys.argv)
        self.TSD_checker = TSD_Checker.Test()

    def test_notok_formula_xls(self):

        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0005_XLS("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0005/doc4_notok_formula_xls.xls"), 0)

    def test_ok_notformula_xls(self):

        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0005_XLS("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0005/doc4_ok_notformula_xls.xls"), 1)

    def test_ok_empty_xls(self):
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0005_XLS("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0005/doc4_ok_empty_xls.xls"), 1)



    def test_notok_formula_xlsx(self):
        self.workbook = openpyxl.load_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0005/doc4_notok_formula_xlsx.xlsx")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0005_XLSX_XLSM(self.workbook), 1)
    def test_ok_notformula_xlsx(self):
        self.workbook = openpyxl.load_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0005/doc4_ok_empty_xlsx.xlsx")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0005_XLSX_XLSM(self.workbook), 1)

    def test_ok_empty_xlsx(self):
        self.workbook = openpyxl.load_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0005/doc4_ok_empty_xlsx.xlsx")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0005_XLSX_XLSM(self.workbook), 1)



    def test_notok_formula_xlsm(self):
        self.workbook = openpyxl.load_workbook( "C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0005/doc4_notok_formula_xlsm.xlsm")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0005_XLSX_XLSM(self.workbook), 1)

    def test_ok_notformula_xlsm(self):
        self.workbook = openpyxl.load_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0005/doc4_ok_empty_xlsm.xlsm")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0005_XLSX_XLSM(self.workbook), 1)

    def test_ok_empty_xlsm(self):
        self.workbook = openpyxl.load_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0005/doc4_ok_empty_xlsm.xlsm")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0005_XLSX_XLSM(self.workbook), 1)




class Test_02043_18_04939_STRUCT_0010(unittest.TestCase):
    def setUp(self):
        self.app = TSD_Checker.QApplication(TSD_Checker.sys.argv)
        self.TSD_checker = TSD_Checker.Test()



    def test_ok_IsReference_xls(self):
        self.workbook = xlrd.open_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0010/doc4_ok_IsReference_xls.xls", formatting_info=True)
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0010_XLS(self.workbook), 1)

    def test_notok_IsEmpty_xls(self):
        self.workbook = xlrd.open_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0010/doc4_notok_IsEmpty_xls.xls", formatting_info=True)
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0010_XLS(self.workbook), 0)


    def test_ok_IsReference_xlsx(self):
        self.workbook = openpyxl.load_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0010/doc4_ok_IsReference_xlsx.xlsx")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0010_XLSX_XLSM(self.workbook), 1)


    def test_notok_IsEmpty_xlsx(self):
        self.workbook = openpyxl.load_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0010/doc4_notok_IsEmpty_xlsx.xlsx")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0010_XLSX_XLSM(self.workbook), 0)


    def test_ok_IsReference_xlsm(self):
        self.workbook = openpyxl.load_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0010/doc4_ok_IsReference_xlsm.xlsm")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0010_XLSX_XLSM(self.workbook), 1)

    def test_notok_IsEmpty_xlsm(self):
        self.workbook = openpyxl.load_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0010/doc4_notok_IsEmpty_xlsm.xlsm")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0010_XLSX_XLSM(self.workbook), 0)





class Test_02043_18_04939_STRUCT_0011(unittest.TestCase):
    def setUp(self):
        self.app = TSD_Checker.QApplication(TSD_Checker.sys.argv)
        self.TSD_checker = TSD_Checker.Test()

    def test_ok_equal_xls(self):
        self.workbook = xlrd.open_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0011/doc4_ok_equal_xls.xls")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0011_XLS(self.workbook), 1)

    def test_notok_notequal_xls(self):
        self.workbook = xlrd.open_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0011/doc4_notok_notequal_xls.xls")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0011_XLS(self.workbook), 0)

    def test_ok_equal_xlsx(self):
        self.workbook = openpyxl.load_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0011/doc4_ok_equal_xlsx.xlsx")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0011_XLSX_XLSM(self.workbook), 1)

    def test_notok_notequal_xlsx(self):
        self.workbook = openpyxl.load_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0011/doc4_notok_notequal_xlsx.xlsx")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0011_XLSX_XLSM(self.workbook), 0)





class Test_02043_18_04939_STRUCT_0020(unittest.TestCase):
    def setUp(self):
        self.app = TSD_Checker.QApplication(TSD_Checker.sys.argv)
        self.TSD_checker = TSD_Checker.Test()

    def test_ok_xls(self):
        self.workbook = xlrd.open_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0020/doc4_ok_xls.xls")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0020_XLS(self.workbook), 1)

    def test_ok_upper_xls(self):
        self.workbook = xlrd.open_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0020/doc4_ok_upper_xls.xls")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0020_XLS(self.workbook), 1)

    def test_notok_xls(self):
        self.workbook = xlrd.open_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0020/doc4_notok_xls.xls")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0020_XLS(self.workbook), 0)

    def test_ok_xlsx(self):
        self.workbook = openpyxl.load_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0020/doc4_ok_xlsx.xlsx")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0020_XLSX_XLSM(self.workbook), 1)

    def test_notok_xlsx(self):
        self.workbook = openpyxl.load_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0020/doc4_notok_xlsx.xlsx")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0020_XLSX_XLSM(self.workbook), 0)


class Test_02043_18_04939_STRUCT_0025_0030_0035_0040(unittest.TestCase):
    def setUp(self):
        self.app = TSD_Checker.QApplication(TSD_Checker.sys.argv)
        self.TSD_checker = TSD_Checker.Test()

    def test_ok_xls(self):
        self.workbook = xlrd.open_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0025_0030_0035_0040/doc4_ok_xls.xls")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0025_XLS(self.workbook), 1)

    def test_ok_french_xls(self):
        self.workbook = xlrd.open_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0025_0030_0035_0040/doc4_ok_french_xls.xls")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0025_XLS(self.workbook), 1)
        
    def test_notok_xls(self):
        self.workbook = xlrd.open_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0025_0030_0035_0040/doc4_notok_xls.xls")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0025_XLS(self.workbook), 0)


    def test_ok_xlsx(self):
        self.workbook =  openpyxl.load_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0025_0030_0035_0040/doc4_ok_xlsx.xlsx")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0025_XLSX_XLSM(self.workbook), 1)


    def test_ok_xls(self):
        self.workbook = xlrd.open_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0025_0030_0035_0040/doc4_ok_xls.xls")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0030_XLS(self.workbook), 1)



class Test_02043_18_04939_STRUCT_0051(unittest.TestCase):
    def setUp(self):
        self.app = TSD_Checker.QApplication(TSD_Checker.sys.argv)
        self.TSD_checker = TSD_Checker.Test()

    def test_ok_xls(self):
        self.workbook = xlrd.open_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0051/doc_ok_xls.xls")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0051_XLS(self.workbook), 1)

    def test_notok_xls(self):
        self.workbook = xlrd.open_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0051/doc_notok_xls.xls")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0051_XLS(self.workbook), 0)


    def test_ok_xlsx(self):
        self.workbook = openpyxl.load_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0051/doc_ok_xlsx.xlsx")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0051_XLSX_XLSM(self.workbook), 1)

    def test_notok_xlsx(self):
        self.workbook = openpyxl.load_workbook("C:/Users/admacesanu/Desktop/EXCEL_TEST/02043_18_04939_STRUCT_0051/doc_notok_xlsx.xlsx")
        self.assertEqual(self.TSD_checker.Test_02043_18_04939_STRUCT_0051_XLSX_XLSM(self.workbook), 0)



if __name__ == '__main__':
    unittest.main()