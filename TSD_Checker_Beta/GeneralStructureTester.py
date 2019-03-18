import TSD_Checker_V0_5_2
import inspect
from ExcelEdit import TestReturn as result
from ErrorMessages import errorMessagesDict as error

class WorkbookProperties:
    def __init__(self):
        self.sheetNames = list()
        self.hasInfGen = True
        self.InfGenIndex = 0
        self.hasSupp = True
        self.SuppIndex = 0
        self.hasRefDocs = True
        self.refDocsIndex = 0
        self.nameRefDocsIndex = 0
        self.refRefDocsIndex = 0
        self.hasTable = True
        self.hasCode = True
        self.hasMeasure = True
        self.tableIndex = 5
        self.tableLastRow = 0
        self.codeIndex = 6
        self.codeLastRow = 0
        self.measureIndex = 6
        self.measureLastRow = 0


def Test_02043_18_04939_STRUCT_0000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    temp = workBook.Sheets
    sheetNames = list()
    for sheet in temp:
        sheetNames.append(sheet.Name.strip().casefold())
    TSDApp.WorkbookStats.sheetNames = sheetNames
    if "informations générales" in sheetNames or "general information" in sheetNames:
        localisation = None
        try:
            index = sheetNames.index("informations générales") + 1
        except:
            index = sheetNames.index("general information") + 1
        TSDApp.WorkbookStats.InfGenIndex = index
    else:
        localisation = ""
        TSDApp.WorkbookStats.hasInfGen = False

    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)


def Test_02043_18_04939_STRUCT_0005(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasInfGen == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        if workBook.Sheets(TSDApp.WorkbookStats.InfGenIndex).Range("C49").HasFormula is False:
            localisation = None
        else:
            localisation = workBook.Sheets(TSDApp.WorkbookStats.InfGenIndex).Range("C49")
            add = localisation.Address

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)


def Test_02043_18_04939_STRUCT_0010(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasInfGen == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.InfGenIndex)
        cell = workSheet.Cells(52,2)
        if cell.Value is None:
            localisation = cell
        else:
            localisation = None
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)



def Test_02043_18_04939_STRUCT_0011(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasInfGen == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.InfGenIndex)
        cell = workSheet.Cells(52,2)

        if cell.Value in {"AEEV_IAEE07_0033", "02043_12_01665", "02043_12_01666"}:
            localisation = None
        else:
            localisation = cell
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)



def Test_02043_18_04939_STRUCT_0020(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "suppression" in TSDApp.WorkbookStats.sheetNames:
        index = TSDApp.WorkbookStats.sheetNames.index("suppression") + 1
        TSDApp.WorkbookStats.SuppIndex = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasSupp = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)


def Test_02043_18_04939_STRUCT_0025(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasSupp == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.SuppIndex)
        row1Values = workSheet.Rows(1).Value
        localisation = workSheet.Rows(1)
        row1Values = row1Values[0]
        for value in row1Values:
            if str(value).casefold().strip() in {"sheet", "onglet"}:
                localisation = None
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)




def Test_02043_18_04939_STRUCT_0030(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasSupp == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.SuppIndex)
        row1Values = workSheet.Rows(1).Value
        row1Values = row1Values[0]
        localisation = workSheet.Rows(1)
        for value in row1Values:
            if str(value).casefold().strip() in {"référence de la ligne", "line number"}:
                localisation = None
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)



def Test_02043_18_04939_STRUCT_0035(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasSupp == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.SuppIndex)
        row1Values = workSheet.Rows(1).Value
        row1Values = row1Values[0]
        localisation = workSheet.Rows(1)
        for value in row1Values:
            if str(value).casefold().strip() in {"version du tsd", "version of the document"}:
                localisation = None
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)



def Test_02043_18_04939_STRUCT_0040(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasSupp == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.SuppIndex)
        row1Values = workSheet.Rows(1).Value
        row1Values = row1Values[0]
        localisation = workSheet.Rows(1)
        for value in row1Values:
            if str(value).casefold().strip() in {"justification de la modification", "change reason"}:
                localisation = None
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)





def Test_02043_18_04939_STRUCT_0051(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "reference docs" in TSDApp.WorkbookStats.sheetNames:
        index = TSDApp.WorkbookStats.sheetNames.index("reference docs") + 1
        TSDApp.WorkbookStats.refDocsIndex = index
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.refDocsIndex)
        workSheetRange = workSheet.UsedRange
        for rowRange in workSheetRange:
            for cell in rowRange:
                if str(cell.Value).casefold().strip() == "name":
                    nameColIndex = cell.Column
                if str(cell.Value).strip().casefold() == "reference":
                    refColIndex = cell.Column
        colName = workSheetRange.Columns(nameColIndex)
        TSDApp.WorkbookStats.nameRefDocsIndex = nameColIndex
        colRef = workSheetRange.Columns(refColIndex)
        TSDApp.WorkbookStats.refRefDocsIndex = refColIndex
        localisation = None
        for cell in colName.Value:
                if str(cell[0]).casefold().strip() in [ "vehicle architecture schematic", "planche d'architecture véhicule"]:
                    if str(workSheet.Cells(colName.Value.index(cell) +1, refColIndex).Value).strip() in ["None",""]:
                        localisation = workSheet.Rows(colName.Value.index(cell) +1)
    else:
        TSDApp.WorkbookStats.hasRefDocs = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)


def Test_02043_18_04939_STRUCT_0052(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasRefDocs == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)

    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.refDocsIndex)
        workSheetRange = workSheet.UsedRange
        colName = workSheetRange.Columns(TSDApp.WorkbookStats.nameRefDocsIndex)
        colRef = workSheetRange.Columns(TSDApp.WorkbookStats.refRefDocsIndex)
        localisation = None
        for cell in colName.Value:
            if str(cell[0]).casefold().strip() in ["diagnostic matrix", "matrice diag"]:
                if str(workSheet.Cells(colName.Value.index(cell) + 1, TSDApp.WorkbookStats.refRefDocsIndex).Value).strip() in ["None", ""]:
                    localisation = workSheet.Rows(colName.Value.index(cell) + 1)
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

'''
def Test_02043_18_04939_STRUCT_0000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name

def Test_02043_18_04939_STRUCT_0000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name

def Test_02043_18_04939_STRUCT_0000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name

def Test_02043_18_04939_STRUCT_0000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name

def Test_02043_18_04939_STRUCT_0000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name

def Test_02043_18_04939_STRUCT_0000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name

def Test_02043_18_04939_STRUCT_0000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name

def Test_02043_18_04939_STRUCT_0000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name

def Test_02043_18_04939_STRUCT_0000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name

def Test_02043_18_04939_STRUCT_0000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name

def Test_02043_18_04939_STRUCT_0000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name

def Test_02043_18_04939_STRUCT_0000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name

'''




