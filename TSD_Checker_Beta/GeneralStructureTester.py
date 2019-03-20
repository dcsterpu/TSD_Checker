import TSD_Checker_V0_5_2
import inspect
from ExcelEdit import TestReturn as result
from ErrorMessages import errorMessagesDict as error

class WorkbookProperties:
    def __init__(self):
        self.sheetNames = list()
        self.hasInfGen = True
        self.hasTable = True
        self.hasCode = True
        self.hasMDD = True
        self.hasSitDeVie = True
        self.hasConstituants = True
        self.hasER = True
        self.hasEffClients = True
        self.hasDiagDeb = True
        self.hasMeasure = True
        self.hasSupp = True
        self.hasRefDocs = True

        self.InfGenIndex = 0
        self.SuppIndex = 0
        self.refDocsIndex = 0
        self.nameRefDocsIndex = 0
        self.refRefDocsIndex = 0
        self.tableIndex = 0
        self.codeIndex = 0
        self.codeLastRow = 0
        self.measureIndex = 0
        self.DiagDebIndex = 0
        self.EffClientsIndex = 0
        self.ERIndex = 0
        self.constituantsIndex = 0
        self.SitDeVieIndex = 0
        self.MDDIndex = 0
#DOC4
        self.hasTableDoc4 = True
        self.tableIndexDOC4 = 0
        self.hasDiagNeeds = True
        self.DiagNeedsIndex = 0
        self.hasCustEff = True
        self.CustEffIndex = 0
        self.hasFearedEvent = True
        self.FearedEvent = 0
        self.hasSystem = True
        self.SystemIndex = 0
        self.hasOpSit = True
        self.OpSitIndex = 0
        self.hasTechEff = True
        self.TechEffIndex = 0
#DOC5
        self.hasTableDoc5 = True
        self.tableIndexDOC5 = 0
        self.hasDataCodes = True
        self.DataCodesIndex = 0
        self.hasReadDataIO = True
        self.ReadDataIOIndex = 0
        self.hasNotEmbDiag = True
        self.NotEmbDiagIndex = 0
        self.hasCustEffDoc5 = True
        self.CustEffIndexDoc5 = 0
        self.hasFearedEventDoc5 = True
        self.FearedEventIndexDoc5 = 0
        self.hasNotEmbDiag = True
        self.NotEmbDiagIndex = 0
        self.hasConstituantsDoc5 = True
        self.ConstituantsIndexDoc5 = 0
        self.hasSitDeVieDoc5 = True
        self.SitDeVieIndexDoc5 = 0
        self.hasMDDDoc5 = True
        self.MDDIndexDoc5 = 0
        self.hasTechEffDoc5 = True
        self.TechEffIndexDoc5 = 0
        self.hasVariant = True
        self.VariantIndex = 0
        self.hasNotEmbDiag = True
        self.NotEmbDiagIndex = 0

        self.tableLastRow = 0
        self.measureLastRow = 0
        self.DiagDebLastRow = 0
        self.MDDLastRow = 0

#General Structure

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

def Test_02043_18_04939_STRUCT_0053(workBook, TSDApp):
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
            if str(cell[0]).casefold().strip() in ["fault tree", "amdec"]:
                if str(workSheet.Cells(colName.Value.index(cell) + 1, TSDApp.WorkbookStats.refRefDocsIndex).Value).strip() in ["None", ""]:
                    localisation = workSheet.Rows(colName.Value.index(cell) + 1)
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0054(workBook, TSDApp):
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
            if str(cell[0]).casefold().strip() in ["ecu schematic", "synoptique ecu"]:
                if str(workSheet.Cells(colName.Value.index(cell) + 1, TSDApp.WorkbookStats.refRefDocsIndex).Value).strip() in ["None", ""]:
                    localisation = workSheet.Rows(colName.Value.index(cell) + 1)
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0055(workBook, TSDApp):
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
            if str(cell[0]).casefold().strip() in ["std"]:
                if str(workSheet.Cells(colName.Value.index(cell) + 1, TSDApp.WorkbookStats.refRefDocsIndex).Value).strip() in ["None", ""]:
                    localisation = workSheet.Rows(colName.Value.index(cell) + 1)
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0056(workBook, TSDApp):
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
            if str(cell[0]).casefold().strip() in ["complexity matrix (decli ee)"]:
                if str(workSheet.Cells(colName.Value.index(cell) + 1, TSDApp.WorkbookStats.refRefDocsIndex).Value).strip() in ["None", ""]:
                    localisation = workSheet.Rows(colName.Value.index(cell) + 1)
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0057(workBook, TSDApp):
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
            if str(cell[0]).casefold().strip() in ["décli"]:
                if str(workSheet.Cells(colName.Value.index(cell) + 1, TSDApp.WorkbookStats.refRefDocsIndex).Value).strip() in ["None", ""]:
                    localisation = workSheet.Rows(colName.Value.index(cell) + 1)
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0058(workBook, TSDApp):
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
            if str(cell[0]).casefold().strip() in ["dcee"]:
                if str(workSheet.Cells(colName.Value.index(cell) + 1, TSDApp.WorkbookStats.refRefDocsIndex).Value).strip() in ["None", ""]:
                    localisation = workSheet.Rows(colName.Value.index(cell) + 1)
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0059(workBook, TSDApp):
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
            if str(cell[0]).casefold().strip() in ["eead"]:
                if str(workSheet.Cells(colName.Value.index(cell) + 1, TSDApp.WorkbookStats.refRefDocsIndex).Value).strip() in ["None", ""]:
                    localisation = workSheet.Rows(colName.Value.index(cell) + 1)
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0060(workBook, TSDApp):
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
            if str(cell[0]).casefold().strip() in ["tfd"]:
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

'''

#[DOC3]

def Test_02043_18_04939_STRUCT_0100(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "tableau" in TSDApp.WorkbookStats.sheetNames:
        index = TSDApp.WorkbookStats.sheetNames.index("tableau") + 1
        TSDApp.WorkbookStats.tableIndex = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasTable = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0110(ExcelApp, workBook, TSDApp, DOC3Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        list_test = list()


        for row in range(4,5):
            for col in range(1,nrCols):
                dict = {}
                dict['1'] = workSheet.Cells(row - 1, col).Value
                dict['2'] = workSheet.Cells(row, col).Value
                dict['3'] = col
                dict['4'] = row
                list_test.append(dict)


        DOC3 = ExcelApp.Workbooks.Open(DOC3Name)
        workSheetRef = DOC3.Sheets("tableau")

        list_ref = list()

        for row in range(4, 5):
            for col in range(1, nrCols):
                dict = {}
                dict['1'] = workSheetRef.Cells(row - 1, col).Value
                dict['2'] = workSheetRef.Cells(row, col).Value
                dict['3'] = col
                dict['4'] = row
                list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['3']
                RowIndex = element['4']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0120(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "codes défauts" in TSDApp.WorkbookStats.sheetNames:
        index = TSDApp.WorkbookStats.sheetNames.index("codes défauts") + 1
        TSDApp.WorkbookStats.codeIndex = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasCode = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0130(ExcelApp, workBook, TSDApp, DOC3Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count

        list_test = list()


        for row in range(4,5):
            for col in range(1,nrCols):
                dict = {}
                dict['1'] = workSheet.Cells(row - 1, col).Value
                dict['2'] = workSheet.Cells(row, col).Value
                dict['3'] = col
                dict['4'] = row
                list_test.append(dict)


        DOC3 = ExcelApp.Workbooks.Open(DOC3Name)
        workSheetRef = DOC3.Sheets("codes défauts")

        list_ref = list()

        for row in range(4, 5):
            for col in range(1, nrCols):
                dict = {}
                dict['1'] = workSheetRef.Cells(row - 1, col).Value
                dict['2'] = workSheetRef.Cells(row, col).Value
                dict['3'] = col
                dict['4'] = row
                list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['3']
                RowIndex = element['4']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0140(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "mesures et commandes" in TSDApp.WorkbookStats.sheetNames:
        index = TSDApp.WorkbookStats.sheetNames.index("mesures et commandes") + 1
        TSDApp.WorkbookStats.measureIndex = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasMeasure = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0150(ExcelApp, workBook, TSDApp, DOC3Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count

        list_test = list()


        for row in range(4,5):
            for col in range(1,nrCols):
                dict = {}
                dict['1'] = workSheet.Cells(row - 1, col).Value
                dict['2'] = workSheet.Cells(row, col).Value
                dict['3'] = col
                dict['4'] = row
                list_test.append(dict)


        DOC3 = ExcelApp.Workbooks.Open(DOC3Name)
        workSheetRef = DOC3.Sheets("mesures et commandes")

        list_ref = list()

        for row in range(4, 5):
            for col in range(1, nrCols):
                dict = {}
                dict['1'] = workSheetRef.Cells(row - 1, col).Value
                dict['2'] = workSheetRef.Cells(row, col).Value
                dict['3'] = col
                dict['4'] = row
                list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['3']
                RowIndex = element['4']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0160(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "diagnostic débarqués" in TSDApp.WorkbookStats.sheetNames:
        index = TSDApp.WorkbookStats.sheetNames.index("diagnostic débarqués") + 1
        TSDApp.WorkbookStats.DiagDebIndex = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasDiagDeb = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0170(ExcelApp, workBook, TSDApp, DOC3Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagDebIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count

        list_test = list()


        for row in range(4,5):
            for col in range(1,nrCols):
                dict = {}
                dict['1'] = workSheet.Cells(row - 1, col).Value
                dict['2'] = workSheet.Cells(row, col).Value
                dict['3'] = col
                dict['4'] = row
                list_test.append(dict)


        DOC3 = ExcelApp.Workbooks.Open(DOC3Name)
        workSheetRef = DOC3.Sheets("diagnostic débarqués")

        list_ref = list()

        for row in range(4, 5):
            for col in range(1, nrCols):
                dict = {}
                dict['1'] = workSheetRef.Cells(row - 1, col).Value
                dict['2'] = workSheetRef.Cells(row, col).Value
                dict['3'] = col
                dict['4'] = row
                list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['3']
                RowIndex = element['4']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0180(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "effets clients" in TSDApp.WorkbookStats.sheetNames:
        index = TSDApp.WorkbookStats.sheetNames.index("effets clients") + 1
        TSDApp.WorkbookStats.EffClientsIndex = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasEffClients = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0190(ExcelApp, workBook, TSDApp, DOC3Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.EffClientsIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count

        list_test = list()


        for row in range(4,5):
            for col in range(1,nrCols):
                dict = {}
                dict['1'] = workSheet.Cells(row - 1, col).Value
                dict['2'] = workSheet.Cells(row, col).Value
                dict['3'] = col
                dict['4'] = row
                list_test.append(dict)


        DOC3 = ExcelApp.Workbooks.Open(DOC3Name)
        workSheetRef = DOC3.Sheets("effets clients")

        list_ref = list()

        for row in range(4, 5):
            for col in range(1, nrCols):
                dict = {}
                dict['1'] = workSheetRef.Cells(row - 1, col).Value
                dict['2'] = workSheetRef.Cells(row, col).Value
                dict['3'] = col
                dict['4'] = row
                list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['3']
                RowIndex = element['4']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0200(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "er" in TSDApp.WorkbookStats.sheetNames:
        index = TSDApp.WorkbookStats.sheetNames.index("er") + 1
        TSDApp.WorkbookStats.ERIndex = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasER = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0210(ExcelApp, workBook, TSDApp, DOC3Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasER == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ERIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count

        list_test = list()


        for row in range(4,5):
            for col in range(1,nrCols):
                dict = {}
                dict['1'] = workSheet.Cells(row - 1, col).Value
                dict['2'] = workSheet.Cells(row, col).Value
                dict['3'] = col
                dict['4'] = row
                list_test.append(dict)


        DOC3 = ExcelApp.Workbooks.Open(DOC3Name)
        workSheetRef = DOC3.Sheets("er")

        list_ref = list()

        for row in range(4, 5):
            for col in range(1, nrCols):
                dict = {}
                dict['1'] = workSheetRef.Cells(row - 1, col).Value
                dict['2'] = workSheetRef.Cells(row, col).Value
                dict['3'] = col
                dict['4'] = row
                list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['3']
                RowIndex = element['4']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0220(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "constituants" in TSDApp.WorkbookStats.sheetNames:
        index = TSDApp.WorkbookStats.sheetNames.index("constituants") + 1
        TSDApp.WorkbookStats.constituantsIndex = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasConstituants = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0230(ExcelApp, workBook, TSDApp, DOC3Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasConstituants == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.constituantsIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count

        list_test = list()


        for row in range(4,5):
            for col in range(1,nrCols):
                dict = {}
                dict['1'] = workSheet.Cells(row - 1, col).Value
                dict['2'] = workSheet.Cells(row, col).Value
                dict['3'] = col
                dict['4'] = row
                list_test.append(dict)


        DOC3 = ExcelApp.Workbooks.Open(DOC3Name)
        workSheetRef = DOC3.Sheets("constituants")

        list_ref = list()

        for row in range(4, 5):
            for col in range(1, nrCols):
                dict = {}
                dict['1'] = workSheetRef.Cells(row - 1, col).Value
                dict['2'] = workSheetRef.Cells(row, col).Value
                dict['3'] = col
                dict['4'] = row
                list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['3']
                RowIndex = element['4']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0240(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "situations de vie" in TSDApp.WorkbookStats.sheetNames:
        index = TSDApp.WorkbookStats.sheetNames.index("situations de vie") + 1
        TSDApp.WorkbookStats.SitDeVieIndex = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasSitDeVie = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0250(ExcelApp, workBook, TSDApp, DOC3Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasSitDeVie == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.SitDeVieIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count

        list_test = list()


        for row in range(4,5):
            for col in range(1, nrCols):
                dict = {}
                dict['1'] = workSheet.Cells(row - 1, col).Value
                dict['2'] = workSheet.Cells(row, col).Value
                dict['3'] = col
                dict['4'] = row
                list_test.append(dict)


        DOC3 = ExcelApp.Workbooks.Open(DOC3Name)
        workSheetRef = DOC3.Sheets("situations de vie")

        list_ref = list()

        for row in range(4, 5):
            for col in range(1, nrCols):
                dict = {}
                dict['1'] = workSheetRef.Cells(row - 1, col).Value
                dict['2'] = workSheetRef.Cells(row, col).Value
                dict['3'] = col
                dict['4'] = row
                list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['3']
                RowIndex = element['4']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0260(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "liste mdd" in TSDApp.WorkbookStats.sheetNames:
        index = TSDApp.WorkbookStats.sheetNames.index("liste mdd") + 1
        TSDApp.WorkbookStats.MDDIndex = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasMDD = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0270(ExcelApp, workBook, TSDApp, DOC3Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasMDD == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.MDDIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count

        list_test = list()


        for row in range(2,3):
            for col in range(1,nrCols):
                dict = {}
                dict['1'] = workSheet.Cells(row - 1, col).Value
                dict['2'] = workSheet.Cells(row, col).Value
                dict['3'] = col
                dict['4'] = row
                list_test.append(dict)


        DOC3 = ExcelApp.Workbooks.Open(DOC3Name)
        workSheetRef = DOC3.Sheets("liste mdd")

        list_ref = list()

        for row in range(4, 5):
            for col in range(1, nrCols):
                dict = {}
                dict['1'] = workSheetRef.Cells(row - 1, col).Value
                dict['2'] = workSheetRef.Cells(row, col).Value
                dict['3'] = col
                dict['4'] = row
                list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['3']
                RowIndex = element['4']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)



#DOC4

def Test_02043_18_04939_STRUCT_0400(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    temp = workBook.Sheets
    sheetNames = list()
    for sheet in temp:
        sheetNames.append(sheet.Name.strip().casefold())
    TSDApp.WorkbookStats.sheetNames = sheetNames
    if "tableau" in TSDApp.WorkbookStats.sheetNames or "table" in TSDApp.WorkbookStats.sheetNames:
        try:
            index = TSDApp.WorkbookStats.sheetNames.index("tableau") + 1
        except:
            index = TSDApp.WorkbookStats.sheetNames.index("table") + 1
        TSDApp.WorkbookStats.tableIndexDOC4 = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasTableDoc4 = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0410(ExcelApp, workBook, TSDApp, DOC4Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTableDoc4 == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndexDOC4)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count

        list_test = list()


        for row in range(3,5):
            for col in range(1,nrCols):
                dict = {}
                dict['1'] = workSheet.Cells(row - 2, col).Value
                dict['2'] = workSheet.Cells(row - 1, col).Value
                dict['3'] = workSheet.Cells(row, col).Value
                dict['4'] = col
                dict['5'] = row
                list_test.append(dict)


        DOC4 = ExcelApp.Workbooks.Open(DOC4Name)

        try:
            workSheetRef = DOC4.Sheets("tableau")
        except:
            workSheetRef = DOC4.Sheets("table")

        list_ref = list()

        for row in range(3, 5):
            for col in range(1, nrCols):
                dict = {}
                dict['1'] = workSheetRef.Cells(row - 2, col).Value
                dict['2'] = workSheetRef.Cells(row - 1, col).Value
                dict['3'] = workSheetRef.Cells(row, col).Value
                dict['4'] = col
                dict['5'] = row
                list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['4']
                RowIndex = element['5']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0420(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "diagnostic needs" in TSDApp.WorkbookStats.sheetNames:
        index = TSDApp.WorkbookStats.sheetNames.index("diagnostic needs") + 1
        TSDApp.WorkbookStats.DiagNeedsIndex = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasDiagNeeds = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0430(ExcelApp, workBook, TSDApp, DOC4Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagNeedsIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count

        list_test = list()

        row = 1
        for col in range(1,nrCols):
            dict = {}
            dict['1'] = workSheet.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_test.append(dict)


        DOC4 = ExcelApp.Workbooks.Open(DOC4Name)
        workSheetRef = DOC4.Sheets("diagnostic needs")

        list_ref = list()
        row = 1
        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_ref.append(dict)

        localisation = list()



        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['2']
                RowIndex = element['3']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0440(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "customer effects" in TSDApp.WorkbookStats.sheetNames:
        index = TSDApp.WorkbookStats.sheetNames.index("customer effects") + 1
        TSDApp.WorkbookStats.CustEffIndex = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasCustEff = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0450(ExcelApp, workBook, TSDApp, DOC4Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasCustEff == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.CustEffIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        list_test = list()


        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheet.Cells(1, col).Value
            dict['2'] = col
            dict['3'] = 1
            list_test.append(dict)


        DOC4 = ExcelApp.Workbooks.Open(DOC4Name)
        workSheetRef = DOC4.Sheets("customer effects")

        list_ref = list()

        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.Cells(1, col).Value
            dict['2'] = col
            dict['3'] = 1
            list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['2']
                RowIndex = element['3']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0460(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "feared events" in TSDApp.WorkbookStats.sheetNames or "er" in TSDApp.WorkbookStats.sheetNames:
        try:
            index = TSDApp.WorkbookStats.sheetNames.index("feared events") + 1
        except:
            index = TSDApp.WorkbookStats.sheetNames.index("er") + 1
        TSDApp.WorkbookStats.FearedEventIndex = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasFearedEvent = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0470(ExcelApp, workBook, TSDApp, DOC4Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.FearedEventIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        list_test = list()


        row = 1

        for col in range(1,nrCols):
            dict = {}
            dict['1'] = workSheet.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_test.append(dict)


        DOC4 = ExcelApp.Workbooks.Open(DOC4Name)
        try:
            workSheetRef = DOC4.Sheets("feared events")
        except:
            workSheetRef = DOC4.Sheets("er")

        list_ref = list()

        row = 1

        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['2']
                RowIndex = element['3']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0480(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "system" in TSDApp.WorkbookStats.sheetNames or "système":
        try:
            index = TSDApp.WorkbookStats.sheetNames.index("system") + 1
        except:
            index = TSDApp.WorkbookStats.sheetNames.index("système") + 1
        TSDApp.WorkbookStats.SystemIndex = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasSystem = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0490(ExcelApp, workBook, TSDApp, DOC4Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasSystem == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.SystemIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        list_test = list()

        row = 1
        for col in range(1,nrCols):
            dict = {}
            dict['1'] = workSheet.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_test.append(dict)


        DOC4 = ExcelApp.Workbooks.Open(DOC4Name)
        try:
            workSheetRef = DOC4.Sheets("system")
        except:
            workSheetRef = DOC4.Sheets("système")

        list_ref = list()

        row = 1
        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['2']
                RowIndex = element['3']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0500(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "operation situation" in TSDApp.WorkbookStats.sheetNames:
        index = TSDApp.WorkbookStats.sheetNames.index("operation situation") + 1
        TSDApp.WorkbookStats.OpSitIndex = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasOpSit = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0510(ExcelApp, workBook, TSDApp, DOC4Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasOpSit == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.OpSitIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count

        list_test = list()

        row = 1
        for col in range(1,nrCols):
            dict = {}
            dict['1'] = workSheet.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_test.append(dict)


        DOC4 = ExcelApp.Workbooks.Open(DOC4Name)
        workSheetRef = DOC4.Sheets("operation situation")

        list_ref = list()

        row = 1
        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['2']
                RowIndex = element['3']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0520(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "req. of tech. effects" in TSDApp.WorkbookStats.sheetNames:
        index = TSDApp.WorkbookStats.sheetNames.index("req. of tech. effects") + 1
        TSDApp.WorkbookStats.TechEffIndex = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasTechEff = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0530(ExcelApp, workBook, TSDApp, DOC4Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTechEff == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.TechEffIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count

        list_test = list()

        row = 1
        for col in range(1,nrCols):
            dict = {}
            dict['1'] = workSheet.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_test.append(dict)


        DOC4 = ExcelApp.Workbooks.Open(DOC4Name)
        workSheetRef = DOC4.Sheets("req. of tech. effects")

        list_ref = list()

        row = 1
        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['2']
                RowIndex = element['3']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)


#DOC5


def Test_02043_18_04939_STRUCT_0700(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    temp = workBook.Sheets
    sheetNames = list()
    for sheet in temp:
        sheetNames.append(sheet.Name.strip().casefold())
    TSDApp.WorkbookStats.sheetNames = sheetNames
    if "tableau" in TSDApp.WorkbookStats.sheetNames or "table" in TSDApp.WorkbookStats.sheetNames:
        try:
            index = TSDApp.WorkbookStats.sheetNames.index("tableau") + 1
        except:
            index = TSDApp.WorkbookStats.sheetNames.index("table") + 1
        TSDApp.WorkbookStats.tableIndexDOC5 = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasTableDoc5 = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0710(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTableDoc5 == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndexDOC5)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count

        list_test = list()


        for row in range(2,3):
            for col in range(1,nrCols):
                dict = {}
                dict['1'] = workSheet.Cells(row - 1, col).Value
                dict['2'] = workSheet.Cells(row, col).Value
                dict['3'] = col
                dict['4'] = row
                list_test.append(dict)


        DOC5 = ExcelApp.Workbooks.Open(DOC5Name)

        try:
            workSheetRef = DOC5.Sheets("tableau")
        except:
            workSheetRef = DOC5.Sheets("table")

        list_ref = list()

        for row in range(2, 3):
            for col in range(1, nrCols):
                dict = {}
                dict['1'] = workSheetRef.Cells(row - 1, col).Value
                dict['2'] = workSheetRef.Cells(row, col).Value
                dict['3'] = col
                dict['4'] = row
                list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['3']
                RowIndex = element['4']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0720(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "data trouble codes" in TSDApp.WorkbookStats.sheetNames or "codes défauts" in TSDApp.WorkbookStats.sheetNames:
        try:
            index = TSDApp.WorkbookStats.sheetNames.index("data trouble codes") + 1
        except:
            index = TSDApp.WorkbookStats.sheetNames.index("codes défauts") + 1
        TSDApp.WorkbookStats.DataCodesIndex = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasDataCodes = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0730(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DataCodesIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count

        list_test = list()


        for col in range(1,nrCols):
            dict = {}
            dict['1'] = workSheet.Cells(1, col).Value
            dict['2'] = col
            dict['3'] = 1
            list_test.append(dict)


        DOC5 = ExcelApp.Workbooks.Open(DOC5Name)
        try:
            workSheetRef = DOC5.Sheets("data trouble codes")
        except:
            workSheetRef = DOC5.Sheets("codes défauts")


        list_ref = list()

        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.Cells(1, col).Value
            dict['2'] = col
            dict['3'] = 1
            list_ref.append(dict)

        localisation = list()



        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['2']
                RowIndex = element['3']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0740(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "read data and io control" in TSDApp.WorkbookStats.sheetNames or "mesures et commandes" in TSDApp.WorkbookStats.sheetNames:
        try:
            index = TSDApp.WorkbookStats.sheetNames.index("read data and io control") + 1
        except:
            index = TSDApp.WorkbookStats.sheetNames.index("mesures et commandes") + 1
        TSDApp.WorkbookStats.ReadDataIOIndex = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasReadDataIO = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0750(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReadDataIOIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        list_test = list()


        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheet.Cells(1, col).Value
            dict['2'] = col
            dict['3'] = 1
            list_test.append(dict)


        DOC5 = ExcelApp.Workbooks.Open(DOC5Name)
        try:
            workSheetRef = DOC5.Sheets("read data and io control")
        except:
            workSheetRef = DOC5.Sheets("mesures et commandes")

        list_ref = list()

        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.Cells(1, col).Value
            dict['2'] = col
            dict['3'] = 1
            list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['2']
                RowIndex = element['3']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0760(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "not embedded diagnosis" in TSDApp.WorkbookStats.sheetNames or "read data and io control" in TSDApp.WorkbookStats.sheetNames:
        try:
            index = TSDApp.WorkbookStats.sheetNames.index("not embedded diagnosis") + 1
        except:
            index = TSDApp.WorkbookStats.sheetNames.index("read data and io control") + 1
        TSDApp.WorkbookStats.NotEmbDiagIndex = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasNotEmbDiag = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0770(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasNotEmbDiag == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.NotEmbDiagIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        list_test = list()


        row = 1

        for col in range(1,nrCols):
            dict = {}
            dict['1'] = workSheet.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_test.append(dict)


        DOC5 = ExcelApp.Workbooks.Open(DOC5Name)
        try:
            workSheetRef = DOC5.Sheets("not embedded diagnosis")
        except:
            workSheetRef = DOC5.Sheets("read data and io control")

        list_ref = list()

        row = 1

        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['2']
                RowIndex = element['3']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0780(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "customer effect" in TSDApp.WorkbookStats.sheetNames or "effets clients":
        try:
            index = TSDApp.WorkbookStats.sheetNames.index("customer effect") + 1
        except:
            index = TSDApp.WorkbookStats.sheetNames.index("effets clients") + 1
        TSDApp.WorkbookStats.CustEffIndexDoc5 = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasCustEffDoc5 = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0790(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasCustEffDoc5 == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.CustEffIndexDoc5)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        list_test = list()

        row = 1
        for col in range(1,nrCols):
            dict = {}
            dict['1'] = workSheet.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_test.append(dict)


        DOC5 = ExcelApp.Workbooks.Open(DOC5Name)
        try:
            workSheetRef = DOC5.Sheets("customer effect")
        except:
            workSheetRef = DOC5.Sheets("effets clients")

        list_ref = list()

        row = 1
        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['2']
                RowIndex = element['3']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0800(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "feared events" in TSDApp.WorkbookStats.sheetNames or "er" in TSDApp.WorkbookStats.sheetNames:
        try:
            index = TSDApp.WorkbookStats.sheetNames.index("feared events") + 1
        except:
            index = TSDApp.WorkbookStats.sheetNames.index("er") + 1
        TSDApp.WorkbookStats.FearedEventIndexDoc5 = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasFearedEventDoc5 = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0810(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasFearedEventDoc5 == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.FearedEventIndexDoc5)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count

        list_test = list()

        row = 1
        for col in range(1,nrCols):
            dict = {}
            dict['1'] = workSheet.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_test.append(dict)


        DOC5 = ExcelApp.Workbooks.Open(DOC5Name)
        try:
            workSheetRef = DOC5.Sheets("feared events")
        except:
            workSheetRef = DOC5.Sheets("er")

        list_ref = list()

        row = 1
        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['2']
                RowIndex = element['3']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0820(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "parts" in TSDApp.WorkbookStats.sheetNames or "constituants" in TSDApp.WorkbookStats.sheetNames:
        try:
            index = TSDApp.WorkbookStats.sheetNames.index("parts") + 1
        except:
            index = TSDApp.WorkbookStats.sheetNames.index("constituants") + 1
        TSDApp.WorkbookStats.ConstituantsIndexDoc5 = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasConstituantsDoc5 = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0830(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasConstituantsDoc5 == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ConstituantsIndexDoc5)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count

        list_test = list()

        row = 1
        for col in range(1,nrCols):
            dict = {}
            dict['1'] = workSheet.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_test.append(dict)


        DOC5 = ExcelApp.Workbooks.Open(DOC5Name)
        try:
            workSheetRef = DOC5.Sheets("constituants")
        except:
            workSheetRef = DOC5.Sheets("parts")

        list_ref = list()

        row = 1
        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['2']
                RowIndex = element['3']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0840(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "situation" in TSDApp.WorkbookStats.sheetNames or "situation de vie" in TSDApp.WorkbookStats.sheetNames:
        try:
            index = TSDApp.WorkbookStats.sheetNames.index("situation") + 1
        except:
            index = TSDApp.WorkbookStats.sheetNames.index("situation de vie") + 1
        TSDApp.WorkbookStats.SitDeVieIndexDoc5 = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasSitDeVieDoc5 = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0850(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasSitDeVieDoc5 == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.SitDeVieIndexDoc5)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count

        list_test = list()

        row = 1
        for col in range(1,nrCols):
            dict = {}
            dict['1'] = workSheet.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_test.append(dict)


        DOC5 = ExcelApp.Workbooks.Open(DOC5Name)
        try:
            workSheetRef = DOC5.Sheets("situation")
        except:
            workSheetRef = DOC5.Sheets("situation de vie")

        list_ref = list()

        row = 1
        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['2']
                RowIndex = element['3']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0860(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "degraded mode" in TSDApp.WorkbookStats.sheetNames or "liste mdd" in TSDApp.WorkbookStats.sheetNames:
        try:
            index = TSDApp.WorkbookStats.sheetNames.index("degraded mode") + 1
        except:
            index = TSDApp.WorkbookStats.sheetNames.index("liste mdd") + 1
        TSDApp.WorkbookStats.MDDIndexDoc5 = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasMDDDoc5 = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0870(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasMDDDoc5 == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.MDDIndexDoc5)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count

        list_test = list()

        row = 1
        for col in range(1,nrCols):
            dict = {}
            dict['1'] = workSheet.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_test.append(dict)


        DOC5 = ExcelApp.Workbooks.Open(DOC5Name)
        try:
            workSheetRef = DOC5.Sheets("degraded mode")
        except:
            workSheetRef = DOC5.Sheets("liste mdd")

        list_ref = list()

        row = 1
        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['2']
                RowIndex = element['3']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0880(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "technical effect" in TSDApp.WorkbookStats.sheetNames or "effets techniques" in TSDApp.WorkbookStats.sheetNames:
        try:
            index = TSDApp.WorkbookStats.sheetNames.index("technical effect") + 1
        except:
            index = TSDApp.WorkbookStats.sheetNames.index("effets techniques") + 1
        TSDApp.WorkbookStats.TechEffIndexDoc5 = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasTechEffDoc5 = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0890(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTechEffDoc5 == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.TechEffIndexDoc5)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count

        list_test = list()

        row = 1
        for col in range(1,nrCols):
            dict = {}
            dict['1'] = workSheet.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_test.append(dict)


        DOC5 = ExcelApp.Workbooks.Open(DOC5Name)
        try:
            workSheetRef = DOC5.Sheets("technical effect")
        except:
            workSheetRef = DOC5.Sheets("effets techniques")


        list_ref = list()

        row = 1
        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['2']
                RowIndex = element['3']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0900(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "variant" in TSDApp.WorkbookStats.sheetNames or "variantes" in TSDApp.WorkbookStats.sheetNames:
        try:
            index = TSDApp.WorkbookStats.sheetNames.index("variant") + 1
        except:
            index = TSDApp.WorkbookStats.sheetNames.index("variantes") + 1
        TSDApp.WorkbookStats.VariantIndex = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasVariant = False
        localisation = ""
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0910(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasVariant == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.VariantIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count

        list_test = list()

        row = 1
        for col in range(1,nrCols):
            dict = {}
            dict['1'] = workSheet.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_test.append(dict)


        DOC5 = ExcelApp.Workbooks.Open(DOC5Name)
        try:
            workSheetRef = DOC5.Sheets("variant")
        except:
            workSheetRef = DOC5.Sheets("variantes")


        list_ref = list()

        row = 1
        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.Cells(row, col).Value
            dict['2'] = col
            dict['3'] = row
            list_ref.append(dict)

        localisation = list()

        for element in list_ref:
            if element in list_test:
                pass
            else:
                ColIndex = element['2']
                RowIndex = element['3']
                localisation.append(workSheet.Cells(RowIndex, ColIndex))

        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
