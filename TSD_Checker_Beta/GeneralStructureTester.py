import TSD_Checker_V7_5
import inspect
from ExcelEdit import TestReturn as result
from ExcelEdit import TestReturnName as show
from ErrorMessages import errorMessagesDict as error
import xlrd

class WorkbookProperties:
    def __init__(self):
        self.sheetNames = list()
        self.hasInfGen = False
        self.hasTable = False
        self.hasCode = False
        self.hasMDD = False
        self.hasSitDeVie = False
        self.hasConstituants = False
        self.hasER = False
        self.hasEffClients = False
        self.hasNotEmbDiag = False
        self.hasDiagDeb = False
        self.hasMeasure = False
        self.hasSupp = False
        self.hasRefDocs = False
        self.hasSituation = False
        self.hasDegradedMode = False

        self.tableLanguage = ""
        self.codeLanguage = ""
        self.DataCodesLanguage = ""
        self.TechEffLanguage = ""
        self.EffClientsLanguage = ""
        self.FearedEventLanguage = ""
        self.PartsLanguage = ""
        self.VariantLanguage = ""
        self.SituationLanguage = ""
        self.DegradedModeLanguage = ""
        self.SystemLanguage = ""

        self.DegradedModeIndex = 0
        self.DegradedModeLastRow = 0
        self.DegradedModeLastCol = 0
        self.SituationIndex = 0
        self.SituationLastCol = 0
        self.SituationLastRow = 0
        self.constituantsRefRowIndex = 0
        self.constituantsRefColIndex = 0
        self.ReqTechRefColIndex = 0
        self.ReqTechRefRowIndex = 0
        self.tableRefColIndex = 0
        self.tableRefRowIndex = 0
        self.codeRefColIndex = 0
        self.codeRefRowIndex = 0
        self.measureRefColIndex = 0
        self.measureRefRowIndex = 0
        self.DiagDebRefColIndex = 0
        self.DiagDebRefRowIndex = 0
        self.MDDRefColIndex = 0
        self.MDDRefRowIndex = 0

        self.InfGenIndex = 0
        self.SuppIndex = 0
        self.refDocsIndex = 0
        self.nameRefDocsIndex = 0
        self.refRefDocsIndex = 0
        self.tableIndex = 0
        self.codeIndex = 0
        self.codeLastRow = 0
        self.codeLastCol= 0
        self.DiagDebLastRow = 0
        self.DiagDebLastCol = 0
        self.measureIndex = 0
        self.DiagDebIndex = 0
        self.EffClientsIndex = 0
        self.ERIndex = 0
        self.constituantsIndex = 0
        self.SitDeVieIndex = 0
        self.MDDIndex = 0
#DOC4
        self.hasTable = False
        self.tableIndex = 0
        self.hasDiagNeeds = False
        self.DiagNeedsIndex = 0
        self.DiagNeedsLastRow = 0
        self.DiagNeedsLastCol = 0
        #self.hasCustEff = True
        #self.CustEffIndex = 0
        self.hasFearedEvent = False
        self.FearedEvent = 0
        self.hasSystem = False
        self.SystemIndex = 0
        self.hasOpSit = False
        self.OpSitIndex = 0
        self.hasTechEff = False
        self.TechEffIndex = 0
        self.hasReqTech = False
        self.ReqTechIndex = 0
#DOC5
        self.hasParts = False
        self.PartsIndex = 0
        self.PartsLastRow = 0
        self.PartsLastCol = 0
        self.TechEffLastRow = 0
        self.TechEffLastCol = 0
        self.OpSitLastRow = 0
        self.OpSitLastCol = 0
        self.SystemLastRow = 0
        self.SystemLastCol = 0
        self.NotEmbDiagLastRow = 0
        self.NotEmbDiagLastCol = 0
        self.DataCodesLastRow = 0
        self.DataCodesLastCol = 0
        self.hasDataCodes = False
        self.DataCodesIndex = 0
        self.ReadDataIOLastRow = 0
        self.ReadDataIOLastCol = 0
        self.hasReadDataIO = False
        self.ReadDataIOIndex = 0
        self.hasNotEmbDiag = False
        self.NotEmbDiagIndex = 0
        self.hasCustEff = False
        self.CustEffIndex = 0

        self.FearedEventIndex = 0
        self.FearedEventLastRow = 0
        self.FearedEventLastCol = 0
        self.hasNotEmbDiag = False
        self.hasVariant = False
        self.VariantIndex = 0
        self.VariantLastCol = 0
        self.VariantLastRow = 0
        self.hasNotEmbDiag = False
        self.SitDeVieLastRow = 0
        self.SitDeVieLastCol = 0

        self.ReqTechLastRow = 0
        self.ReqTechLastCol = 0
        self.EffClientsLastRow = 0
        self.EffClientsLastCol = 0
        self.tableLastRow = 0
        self.TableLastCol = 0
        self.measureLastRow = 0
        self.measureLastCol = 0
        self.MDDLastRow = 0
        self.MDDLastCol = 0
        self.codeLastRow = 0
        self.constituantsLastRow = 0
        self.constituantsLastCol = 0
        self.effLastRow = 0
        self.ERLastRow = 0
        self.ERLastCol = 0
        self.CustEffLastRow = 0
        self.TechLastRow = 0

        self.famillyList = list()

#General Structure


def Test_02043_18_04939_STRUCT_0000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    sheetNames = list()
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())
    TSDApp.WorkbookStats.sheetNames = sheetNames
    flag = False
    if "informations générales" in sheetNames or "general information" in sheetNames:
        TSDApp.WorkbookStats.hasInfGen = True
        localisation = None
        try:
            index = sheetNames.index("informations générales")
        except:
            index = sheetNames.index("general information")
        TSDApp.WorkbookStats.InfGenIndex = index
    else:
        localisation = ""
        TSDApp.WorkbookStats.hasInfGen = False
        flag = True

    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0005(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    localisation = []
    if TSDApp.WorkbookStats.hasInfGen == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.InfGenIndex)
        try:
            if str(workSheet.cell(51,1).value)[0] != "=":
                localisation = None
        except:
            localisation.append(("informations générales", 51, 1))

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0010(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    localisation = []
    if TSDApp.WorkbookStats.hasInfGen == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.InfGenIndex)
        cell = workSheet.cell(51,1)
        if cell.value is None:
            localisation.append(("informations générales", 51, 1))
        else:
            localisation = None
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def hasNumbers(inputString):
    return any(char.isdigit() for char in inputString)

def Test_02043_18_04939_STRUCT_0011(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    localisation = []
    if TSDApp.WorkbookStats.hasInfGen == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.InfGenIndex)
        rowIndex = -1
        colIndex = -1
        flag = False
        for index1 in range(0, workSheet.nrows):
            for index2 in range(0, workSheet.ncols):
                if "Ref plan type".casefold() in str(workSheet.cell(index1, index2).value).casefold().strip() or "Standard plan reference".casefold() in str(workSheet.cell(index1, index2).value).casefold().strip() :
                    rowIndex = index1
                    colIndex = index2
                if rowIndex != -1 and colIndex != -1:
                    flag = True
                    break
            if flag:
                break

        try:
            if str(workSheet.cell(rowIndex, colIndex + 1).value) in {"AEEV_IAEE07_0033", "02043_12_01665", "02043_12_01666"} and hasNumbers(str(workSheet.cell(rowIndex, colIndex + 2).value)):
                localisation = None
            else:
                localisation.append(("informations générales", rowIndex, colIndex + 1))
                localisation.append(("informations générales", rowIndex, colIndex + 2))
        except:
            localisation.append(("informations générales", rowIndex, colIndex + 1))
            localisation.append(("informations générales", rowIndex, colIndex + 2))
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0020(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if "Suppression"  in TSDApp.WorkbookStats.sheetNames or "suppression" in TSDApp.WorkbookStats.sheetNames:
        TSDApp.WorkbookStats.hasSupp = True
        try:
            index = TSDApp.WorkbookStats.sheetNames.index("suppression")
        except:
            index = TSDApp.WorkbookStats.sheetNames.index("Suppression")
        TSDApp.WorkbookStats.SuppIndex = index
        localisation = None
    else:
        TSDApp.WorkbookStats.hasSupp = False
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0025(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasSupp == False:
        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.SuppIndex)
        name = []
        name.append("Onglet/sheet")
        flag = False
        for index in range(0, workSheet.ncols):
            if str(workSheet.cell(TSDApp.suppressionHeaderRow, index).value).casefold().strip() == "sheet" or str(workSheet.cell(TSDApp.suppressionHeaderRow, index).value).casefold().strip() == "onglet":
                name = None
                flag = True
            if flag:
                break

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0030(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasSupp == False:
        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.SuppIndex)
        name = []
        name.append("Référence de la ligne/Line number")
        flag = False
        for index in range(0, workSheet.ncols):
            if str(workSheet.cell(TSDApp.suppressionHeaderRow, index).value).casefold().strip() == "référence de la ligne" or str(workSheet.cell(TSDApp.suppressionHeaderRow, index).value).casefold().strip() == "line number":
                name = None
                flag = True
            if flag:
                break

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0035(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasSupp == False:
        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.SuppIndex)
        name = []
        name.append("Version du TSD/Version of the document")
        flag = False
        for index in range(0, workSheet.ncols):
            if str(workSheet.cell(TSDApp.suppressionHeaderRow, index).value).casefold().strip() == "version du tsd" or str(workSheet.cell(TSDApp.suppressionHeaderRow, index).value).casefold().strip() == "version of the document":
                name = None
                flag = True
            if flag:
                break

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0040(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasSupp == False:
        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.SuppIndex)
        name = []
        name.append("Justification de la modification/Change reason")
        flag = False
        for index in range(0, workSheet.ncols):
            if str(workSheet.cell(TSDApp.suppressionHeaderRow, index).value).casefold().strip() == "justification de la modification" or str(
                    workSheet.cell(TSDApp.suppressionHeaderRow, index).value).casefold().strip() == "change reason":
                name = None
                flag = True
            if flag:
                break

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)


def Test_02043_18_04939_STRUCT_0046(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if "reference docs" in TSDApp.WorkbookStats.sheetNames:
        TSDApp.WorkbookStats.hasRefDocs = True
        index = TSDApp.WorkbookStats.sheetNames.index("reference docs")
        TSDApp.WorkbookStats.refDocsIndex = index
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], "", workBook, TSDApp)
    else:
        TSDApp.WorkbookStats.hasRefDocs = False
        name = []
        name.append("Missing Reference docs sheet")
        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)


def Test_02043_18_04939_STRUCT_0051(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    localisation1 = []
    if TSDApp.WorkbookStats.hasRefDocs == False:
        name = []
        name.append("Missing Reference docs sheet")
        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.refDocsIndex)
        localisation1 = []

        for index in range(0, workSheet.nrows):
                if str(workSheet.cell(index, 0).value).casefold().strip() in [ "vehicle architecture schematic", "planche d'architecture véhicule"]:
                    if str(workSheet.cell(index, 2).value).casefold().strip() == "":
                        localisation1.append(("reference docs", index, 2))
                        break

        if not localisation1:
            localisation1 = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation1, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0052(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    localisation1 = []
    if TSDApp.WorkbookStats.hasRefDocs == False:
        name = []
        name.append("Missing Reference docs sheet")
        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.refDocsIndex)

        for index in range(0, workSheet.nrows):
            if str(workSheet.cell(index, 0).value).casefold().strip() in ["diagnostic matrix", "matrice diag"]:
                if str(workSheet.cell(index, 2).value).casefold().strip() == "":
                    localisation1.append(("reference docs", index, 2))
                    break
        if not localisation1:
            localisation1 = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation1, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0053(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    localisation1 = []
    if TSDApp.WorkbookStats.hasRefDocs == False:
        name = []
        name.append("Missing Reference docs sheet")
        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.refDocsIndex)
        for index in range(0, workSheet.nrows):
            if str(workSheet.cell(index, 0).value).casefold().strip() in ["fault tree", "amdec"]:
                if str(workSheet.cell(index, 2).value).casefold().strip() == "":
                    localisation1.append(("reference docs", index, 2))
                    break
        if not localisation1:
            localisation1 = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation1, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0054(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    localisation1 = []
    if TSDApp.WorkbookStats.hasRefDocs == False:
        name = []
        name.append("Missing Reference docs sheet")
        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.refDocsIndex)
        for index in range(0, workSheet.nrows):
            if str(workSheet.cell(index, 0).value).casefold().strip() in ["ecu schematic", "synoptique ecu"]:
                if str(workSheet.cell(index, 2).value).casefold().strip() == "":
                    localisation1.append(("reference docs", index, 2))
                    break
        if not localisation1:
            localisation1 = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation1, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0055(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    localisation1 = []
    if TSDApp.WorkbookStats.hasRefDocs == False:
        name = []
        name.append("Missing Reference docs sheet")
        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.refDocsIndex)
        for index in range(0, workSheet.nrows):
            if str(workSheet.cell(index, 0).value).casefold().strip() in ["std"]:
                if str(workSheet.cell(index, 2).value).casefold().strip() == "":
                    localisation1.append(("reference docs", index, 2))
                    break
        if not localisation1:
            localisation1 = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation1, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0056(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    localisation1 = []
    if TSDApp.WorkbookStats.hasRefDocs == False:
        name = []
        name.append("Missing Reference docs sheet")
        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.refDocsIndex)
        for index in range(0, workSheet.nrows):
            if str(workSheet.cell(index, 0).value).casefold().strip() in ["complexity matrix (decli ee)"]:
                if str(workSheet.cell(index, 2).value).casefold().strip() == "":
                    localisation1.append(("reference docs", index, 2))
                    break
        if not localisation1:
            localisation1 = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation1, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0057(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    localisation1 = []
    if TSDApp.WorkbookStats.hasRefDocs == False:
        name = []
        name.append("Missing Reference docs sheet")
        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.refDocsIndex)
        for index in range(0, workSheet.nrows):
            if str(workSheet.cell(index, 0).value).casefold().strip() in ["décli"]:
                if str(workSheet.cell(index, 2).value).casefold().strip() == "":
                    localisation1.append(("reference docs", index, 2))
                    break
        if not localisation1:
            localisation1 = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation1, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0058(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    localisation1 = []
    if TSDApp.WorkbookStats.hasRefDocs == False:
        name = []
        name.append("Missing Reference docs sheet")
        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.refDocsIndex)
        for index in range(0, workSheet.nrows):
            if str(workSheet.cell(index, 0).value).casefold().strip() in ["dcee"]:
                if str(workSheet.cell(index, 2).value).casefold().strip() == "":
                    localisation1.append(("reference docs", index, 2))
                    break
        if not localisation1:
            localisation1 = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation1, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0059(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    localisation1 = []
    if TSDApp.WorkbookStats.hasRefDocs == False:
        name = []
        name.append("Missing Reference docs sheet")
        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.refDocsIndex)
        for index in range(0, workSheet.nrows):
            if str(workSheet.cell(index, 0).value).casefold().strip() in ["eead"]:
                if str(workSheet.cell(index, 2).value).casefold().strip() == "":
                    localisation1.append(("reference docs", index, 2))
                    break
        if not localisation1:
            localisation1 = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation1, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0060(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    localisation1 = []
    if TSDApp.WorkbookStats.hasRefDocs == False:
        name = []
        name.append("Missing Reference docs sheet")
        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.refDocsIndex)
        for index in range(0, workSheet.nrows):
            if str(workSheet.cell(index, 0).value).casefold().strip() in ["tfd"]:
                if str(workSheet.cell(index, 2).value).casefold().strip() == "":
                    localisation1.append(("reference docs", index, 2))
                    break

        if not localisation1:
            localisation1 = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation1, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0061(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    localisation1 = []
    if TSDApp.WorkbookStats.hasRefDocs == False:
        name = []
        name.append("Missing Reference docs sheet")
        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.refDocsIndex)
        for index in range(0, workSheet.nrows):
            if str(workSheet.cell(index, 0).value).casefold().strip() in ["sto"]:
                if str(workSheet.cell(index, 2).value).casefold().strip() == "":
                    localisation1.append(("reference docs", index, 2))
                    break
        if not localisation1:
            localisation1 = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation1, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0062(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    localisation1 = []
    if TSDApp.WorkbookStats.hasRefDocs == False:
        name = []
        name.append("Missing Reference docs sheet")
        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.refDocsIndex)
        for index in range(0, workSheet.nrows):
            if str(workSheet.cell(index, 0).value).casefold().strip() in ["view 5 and 8"]:
                if str(workSheet.cell(index, 2).value).casefold().strip() == "":
                    localisation1.append(("reference docs", index, 2))
                    break
        if not localisation1:
            localisation1 = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation1, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0063(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    localisation1 = []
    if TSDApp.WorkbookStats.hasRefDocs == False:
        name = []
        name.append("Missing Reference docs sheet")
        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.refDocsIndex)
        for index in range(0, workSheet.nrows):
            if str(workSheet.cell(index, 0).value).casefold().strip() in ["allocation matrix"]:
                if str(workSheet.cell(index, 2).value).casefold().strip() == "":
                    localisation1.append(("reference docs", index, 2))
                    break
        if not localisation1:
            localisation1 = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation1, workBook, TSDApp)

#[DOC3]

def Test_02043_18_04939_STRUCT_0100(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if  TSDApp.WorkbookStats.hasTable == True:
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0110(ExcelApp, workBook, TSDApp, DOC3Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        list_test = []

        for col in range(0,TSDApp.WorkbookStats.tableLastCol):
            dict = {}
            dict['1'] = workSheet.cell(TSDApp.tableHeaderRow - 1, col).value
            dict['2'] = workSheet.cell(TSDApp.tableHeaderRow, col).value
            dict['3'] = col
            dict['4'] = TSDApp.tableHeaderRow
            list_test.append(dict)


        DOC3 = xlrd.open_workbook(DOC3Name, on_demand=True)
        workSheetRef = DOC3.sheet_by_name("tableau")
        nrCols = workSheetRef.ncols
        list_ref = list()


        for col in range(0, nrCols):
            dict = {}
            dict['1'] = workSheetRef.cell(TSDApp.tableHeaderRow - 1, col).value
            dict['2'] = workSheetRef.cell(TSDApp.tableHeaderRow, col).value
            dict['3'] = col
            dict['4'] = TSDApp.tableHeaderRow
            list_ref.append(dict)

        name = list()

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                    if elem1['2'] == elem2['2']:
                        found = True
            if not found:
                name.append((elem1['2']))
        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0120(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasCode == True:
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0130(ExcelApp, workBook, TSDApp, DOC3Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        list_test = []


        for col in range(0,TSDApp.WorkbookStats.codeLastCol):
            dict = {}
            dict['2'] = workSheet.cell(TSDApp.codeHeaderRow, col).value
            dict['3'] = col
            dict['4'] = TSDApp.codeHeaderRow
            list_test.append(dict)


        DOC3 = xlrd.open_workbook(DOC3Name, on_demand=True)
        workSheetRef = DOC3.sheet_by_name("codes défauts")
        nrCols = workSheetRef.ncols
        list_ref = []

        for col in range(0, nrCols):
            dict = {}
            dict['2'] = workSheetRef.cell(TSDApp.codeHeaderRow, col).value
            dict['3'] = col
            dict['4'] = TSDApp.codeHeaderRow
            list_ref.append(dict)

        name = []

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['2'] == elem2['2']:
                    found = True
            if not found and elem1['2'] != "":
                name.append((elem1['2']))
        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0140(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if  TSDApp.WorkbookStats.hasMeasure == True:
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0150(ExcelApp, workBook, TSDApp, DOC3Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        list_test = []

        for col in range(0,TSDApp.WorkbookStats.measureLastCol):
            dict = {}
            dict['2'] = workSheet.cell(TSDApp.measureHeaderRow, col).value
            dict['3'] = col
            dict['4'] = TSDApp.measureHeaderRow
            list_test.append(dict)

        DOC3 = xlrd.open_workbook(DOC3Name, on_demand=True)
        workSheetRef = DOC3.sheet_by_name("mesures et commandes")
        nrCols = workSheetRef.ncols
        list_ref = []

        for col in range(0, nrCols):
            dict = {}
            dict['2'] = workSheetRef.cell(TSDApp.measureHeaderRow, col).value
            dict['3'] = col
            dict['4'] = TSDApp.measureHeaderRow
            list_ref.append(dict)

        name = []

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['2'] == elem2['2']:
                    found = True
            if not found and elem1['2'] != "":
                name.append((elem1['2']))
        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0160(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasDiagDeb == True:
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0170(ExcelApp, workBook, TSDApp, DOC3Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagDebIndex)
        list_test = []

        for col in range(0,TSDApp.WorkbookStats.DiagDebLastCol):
            dict = {}
            dict['2'] = workSheet.cell(TSDApp.diagDebHeaderRow, col).value
            dict['3'] = col
            dict['4'] = TSDApp.diagDebHeaderRow
            list_test.append(dict)

        DOC3 = xlrd.open_workbook(DOC3Name, on_demand=True)
        workSheetRef = DOC3.sheet_by_name("Diagnostic débarqués")
        nrCols = workSheetRef.ncols
        list_ref = []

        for col in range(0, nrCols):
            dict = {}
            dict['2'] = workSheetRef.cell(TSDApp.diagDebHeaderRow, col).value
            dict['3'] = col
            dict['4'] = TSDApp.diagDebHeaderRow
            list_ref.append(dict)

        name = []

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['2'] == elem2['2']:
                    found = True
            if not found:
                name.append((elem1['2']))
        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0180(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasEffClients == True:
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0190(ExcelApp, workBook, TSDApp, DOC3Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.EffClientsIndex)
        list_test = []

        for col in range(0,TSDApp.WorkbookStats.EffClientsLastCol):
            dict = {}
            dict['2'] = workSheet.cell(TSDApp.effClientsHeaderRow, col).value
            dict['3'] = col
            dict['4'] = TSDApp.effClientsHeaderRow
            list_test.append(dict)


        DOC3 = xlrd.open_workbook(DOC3Name, on_demand=True)
        workSheetRef = DOC3.sheet_by_name("Effets clients")
        nrCols = workSheetRef.ncols
        list_ref = []

        for col in range(0, nrCols):
            dict = {}
            dict['2'] = workSheetRef.cell(TSDApp.effClientsHeaderRow, col).value
            dict['3'] = col
            dict['4'] = TSDApp.effClientsHeaderRow
            list_ref.append(dict)

        name = []

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['2'] == elem2['2']:
                    found = True
            if not found:
                name.append((elem1['2']))
        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0200(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasER == True:
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag


def Test_02043_18_04939_STRUCT_0210(ExcelApp, workBook, TSDApp, DOC3Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasER == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ERIndex)
        nrCols = workSheet.ncols
        list_test = list()

        for col in range(0,nrCols):
            dict = {}
            dict['2'] = workSheet.cell(TSDApp.ERHeaderRow, col).value
            dict['3'] = col
            dict['4'] = TSDApp.ERHeaderRow
            list_test.append(dict)


        DOC3 = xlrd.open_workbook(DOC3Name, on_demand=True)
        workSheetRef = DOC3.sheet_by_name("ER")
        nrCols = workSheetRef.ncols
        list_ref = list()

        for col in range(0, nrCols):
            dict = {}
            dict['2'] = workSheetRef.cell(TSDApp.ERHeaderRow, col).value
            dict['3'] = col
            dict['4'] = TSDApp.ERHeaderRow
            list_ref.append(dict)

        name = list()

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['2'] == elem2['2']:
                    found = True
            if not found:
                name.append((elem1['2']))
        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0220(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasConstituants == True:
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0230(ExcelApp, workBook, TSDApp, DOC3Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasConstituants == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.constituantsIndex)
        nrCols = workSheet.ncols
        list_test = list()

        for col in range(0, nrCols):
            dict = {}
            dict['2'] = workSheet.cell(TSDApp.constituantsHeaderRow, col).value
            dict['3'] = col
            dict['4'] = TSDApp.constituantsHeaderRow
            list_test.append(dict)


        DOC3 = xlrd.open_workbook(DOC3Name, on_demand=True)
        workSheetRef = DOC3.sheet_by_name("Constituants")
        nrCols = workSheetRef.ncols
        list_ref = list()

        for col in range(0, nrCols):
            dict = {}
            dict['2'] = workSheetRef.cell(TSDApp.constituantsHeaderRow, col).value
            dict['3'] = col
            dict['4'] = TSDApp.constituantsHeaderRow
            list_ref.append(dict)

        name = list()

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['2'] == elem2['2']:
                    found = True
            if not found:
                name.append((elem1['2']))
        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0240(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasSitDeVie == True:
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0250(ExcelApp, workBook, TSDApp, DOC3Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasSitDeVie == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.SitDeVieIndex)
        nrCols = workSheet.ncols
        list_test = list()

        for col in range(0, nrCols):
            dict = {}
            dict['2'] = workSheet.cell(TSDApp.sitDeVieHeaderRow, col).value
            dict['3'] = col
            dict['4'] = TSDApp.sitDeVieHeaderRow
            list_test.append(dict)


        DOC3 = xlrd.open_workbook(DOC3Name, on_demand=True)
        workSheetRef = DOC3.sheet_by_name("situations de vie")
        nrCols = workSheetRef.ncols
        list_ref = list()

        for col in range(0, nrCols):
            dict = {}
            dict['2'] = workSheetRef.cell(TSDApp.sitDeVieHeaderRow, col).value
            dict['3'] = col
            dict['4'] = TSDApp.sitDeVieHeaderRow
            list_ref.append(dict)

        name = list()

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['2'] == elem2['2']:
                    found = True
            if not found:
                name.append((elem1['2']))
        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0260(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasMDD == True:
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0270(ExcelApp, workBook, TSDApp, DOC3Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasMDD == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.MDDIndex)
        list_test = []

        for col in range(0,TSDApp.WorkbookStats.MDDLastCol):
            dict = {}
            dict['2'] = workSheet.cell(TSDApp.listeMDDHeaderRow, col).value
            dict['3'] = col
            dict['4'] = TSDApp.listeMDDHeaderRow
            list_test.append(dict)


        DOC3 = xlrd.open_workbook(DOC3Name, on_demand=True)
        workSheetRef = DOC3.sheet_by_name("Liste MDD")
        nrCols = workSheetRef.ncols
        list_ref = []

        for col in range(0, nrCols):
            dict = {}
            dict['2'] = workSheetRef.cell(TSDApp.listeMDDHeaderRow, col).value
            dict['3'] = col
            dict['4'] = TSDApp.listeMDDHeaderRow
            list_ref.append(dict)

        name = []

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['2'] == elem2['2']:
                    found = True
            if not found:
                name.append((elem1['2']))

        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)


#DOC4

def Test_02043_18_04939_STRUCT_0400(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasTable == True:
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0410(ExcelApp, workBook, TSDApp, DOC4Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        list_test = []

        for col in range(0,TSDApp.WorkbookStats.tableLastCol):
            dict = {}
            dict['1'] = workSheet.cell(TSDApp.tableHeaderRow - 2, col).value
            dict['2'] = workSheet.cell(TSDApp.tableHeaderRow - 1, col).value
            dict['3'] = workSheet.cell(TSDApp.tableHeaderRow, col).value
            dict['4'] = col
            dict['5'] = TSDApp.tableHeaderRow
            list_test.append(dict)

        DOC4 = xlrd.open_workbook(DOC4Name, on_demand=True)
        try:
            workSheetRef = DOC4.sheet_by_name("tableau")
        except:
            workSheetRef = DOC4.sheet_by_name("Table")

        nrCols = workSheetRef.ncols
        list_ref = []

        for col in range(0, nrCols):
            dict = {}
            dict['1'] = workSheetRef.cell(TSDApp.tableHeaderRow - 2, col).value
            dict['2'] = workSheetRef.cell(TSDApp.tableHeaderRow - 1, col).value
            dict['3'] = workSheetRef.cell(TSDApp.tableHeaderRow, col).value
            dict['4'] = col
            dict['5'] = TSDApp.tableHeaderRow
            list_ref.append(dict)

        name = list()

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                    if elem1['1'] == elem2['1'] and elem1['2'] == elem2['2'] and elem1['3'] == elem2['3']:
                        found = True
            if not found:
                name.append((elem1['3']))

        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0420(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasDiagNeeds == True:
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0430(ExcelApp, workBook, TSDApp, DOC4Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagNeedsIndex)
        nrCols = workSheet.ncols
        list_test = list()


        for col in range(0,nrCols):
            dict = {}
            dict['1'] = workSheet.cell(TSDApp.diagNeedsHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.diagNeedsHeaderRow
            list_test.append(dict)


        DOC4 = xlrd.open_workbook(DOC4Name, on_demand=True)
        workSheetRef = DOC4.sheet_by_name("Diagnostic Needs")
        nrCols = workSheetRef.ncols

        list_ref = list()
        for col in range(0, nrCols):
            dict = {}
            dict['1'] = workSheetRef.cell(TSDApp.diagNeedsHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.diagNeedsHeaderRow
            list_ref.append(dict)

        name = list()

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['1'] == elem2['1'] :
                    found = True
            if not found:
                name.append((elem1['1']))

        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0440(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasCustEff == True:
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0450(ExcelApp, workBook, TSDApp, DOC4Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.EffClientsIndex)
        list_test = []

        for col in range(0, TSDApp.WorkbookStats.EffClientsLastCol):
            dict = {}
            dict['1'] = workSheet.cell(TSDApp.effClientsHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.effClientsHeaderRow
            list_test.append(dict)


        DOC4 = xlrd.open_workbook(DOC4Name, on_demand=True)
        workSheetRef = DOC4.sheet_by_name("Customer Effects")
        nrCols = workSheetRef.ncols
        list_ref = list()

        for col in range(0, nrCols):
            dict = {}
            dict['1'] = workSheetRef.cell(TSDApp.effClientsHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.effClientsHeaderRow
            list_ref.append(dict)

        name = list()

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['1'] == elem2['1']:
                    found = True
            if not found:
                name.append((elem1['1']))

        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0460(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasFearedEvent == True:
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0470(ExcelApp, workBook, TSDApp, DOC4Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.FearedEventIndex)
        nrCols = workSheet.ncols
        list_test = list()

        for col in range(0, nrCols):
            dict = {}
            dict['1'] = workSheet.cell(TSDApp.fearedEventHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.fearedEventHeaderRow
            list_test.append(dict)

        DOC4 = xlrd.open_workbook(DOC4Name, on_demand=True)
        try:
            workSheetRef = DOC4.sheet_by_name("feared events")
        except:
            workSheetRef = DOC4.sheet_by_name("ER")

        nrCols = workSheetRef.ncols
        list_ref = list()

        for col in range(0, nrCols):
            dict = {}
            dict['1'] = workSheetRef.cell(TSDApp.fearedEventHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.fearedEventHeaderRow
            list_ref.append(dict)

        name = list()

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['1'] == elem2['1']:
                    found = True
            if not found:
                name.append((elem1['1']))

        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0480(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasSystem == True:
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0490(ExcelApp, workBook, TSDApp, DOC4Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasSystem == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.SystemIndex)
        nrCols = workSheet.ncols
        list_test = list()

        for col in range(0, nrCols):
            dict = {}
            dict['1'] = workSheet.cell(TSDApp.systemHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.systemHeaderRow
            list_test.append(dict)


        DOC4 = xlrd.open_workbook(DOC4Name, on_demand=True)
        try:
            workSheetRef = DOC4.sheet_by_name("System")
        except:
            workSheetRef = DOC4.sheet_by_name("Système")

        nrCols = workSheetRef.ncols
        list_ref = list()

        for col in range(0, nrCols):
            dict = {}
            dict['1'] = workSheetRef.cell(TSDApp.systemHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.systemHeaderRow
            list_ref.append(dict)

        name = list()

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['1'] == elem2['1']:
                    found = True
            if not found:
                name.append((elem1['1']))

        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0500(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasOpSit == True:
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0510(ExcelApp, workBook, TSDApp, DOC4Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasOpSit == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.OpSitIndex)
        nrCols = workSheet.ncols
        list_test = list()

        for col in range(0, nrCols):
            dict = {}
            dict['1'] = workSheet.cell(TSDApp.opSitHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.opSitHeaderRow
            list_test.append(dict)


        DOC4 = xlrd.open_workbook(DOC4Name, on_demand=True)
        workSheetRef = DOC4.sheet_by_name("Operation situation")
        nrCols = workSheetRef.ncols
        list_ref = list()

        for col in range(0, nrCols):
            dict = {}
            dict['1'] = workSheetRef.cell(TSDApp.opSitHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.opSitHeaderRow
            list_ref.append(dict)

        name = list()

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['1'] == elem2['1']:
                    found = True
            if not found:
                name.append((elem1['1']))
        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0520(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasReqTech == True:
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0530(ExcelApp, workBook, TSDApp, DOC4Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasReqTech == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ReqTechIndex)
        list_test = []


        for col in range(0, TSDApp.WorkbookStats.ReqTechLastCol):
            dict = {}
            dict['1'] = workSheet.cell(TSDApp.reqTechHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.reqTechHeaderRow
            list_test.append(dict)


        DOC4 = xlrd.open_workbook(DOC4Name)
        workSheetRef = DOC4.sheet_by_name("Req. of tech. effects")
        nrCols = workSheetRef.ncols
        list_ref = []

        for col in range(0, nrCols):
            dict = {}
            dict['1'] = workSheetRef.cell(TSDApp.reqTechHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.reqTechHeaderRow
            list_ref.append(dict)

        name = list()

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['1'] == elem2['1']:
                    found = True
            if not found:
                name.append((elem1['1']))
        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)


#DOC5


def Test_02043_18_04939_STRUCT_0700(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasTable == True:
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0710(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        list_test = []

        for col in range(1,TSDApp.WorkbookStats.tableLastCol):
            dict = {}
            dict['1'] = workSheet.cell(TSDApp.tableHeaderRow - 1, col).value
            dict['2'] = workSheet.cell(TSDApp.tableHeaderRow, col).value
            dict['3'] = col
            dict['4'] = TSDApp.tableHeaderRow
            list_test.append(dict)

        DOC5 = xlrd.open_workbook(DOC5Name, on_demand=True)

        try:
            workSheetRef = DOC5.sheet_by_name("Tableau")
        except:
            workSheetRef = DOC5.sheet_by_name("Table")

        nrCols = workSheetRef.ncols
        list_ref = []

        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.cell(TSDApp.tableHeaderRow - 1, col).value
            dict['2'] = workSheetRef.cell(TSDApp.tableHeaderRow, col).value
            dict['3'] = col
            dict['4'] = TSDApp.tableHeaderRow
            list_ref.append(dict)

        name = []

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['1'] == elem2['1'] and elem1['2'] == elem2['2']:
                    found = True
            if not found:
                name.append(workSheetRef.cell(elem1['4'], elem1['3']).value)

        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0720(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasDataCodes == True or TSDApp.WorkbookStats.hasCode == True:
        if TSDApp.WorkbookStats.hasCode == True:
            TSDApp.WorkbookStats.hasDataCodes = TSDApp.WorkbookStats.hasCode
            TSDApp.WorkbookStats.DataCodesIndex = TSDApp.WorkbookStats.codeIndex
            TSDApp.WorkbookStats.DataCodesLastRow = TSDApp.WorkbookStats.codeLastRow
            TSDApp.WorkbookStats.DataCodesLastCol = TSDApp.WorkbookStats.codeLastCol
        else:
            pass
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0730(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DataCodesIndex)
        list_test = []

        for col in range(1,TSDApp.WorkbookStats.DataCodesLastCol):
            dict = {}
            dict['1'] = workSheet.cell(TSDApp.dataCodesHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.dataCodesHeaderRow
            list_test.append(dict)

        DOC5 = xlrd.open_workbook(DOC5Name, on_demand=True)
        try:
            workSheetRef = DOC5.sheet_by_name("Data trouble codes")
        except:
            workSheetRef = DOC5.sheet_by_name("codes défauts")

        nrCols = workSheetRef.ncols
        list_ref = []

        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.cell(TSDApp.dataCodesHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.dataCodesHeaderRow
            list_ref.append(dict)

        name = []

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['1'] == elem2['1']:
                    found = True
            if not found:
                name.append(workSheetRef.cell(elem1['3'], elem1['2']).value)

        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0740(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasReadDataIO == True or TSDApp.WorkbookStats.hasMeasure == True:
        if TSDApp.WorkbookStats.hasMeasure == True:
            TSDApp.WorkbookStats.hasReadDataIO = TSDApp.WorkbookStats.hasMeasure
            TSDApp.WorkbookStats.ReadDataIOIndex = TSDApp.WorkbookStats.measureIndex
            TSDApp.WorkbookStats.ReadDataIOLastRow = TSDApp.WorkbookStats.measureLastRow
            TSDApp.WorkbookStats.ReadDataIOLastCol = TSDApp.WorkbookStats.measureLastCol
        else:
            pass
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0750(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ReadDataIOIndex)
        list_test = []

        for col in range(1, TSDApp.WorkbookStats.ReadDataIOLastCol):
            dict = {}
            dict['1'] = workSheet.cell(TSDApp.readDataIOHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.readDataIOHeaderRow
            list_test.append(dict)

        DOC5 = xlrd.open_workbook(DOC5Name, on_demand=True)
        try:
            workSheetRef = DOC5.sheet_by_name("Read data and IO control")
        except:
            workSheetRef = DOC5.sheet_by_name("mesures et commandes")

        nrCols = workSheetRef.ncols
        list_ref = []

        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.cell(TSDApp.readDataIOHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.readDataIOHeaderRow
            list_ref.append(dict)

        name = []

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['1'] == elem2['1']:
                    found = True
            if not found:
                name.append(workSheetRef.cell(elem1['3'], elem1['2']).value)

        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0760(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasNotEmbDiag == True or TSDApp.WorkbookStats.hasReadDataIO == True:
        if TSDApp.WorkbookStats.hasReadDataIO == True:
            TSDApp.WorkbookStats.hasNotEmbDiag = TSDApp.WorkbookStats.hasReadDataIO
            TSDApp.WorkbookStats.NotEmbDiagIndex = TSDApp.WorkbookStats.ReadDataIOIndex
            TSDApp.WorkbookStats.NotEmbDiagLastRow = TSDApp.WorkbookStats.ReadDataIOLastRow
            TSDApp.WorkbookStats.NotEmbDiagLastCol = TSDApp.WorkbookStats.ReadDataIOLastCol
        else:
            pass
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0770(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasNotEmbDiag == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.NotEmbDiagIndex)
        list_test = []

        for col in range(1,TSDApp.WorkbookStats.NotEmbDiagLastCol):
            dict = {}
            dict['1'] = workSheet.cell(TSDApp.notEmbDiagHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.notEmbDiagHeaderRow
            list_test.append(dict)

        DOC5 = xlrd.open_workbook(DOC5Name, on_demand=True)
        try:
            workSheetRef = DOC5.sheet_by_name("Not embedded diagnosis")
        except:
            workSheetRef = DOC5.sheet_by_name("Read data and IO control")

        nrCols = workSheetRef.ncols

        list_ref = []
        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.cell(TSDApp.notEmbDiagHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.notEmbDiagHeaderRow
            list_ref.append(dict)

        name = []

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['1'] == elem2['1']:
                    found = True
            if not found:
                name.append(workSheetRef.cell(elem1['3'], elem1['2']).value)

        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0780(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = True
    if TSDApp.WorkbookStats.hasEffClients == True:
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0790(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasCustEff == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.EffClientsLastCol)
        list_test = []
        for col in range(1, TSDApp.WorkbookStats.EffClientsLastCol):
            dict = {}
            dict['1'] = workSheet.cell(TSDApp.effClientsHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.effClientsHeaderRow
            list_test.append(dict)


        DOC5 = xlrd.open_workbook(DOC5Name, on_demand=True)
        try:
            workSheetRef = DOC5.sheet_by_name("Customer effect")
        except:
            workSheetRef = DOC5.sheet_by_name("Effets clients")

        nrCols = workSheetRef.ncols
        list_ref = []

        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.cell(TSDApp.effClientsHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.effClientsHeaderRow
            list_ref.append(dict)

        name = list()

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['1'] == elem2['1']:
                    found = True
            if not found:
                name.append(workSheetRef.cell(elem1['3'], elem1['2']).value)

        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0800(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasFearedEvent == True or TSDApp.WorkbookStats.hasER == True:
        if TSDApp.WorkbookStats.hasER == True:
            TSDApp.WorkbookStats.hasFearedEvent = TSDApp.WorkbookStats.hasER
            TSDApp.WorkbookStats.FearedEventIndex = TSDApp.WorkbookStats.ERIndex
            TSDApp.WorkbookStats.FearedEventLastRow = TSDApp.WorkbookStats.ERLastRow
            TSDApp.WorkbookStats.FearedEventLastCol = TSDApp.WorkbookStats.ERLastCol
        else:
            pass
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0810(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.FearedEventIndex)
        nrCols = workSheet.ncols
        list_test = list()

        for col in range(1,nrCols):
            dict = {}
            dict['1'] = workSheet.cell(TSDApp.fearedEventHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.fearedEventHeaderRow
            list_test.append(dict)


        DOC5 = xlrd.open_workbook(DOC5Name, on_demand=True)
        try:
            workSheetRef = DOC5.sheet_by_name("Feared events")
        except:
            workSheetRef = DOC5.sheet_by_name("ER")

        nrCols = workSheetRef.ncols
        list_ref = list()

        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.cell(TSDApp.fearedEventHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.fearedEventHeaderRow
            list_ref.append(dict)

        name = list()

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['1'] == elem2['1']:
                    found = True
            if not found:
                name.append(workSheetRef.cell(elem1['3'], elem1['2']).value)

        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0820(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasParts == True or TSDApp.WorkbookStats.hasConstituants == True:
        if TSDApp.WorkbookStats.hasConstituants == True:
            TSDApp.WorkbookStats.hasParts = TSDApp.WorkbookStats.hasConstituants
            TSDApp.WorkbookStats.PartsIndex = TSDApp.WorkbookStats.constituantsIndex
            TSDApp.WorkbookStats.PartsLastRow = TSDApp.WorkbookStats.constituantsLastRow
            TSDApp.WorkbookStats.PartsLastCol = TSDApp.WorkbookStats.constituantsLastCol
        else:
            pass
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag


def Test_02043_18_04939_STRUCT_0830(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasConstituants == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.constituantsIndex)
        nrCols = workSheet.ncols

        list_test = list()

        for col in range(1,nrCols):
            dict = {}
            dict['1'] = workSheet.cell(TSDApp.partsHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.partsHeaderRow
            list_test.append(dict)

        DOC5 = xlrd.open_workbook(DOC5Name, on_demand=True)
        try:
            workSheetRef = DOC5.sheet_by_name("Constituants")
        except:
            workSheetRef = DOC5.sheet_by_name("Parts")

        nrCols = workSheetRef.ncols
        list_ref = list()

        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.cell(TSDApp.partsHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.partsHeaderRow
            list_ref.append(dict)

        name = []

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['1'] == elem2['1']:
                    found = True
            if not found:
                name.append(workSheetRef.cell(elem1['3'], elem1['2']).value)

        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0840(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasSituation == True or TSDApp.WorkbookStats.hasSitDeVie == True:
        if TSDApp.WorkbookStats.hasSitDeVie == True:
            TSDApp.WorkbookStats.hasSituation = TSDApp.WorkbookStats.hasSitDeVie
            TSDApp.WorkbookStats.SituationIndex = TSDApp.WorkbookStats.SitDeVieLastIndex
            TSDApp.WorkbookStats.SituationLastRow = TSDApp.WorkbookStats.SitDeVieLastRow
            TSDApp.WorkbookStats.SituationLastCol = TSDApp.WorkbookStats.SitDeVieLastCol
        else:
            pass
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0850(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasSitDeVie == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.SitDeVieIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count

        list_test = []

        for col in range(1,nrCols):
            dict = {}
            dict['1'] = workSheet.cell(TSDApp.situationHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.situationHeaderRow
            list_test.append(dict)

        DOC5 = xlrd.open_workbook( DOC5Name, on_demand=True)
        try:
            workSheetRef = DOC5.sheet_by_name("Situation")
        except:
            workSheetRef = DOC5.sheet_by_name("situation de vie")

        nrCols = workSheetRef.ncols
        list_ref = list()

        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.cell(TSDApp.situationHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.situationHeaderRow
            list_ref.append(dict)

        name = list()

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['1'] == elem2['1']:
                    found = True
            if not found:
                name.append(workSheetRef.cell(elem1['3'], elem1['2']).value)

        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0860(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasDegradedMode == True or TSDApp.WorkbookStats.hasMDD == True:
        if TSDApp.WorkbookStats.hasMDD == True:
            TSDApp.WorkbookStats.hasDegradedMode = TSDApp.WorkbookStats.hasMDD
            TSDApp.WorkbookStats.DegradedModeIndex = TSDApp.WorkbookStats.MDDIndex
            TSDApp.WorkbookStats.DegradedModeLastRow = TSDApp.WorkbookStats.MDDLastRow
            TSDApp.WorkbookStats.DegradedModeLastCol = TSDApp.WorkbookStats.MDDLastCol
        else:
            pass
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0870(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasMDD == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.MDDIndex)
        list_test = []

        for col in range(1,TSDApp.WorkbookStats.DegradedModeLastCol):
            dict = {}
            dict['1'] = workSheet.cell(TSDApp.degradedModeHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.degradedModeHeaderRow
            list_test.append(dict)


        DOC5 = xlrd.open_workbook(DOC5Name, on_demand=True)
        try:
            workSheetRef = DOC5.sheet_by_name("Degraded mode")
        except:
            workSheetRef = DOC5.sheet_by_name("liste MDD")

        nrCols = workSheetRef.ncols
        list_ref = []

        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.cell(TSDApp.degradedModeHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.degradedModeHeaderRow
            list_ref.append(dict)

        name = list()

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['1'] == elem2['1']:
                    found = True
            if not found:
                name.append(workSheetRef.cell(elem1['3'], elem1['2']).value)

        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0880(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasTechEff == True:
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0890(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTechEff == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.TechEffIndex)
        nrCols = workSheet.ncols

        list_test = list()

        for col in range(1,nrCols):
            dict = {}
            dict['1'] = workSheet.cell(TSDApp.techEffHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.techEffHeaderRow
            list_test.append(dict)


        DOC5 = xlrd.open_workbook(DOC5Name, on_demand=True)
        try:
            workSheetRef = DOC5.sheet_by_name("Technical effect")
        except:
            workSheetRef = DOC5.sheet_by_name("Effets techniques")

        nrCols = workSheetRef.ncols
        list_ref = list()

        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.cell(TSDApp.techEffHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.techEffHeaderRow
            list_ref.append(dict)

        name = list()

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['1'] == elem2['1']:
                    found = True
            if not found:
                name.append(workSheetRef.cell(elem1['3'], elem1['2']).value)

        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)

def Test_02043_18_04939_STRUCT_0900(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    flag = False
    if TSDApp.WorkbookStats.hasVariant == True:
        localisation = None
    else:
        localisation = ""
        flag = True
    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return flag

def Test_02043_18_04939_STRUCT_0910(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasVariant == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.VariantIndex)
        nrCols = workSheet.ncols

        list_test = list()

        for col in range(1,nrCols):
            dict = {}
            dict['1'] = workSheet.cell(TSDApp.variantHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.variantHeaderRow
            list_test.append(dict)


        DOC5 = xlrd.open_workbook(DOC5Name, on_demand=True)
        try:
            workSheetRef = DOC5.sheet_by_name("Variant")
        except:
            workSheetRef = DOC5.sheet_by_name("Variantes")

        nrCols = workSheetRef.ncols
        list_ref = list()

        for col in range(1, nrCols):
            dict = {}
            dict['1'] = workSheetRef.cell(TSDApp.variantHeaderRow, col).value
            dict['2'] = col
            dict['3'] = TSDApp.variantHeaderRow
            list_ref.append(dict)

        name = list()

        for elem1 in list_ref:
            found = False
            for elem2 in list_test:
                if elem1['1'] == elem2['1']:
                    found = True
            if not found:
                name.append(workSheetRef.cell(elem1['3'], elem1['2']).value)

        if not name:
            name = None

        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)