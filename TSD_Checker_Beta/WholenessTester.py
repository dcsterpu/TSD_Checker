import TSD_Checker_V0_5_2
import inspect
from ExcelEdit import TestReturn as result
from ErrorMessages import errorMessagesDict as error

def Test_02043_18_04939_WHOLENESS_1000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Référence":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
        refCellRange = workSheet.Cells(refRowIndex,refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()
        firtCell = workSheet.Cells(refRowIndex + nrLines, 1)
        lastCell = workSheet.Cells(workSheetRange.Rows.Count, nrCols)
        workSheetRange = workSheet.Range(firtCell, lastCell)
        flag = False
        for row in workSheetRange.Rows:
            flag = False
            for valueTuple in row.Value:
                for value in valueTuple:
                    if value != None:
                        flag = True
            if flag == False:
                TSDApp.WorkbookStats.tableLastRow = row.Row
                break

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow ):
            if workSheet.Cells(index, refColIndex).Value == None:
                localisation.append(workSheet.Cells(index, refColIndex))
        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)


def Test_02043_18_04939_WHOLENESS_1001(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Version":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                localisation.append(workSheet.Cells(index, refColIndex))
        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)


def Test_02043_18_04939_WHOLENESS_1010(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Référence":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()
        firtCell = workSheet.Cells(refRowIndex + nrLines, 1)
        lastCell = workSheet.Cells(workSheetRange.Rows.Count, nrCols)
        workSheetRange = workSheet.Range(firtCell, lastCell)
        flag = False
        for row in workSheetRange.Rows:
            flag = False
            for valueTuple in row.Value:
                for value in valueTuple:
                    if value != None:
                        flag = True
            if flag == False:
                TSDApp.WorkbookStats.codeLastRow = row.Row
                break

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                localisation.append(workSheet.Cells(index, refColIndex))
        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_WHOLENESS_1011(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Version":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                localisation.append(workSheet.Cells(index, refColIndex))
        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_WHOLENESS_1020(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Référence":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()
        firtCell = workSheet.Cells(refRowIndex + nrLines, 1)
        lastCell = workSheet.Cells(workSheetRange.Rows.Count, nrCols)
        workSheetRange = workSheet.Range(firtCell, lastCell)
        flag = False
        for row in workSheetRange.Rows:
            flag = False
            for valueTuple in row.Value:
                for value in valueTuple:
                    if value != None:
                        flag = True
            if flag == False:
                TSDApp.WorkbookStats.measureLastRow = row.Row
                break

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.measureLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                localisation.append(workSheet.Cells(index, refColIndex))
        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_WHOLENESS_1021(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Version":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.measureLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                localisation.append(workSheet.Cells(index, refColIndex))
        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_WHOLENESS_1080(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Code défaut":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                localisation.append(workSheet.Cells(index, refColIndex))
        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_WHOLENESS_1090(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "supporté par constituant (s)":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                localisation.append(workSheet.Cells(index, refColIndex))
        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_WHOLENESS_1100(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "supporté par constituant (s)":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.measureLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                localisation.append(workSheet.Cells(index, refColIndex))
        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_WHOLENESS_1110(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "libellé (signification)":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                localisation.append(workSheet.Cells(index, refColIndex))
        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_WHOLENESS_1120(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Description de la strategie pour détecter le défaut":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                localisation.append(workSheet.Cells(index, refColIndex))
        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)


'''def Test_02043_18_04939_WHOLENESS_1130(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Seuil de détection  /  valeur  du défaut":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                localisation.append(workSheet.Cells(index, refColIndex))
        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)'''

def Test_02043_18_04939_WHOLENESS_1140(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Temps de confirmation du défaut":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                localisation.append(workSheet.Cells(index, refColIndex))
        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_WHOLENESS_1150(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Description de la strategie de disparition du défaut / Procedure à effectuer pour vérifier la disparition du défaut":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                localisation.append(workSheet.Cells(index, refColIndex))
        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_WHOLENESS_1160(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Mode dégradé":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                localisation.append(workSheet.Cells(index, refColIndex))
        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_WHOLENESS_1170(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Voyant":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                localisation.append(workSheet.Cells(index, refColIndex))
        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)


def Test_02043_18_04939_WHOLENESS_1180(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Constituant défaillant détecté":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                localisation.append(workSheet.Cells(index, refColIndex))
        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_WHOLENESS_1190(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Défaillance constituant":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                localisation.append(workSheet.Cells(index, refColIndex))
        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)


def Test_02043_18_04939_WHOLENESS_1200(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Situation de vie client":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                localisation.append(workSheet.Cells(index, refColIndex))
        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)


def Test_02043_18_04939_WHOLENESS_1210(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Effet(s) client(s)":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                localisation.append(workSheet.Cells(index, refColIndex))
        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)


def Test_02043_18_04939_WHOLENESS_1220(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Code défaut":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                localisation.append(workSheet.Cells(index, refColIndex))
        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)


def Test_02043_18_04939_WHOLENESS_1230(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Défaillance constituant":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                localisation.append(workSheet.Cells(index, refColIndex))
        if str(localisation) == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)