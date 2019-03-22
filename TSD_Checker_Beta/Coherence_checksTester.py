import TSD_Checker_V0_5_2
import inspect
from ExcelEdit import TestReturn as result
from ErrorMessages import errorMessagesDict as error


#Coherence checks requirements

def Test_02043_18_04939_COH_2000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0
        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "mesures et commandes (Mesure Parametre et Test Actionneur) / Tests de cohérence":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        nrCols = refCellRange.Columns.Count
        localisation = list()
        list_table = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
            if workSheet.Cells(index, refColIndex).Value == "N/A" or workSheet.Cells(index, refColIndex).Value == None:
                pass
            else:
                list_table.append(workSheet.Cells(index, refColIndex).Value)

    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        measureColIndex = 0
        var = 0
        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "libellé (signification)":
                    measureColIndex = cellRow.index(cell) + 1
                    measureRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if measureColIndex != 0:
                break

        measureCellRange = workSheet.Cells(measureRowIndex, measureColIndex).MergeArea
        nrLines = measureCellRange.Rows.Count
        nrCols = measureCellRange.Columns.Count
        localisation = list()
        list_measure = list()

        for index in range(measureRowIndex + nrLines, TSDApp.WorkbookStats.measureLastRow):
            if workSheet.Cells(index, measureColIndex).Value == None:
                pass
            else:
                list_measure.append(workSheet.Cells(index, measureColIndex).Value)

    for element in list_table:
        if element in list_measure:
            localisation = None
        else:
            localisation = ""
            break
    if list_table == "[]":
        localisation = None

    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_COH_2005(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        codeColIndex = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Code défaut":
                    codeColIndex = cellRow.index(cell) + 1
                    codeRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if codeColIndex != 0:
                break

        codeCellRange = workSheet.Cells(codeRowIndex, codeColIndex).MergeArea
        nrLines = codeCellRange.Rows.Count
        localisation = list()
        listValues = list()
        firstCell = workSheet.Cells(codeRowIndex + nrLines, 1)
        lastCell = workSheet.Cells(workSheetRange.Rows.Count, nrCols)
        workSheetRange = workSheet.Range(firstCell, lastCell)
        flag = False
        ok = 1

        for row in workSheetRange.Rows:
            flag = False
            for valueTuple in row.Value:
                for value in valueTuple:
                    if value != None:
                        flag = True
            if flag == False:
                TSDApp.WorkbookStats.codeLastRow = row.Row
                break

        for index in range(codeRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
            if workSheet.Cells(index, codeColIndex).Value.count('-') != 2:
                localisation.append(workSheet.Cells(index, codeColIndex))

            else:
                listValues = workSheet.Cells(index, codeColIndex).Value.split('-')
                if not listValues[0].isascii():
                    ok = 0
                if not listValues[1][0].isalpha():
                    ok = 0
                try:
                    int(listValues[1][1:], 16)
                except:
                    ok = 0
                try:
                    int(listValues[2], 16)
                except:
                    ok = 0
                if ok == 1:
                    tempDict = dict()
                    tempDict["value"] = listValues[0]
                    tempDict["codenr"] = listValues[1]
                    tempDict["localisation"] = workSheet.Cells(index, codeColIndex)
                    TSDApp.WorkbookStats.famillyList.append(dict(tempDict))
                else:
                    localisation.append(workSheet.Cells(index, codeColIndex))
                    ok = 1
                    break

        if localisation == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_COH_2006(ExcelApp, workBook, TSDApp, DOC8Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.famillyList == "[]":
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:

        DOC8 = ExcelApp.Workbooks.Open(DOC8Name)
        workSheetRef = DOC8.Sheets("sous familles Cesare 2018 08 30")

        workSheetRange = workSheetRef.UsedRange
        nrCols = workSheetRange.Columns.Count
        nrRows = workSheetRange.Rows.Count
        refColIndex = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == " Nom de la sous famille ":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheetRef.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()
        flag = False
        list_ref  =list()


        for index in range(refRowIndex + nrLines, nrRows + 1):
            if workSheetRef.Cells(index, refColIndex).Value == None:
                pass
            else:
                list_ref.append(workSheetRef.Cells(index, refColIndex).Value)

        for element in TSDApp.WorkbookStats.famillyList:
            if element["value"] in list_ref:
                pass
            else:
               localisation.append(element["localisation"])


        if localisation == "[]":
            localisation = None

    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_COH_2007(ExcelApp, workBook, TSDApp, DOC14Name):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.famillyList == "[]":
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:

        DOC14 = ExcelApp.Workbooks.Open(DOC14Name)
        workSheetRef = DOC14.Sheets("Matrix")

        workSheetRange = workSheetRef.UsedRange
        nrCols = workSheetRange.Columns.Count
        nrRows = workSheetRange.Rows.Count
        refColIndex = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Data Trouble Code (DTC)":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheetRef.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        localisation = list()
        flag = False
        list_ref  =list()


        for index in range(refRowIndex + nrLines, nrRows + 1):
            if workSheetRef.Cells(index, refColIndex).Value == None:
                pass
            else:
                list_ref.append(workSheetRef.Cells(index, refColIndex).Value)

        for element in TSDApp.WorkbookStats.famillyList:
            if element["codenr"] in list_ref:
                pass
            else:
               localisation.append(element["localisation"])


        if localisation == "[]":
            localisation = None

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_COH_2010(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0
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
        nrCols = refCellRange.Columns.Count
        localisation = list()
        list_table = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
            if workSheet.Cells(index, refColIndex).Value == "No DTC" or workSheet.Cells(index, refColIndex).Value == None:
                pass
            else:
                list_table.append(workSheet.Cells(index, refColIndex).Value)

    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        codeColIndex = 0
        var = 0
        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Code défaut":
                    codeColIndex = cellRow.index(cell) + 1
                    codeRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if codeColIndex != 0:
                break

        codeCellRange = workSheet.Cells(codeRowIndex, codeColIndex).MergeArea
        nrLines = codeCellRange.Rows.Count
        nrCols = codeCellRange.Columns.Count
        localisation = list()
        list_code = list()

        for index in range(codeRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
            if workSheet.Cells(index, codeColIndex).Value == None:
                pass
            else:
                list_code.append(workSheet.Cells(index, codeColIndex).Value)

    for element in list_table:
        if element in list_code:
            localisation = None
        else:
            localisation = ""
            break
    if list_table == "[]":
        localisation = None

    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_COH_2020(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0
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
        nrCols = refCellRange.Columns.Count
        localisation = list()
        list_table = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                pass
            else:
                list_table.append(workSheet.Cells(index, refColIndex).Value)

    if TSDApp.WorkbookStats.hasConstituants == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.constituantsIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        constituantsColIndex = 0
        var = 0
        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Noms":
                    constituantsColIndex = cellRow.index(cell) + 1
                    constituantsRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if constituantsColIndex != 0:
                break

        constituantsCellRange = workSheet.Cells(constituantsRowIndex, constituantsColIndex).MergeArea
        nrLines = constituantsCellRange.Rows.Count
        localisation = list()
        firstCell = workSheet.Cells(constituantsRowIndex + nrLines, 1)
        lastCell = workSheet.Cells(workSheetRange.Rows.Count, nrCols)
        workSheetRange = workSheet.Range(firstCell, lastCell)
        flag = False
        list_constituants = list()

        for row in workSheetRange.Rows:
            flag = False
            for valueTuple in row.Value:
                for value in valueTuple:
                    if value != None:
                        flag = True
            if flag == False:
                TSDApp.WorkbookStats.constituantsLastRow = row.Row
                break

        for index in range(constituantsRowIndex + nrLines, TSDApp.WorkbookStats.constituantsLastRow):
            if workSheet.Cells(index, constituantsColIndex).Value == None:
                pass
            else:
                list_constituants.append(workSheet.Cells(index, constituantsColIndex).Value)

    for element in list_table:
        if element in list_constituants:
            localisation = None
        else:
            localisation = ""
            break
    if list_table == "[]":
        localisation = None

    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_COH_2030(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0
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
        nrCols = refCellRange.Columns.Count
        localisation = list()
        list_table = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
            if workSheet.Cells(index, refColIndex).Value == None:
                pass
            else:
                list_table.append(workSheet.Cells(index, refColIndex).Value)

    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.EffClientsIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        effColIndex = 0
        var = 0
        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Noms":
                    effColIndex = cellRow.index(cell) + 1
                    effRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if effColIndex != 0:
                break

        effCellRange = workSheet.Cells(effRowIndex, effColIndex).MergeArea
        nrLines = effCellRange.Rows.Count
        localisation = list()
        firstCell = workSheet.Cells(effRowIndex + nrLines, 1)
        lastCell = workSheet.Cells(workSheetRange.Rows.Count, nrCols)
        workSheetRange = workSheet.Range(firstCell, lastCell)
        flag = False
        list_eff = list()

        for row in workSheetRange.Rows:
            flag = False
            for valueTuple in row.Value:
                for value in valueTuple:
                    if value != None:
                        flag = True
            if flag == False:
                TSDApp.WorkbookStats.effLastRow = row.Row
                break

        for index in range(effRowIndex + nrLines, TSDApp.WorkbookStats.effLastRow):
            if workSheet.Cells(index, effColIndex).Value == None:
                pass
            else:
                list_eff.append(workSheet.Cells(index, effColIndex).Value)

    for element in list_table:
        if element in list_eff:
            localisation = None
        else:
            localisation = ""
            break
    if list_table == "[]":
        localisation = None

    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_COH_2040(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0
        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "DIAGNOSTIC DEBARQUE":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        nrCols = refCellRange.Columns.Count
        localisation = list()
        list_table = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
            if workSheet.Cells(index, refColIndex).Value == "N/A" or workSheet.Cells(index, refColIndex).Value == None:
                pass
            else:
                list_table.append(workSheet.Cells(index, refColIndex).Value)

    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagDebIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        diagColIndex = 0
        var = 0
        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "libellé (signification)":
                    diagColIndex = cellRow.index(cell) + 1
                    diagRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if diagColIndex != 0:
                break

        diagCellRange = workSheet.Cells(diagRowIndex, diagColIndex).MergeArea
        nrLines = diagCellRange.Rows.Count
        nrCols = diagCellRange.Columns.Count
        localisation = list()
        list_diag = list()

        for index in range(diagRowIndex + nrLines, TSDApp.WorkbookStats.DiagDebLastRow):
            if workSheet.Cells(index, diagColIndex).Value == None:
                pass
            else:
                list_code.append(workSheet.Cells(index, diagColIndex).Value)

    for element in list_table:
        if element in list_code:
            localisation = None
        else:
            localisation = ""
            break
    if list_table == "[]":
        localisation = None

    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_COH_2050(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0
        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Evenement(s) redouté(s) (ER)":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break

        refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
        nrLines = refCellRange.Rows.Count
        nrCols = refCellRange.Columns.Count
        localisation = list()
        list_table = list()

        for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
            if workSheet.Cells(index, refColIndex).Value == "No DTC" or workSheet.Cells(index, refColIndex).Value == None:
                pass
            else:
                list_table.append(workSheet.Cells(index, refColIndex).Value)

    if TSDApp.WorkbookStats.hasER == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ERIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        ERColIndex = 0
        var = 0
        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "nom":
                    ERColIndex = cellRow.index(cell) + 1
                    ERRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if ERColIndex != 0:
                break

        ERCellRange = workSheet.Cells(ERRowIndex, ERColIndex).MergeArea
        nrLines = ERCellRange.Rows.Count
        localisation = list()
        firstCell = workSheet.Cells(ERRowIndex + nrLines, 1)
        lastCell = workSheet.Cells(workSheetRange.Rows.Count, nrCols)
        workSheetRange = workSheet.Range(firstCell, lastCell)
        flag = False
        list_ER = list()

        for row in workSheetRange.Rows:
            flag = False
            for valueTuple in row.Value:
                for value in valueTuple:
                    if value != None:
                        flag = True
            if flag == False:
                TSDApp.WorkbookStats.ERLastRow = row.Row
                break

        for index in range(ERRowIndex + nrLines, TSDApp.WorkbookStats.ERLastRow):
            if workSheet.Cells(index, ERColIndex).Value == None:
                pass
            else:
                list_ER.append(workSheet.Cells(index, ERColIndex).Value)

    for element in list_table:
        if element in list_ER:
            localisation = None
        else:
            localisation = ""
            break
    if list_table == "[]":
        localisation = None

    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_COH_2060(ExcelApp, workBook, TSDApp, DOC7Name):
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.EffClientsIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        effColIndex = 0
        var = 0
        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Noms":
                    effColIndex = cellRow.index(cell) + 1
                    effRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if effColIndex != 0:
                break

        effCellRange = workSheet.Cells(effRowIndex, effColIndex).MergeArea
        nrLines = effCellRange.Rows.Count
        localisation = list()
        firstCell = workSheet.Cells(effRowIndex + nrLines, 1)
        lastCell = workSheet.Cells(workSheetRange.Rows.Count, nrCols)
        workSheetRange = workSheet.Range(firstCell, lastCell)
        flag = False
        list_eff = list()

        for row in workSheetRange.Rows:
            flag = False
            for valueTuple in row.Value:
                for value in valueTuple:
                    if value != None:
                        flag = True
            if flag == False:
                TSDApp.WorkbookStats.effLastRow = row.Row
                break

        for index in range(effRowIndex + nrLines, TSDApp.WorkbookStats.effLastRow):
            if workSheet.Cells(index, effColIndex).Value == None:
                pass
            else:
                list_eff.append(workSheet.Cells(index, effColIndex).Value)

        DOC7 = ExcelApp.Workbooks.Open(DOC7Name)
        try:
            workSheetRef = DOC7.Sheets("FR")
        except:
            workSheetRef = DOC7.Sheets("GB")

        workSheetRange = workSheetRef.UsedRange
        nrCols = workSheetRange.Columns.Count
        nrRows = workSheetRange.Rows.Count
        N1ColIndex = 0
        N2ColIndex = 0
        N2ColIndex = 0
        col = 0
        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Libellé N1":
                    N1ColIndex = cellRow.index(cell) + 1
                    N1RowIndex = workSheetRange.Value.index(cellRow) + 1
                    col += 1
                if cell == "Libellé N2":
                    N2ColIndex = cellRow.index(cell) + 1
                    N2RowIndex = workSheetRange.Value.index(cellRow) + 1
                    col += 1
                if cell == "Libellé N3":
                    N3ColIndex = cellRow.index(cell) + 1
                    N3RowIndex = workSheetRange.Value.index(cellRow) + 1
                    col += 1
                if col == 3:
                    break
            if col == 3:
                break

        try:
            refCellRange = workSheetRef.Cells(N1RowIndex, N1ColIndex).MergeArea
        except:
            try:
                refCellRange = workSheetRef.Cells(N2RowIndex, N2ColIndex).MergeArea
            except:
                refCellRange = workSheetRef.Cells(N3RowIndex, N3ColIndex).MergeArea


        nrLines = refCellRange.Rows.Count
        localisation = list()
        flag = False
        list_ref = list()

        if N1RowIndex != 0:
            for index in range(N1RowIndex + nrLines, nrRows + 1):
                if workSheetRef.Cells(index, N1ColIndex).Value == None:
                    pass
                else:
                    list_ref.append(workSheetRef.Cells(index, N1ColIndex).Value)
        elif N2RowIndex != 0:
            for index in range(N2RowIndex + nrLines, nrRows + 1):
                if workSheetRef.Cells(index, N2ColIndex).Value == None:
                    pass
                else:
                    list_ref.append(workSheetRef.Cells(index, N2ColIndex).Value)
        else:
            for index in range(N3RowIndex + nrLines, nrRows + 1):
                if workSheetRef.Cells(index, N3ColIndex).Value == None:
                    pass
                else:
                    list_ref.append(workSheetRef.Cells(index, N3ColIndex).Value)

        for element in list_eff:
            if element in list_ref:
                localisation = None
                pass
            else:
                localisation = ""
                break

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_COH_2070(ExcelApp, workBook, TSDApp, DOC7Name):
    if TSDApp.WorkbookStats.hasCustEff == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.CustEffIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        effColIndex = 0
        var = 0
        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Name":
                    effColIndex = cellRow.index(cell) + 1
                    effRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if effColIndex != 0:
                break

        effCellRange = workSheet.Cells(effRowIndex, effColIndex).MergeArea
        nrLines = effCellRange.Rows.Count
        localisation = list()
        firstCell = workSheet.Cells(effRowIndex + nrLines, 1)
        lastCell = workSheet.Cells(workSheetRange.Rows.Count, nrCols)
        workSheetRange = workSheet.Range(firstCell, lastCell)
        flag = False
        list_eff = list()

        for row in workSheetRange.Rows:
            flag = False
            for valueTuple in row.Value:
                for value in valueTuple:
                    if value != None:
                        flag = True
            if flag == False:
                TSDApp.WorkbookStats.CustEffLastRow = row.Row
                break

        for index in range(effRowIndex + nrLines, TSDApp.WorkbookStats.CustEffLastRow):
            if workSheet.Cells(index, effColIndex).Value == None:
                pass
            else:
                list_eff.append(workSheet.Cells(index, effColIndex).Value)

        DOC7 = ExcelApp.Workbooks.Open(DOC7Name)
        try:
            workSheetRef = DOC7.Sheets("FR")
        except:
            workSheetRef = DOC7.Sheets("GB")

        workSheetRange = workSheetRef.UsedRange
        nrCols = workSheetRange.Columns.Count
        nrRows = workSheetRange.Rows.Count
        N1ColIndex = 0
        N2ColIndex = 0
        N2ColIndex = 0
        col = 0
        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Libellé N1":
                    N1ColIndex = cellRow.index(cell) + 1
                    N1RowIndex = workSheetRange.Value.index(cellRow) + 1
                    col += 1
                if cell == "Libellé N2":
                    N2ColIndex = cellRow.index(cell) + 1
                    N2RowIndex = workSheetRange.Value.index(cellRow) + 1
                    col += 1
                if cell == "Libellé N3":
                    N3ColIndex = cellRow.index(cell) + 1
                    N3RowIndex = workSheetRange.Value.index(cellRow) + 1
                    col += 1
                if col == 3:
                    break
            if col == 3:
                break

        try:
            refCellRange = workSheetRef.Cells(N1RowIndex, N1ColIndex).MergeArea
        except:
            try:
                refCellRange = workSheetRef.Cells(N2RowIndex, N2ColIndex).MergeArea
            except:
                refCellRange = workSheetRef.Cells(N3RowIndex, N3ColIndex).MergeArea


        nrLines = refCellRange.Rows.Count
        localisation = list()
        flag = False
        list_ref = list()

        if N1RowIndex != 0:
            for index in range(N1RowIndex + nrLines, nrRows + 1):
                if workSheetRef.Cells(index, N1ColIndex).Value == None:
                    pass
                else:
                    list_ref.append(workSheetRef.Cells(index, N1ColIndex).Value)
        elif N2RowIndex != 0:
            for index in range(N2RowIndex + nrLines, nrRows + 1):
                if workSheetRef.Cells(index, N2ColIndex).Value == None:
                    pass
                else:
                    list_ref.append(workSheetRef.Cells(index, N2ColIndex).Value)
        else:
            for index in range(N3RowIndex + nrLines, nrRows + 1):
                if workSheetRef.Cells(index, N3ColIndex).Value == None:
                    pass
                else:
                    list_ref.append(workSheetRef.Cells(index, N3ColIndex).Value)

        for element in list_eff:
            if element in list_ref:
                localisation = None
                pass
            else:
                localisation = ""
                break

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

def Test_02043_18_04939_COH_2080(ExcelApp, workBook, TSDApp, DOC7Name):
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.EffClientsIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        effColIndex = 0
        var = 0
        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Noms":
                    effColIndex = cellRow.index(cell) + 1
                    effRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if effColIndex != 0:
                break

        effCellRange = workSheet.Cells(effRowIndex, effColIndex).MergeArea
        nrLines = effCellRange.Rows.Count
        localisation = list()
        firstCell = workSheet.Cells(effRowIndex + nrLines, 1)
        lastCell = workSheet.Cells(workSheetRange.Rows.Count, nrCols)
        workSheetRange = workSheet.Range(firstCell, lastCell)
        flag = False
        list_eff = list()

        for row in workSheetRange.Rows:
            flag = False
            for valueTuple in row.Value:
                for value in valueTuple:
                    if value != None:
                        flag = True
            if flag == False:
                TSDApp.WorkbookStats.effLastRow = row.Row
                break

        for index in range(effRowIndex + nrLines, TSDApp.WorkbookStats.effLastRow):
            if workSheet.Cells(index, effColIndex).Value == None:
                pass
            else:
                list_eff.append(workSheet.Cells(index, effColIndex).Value)

        DOC7 = ExcelApp.Workbooks.Open(DOC7Name)
        try:
            workSheetRef = DOC7.Sheets("FR")
        except:
            workSheetRef = DOC7.Sheets("GB")

        workSheetRange = workSheetRef.UsedRange
        nrCols = workSheetRange.Columns.Count
        nrRows = workSheetRange.Rows.Count
        N1ColIndex = 0
        N2ColIndex = 0
        N2ColIndex = 0
        col = 0
        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Libellé N1":
                    N1ColIndex = cellRow.index(cell) + 1
                    N1RowIndex = workSheetRange.Value.index(cellRow) + 1
                    col += 1
                if cell == "Libellé N2":
                    N2ColIndex = cellRow.index(cell) + 1
                    N2RowIndex = workSheetRange.Value.index(cellRow) + 1
                    col += 1
                if cell == "Libellé N3":
                    N3ColIndex = cellRow.index(cell) + 1
                    N3RowIndex = workSheetRange.Value.index(cellRow) + 1
                    col += 1
                if col == 3:
                    break
            if col == 3:
                break

        try:
            refCellRange = workSheetRef.Cells(N1RowIndex, N1ColIndex).MergeArea
        except:
            try:
                refCellRange = workSheetRef.Cells(N2RowIndex, N2ColIndex).MergeArea
            except:
                refCellRange = workSheetRef.Cells(N3RowIndex, N3ColIndex).MergeArea


        nrLines = refCellRange.Rows.Count
        localisation = list()
        flag = False
        list_ref = list()

        if N1RowIndex != 0:
            for index in range(N1RowIndex + nrLines, nrRows + 1):
                if workSheetRef.Cells(index, N1ColIndex).Value == None:
                    pass
                else:
                    list_ref.append(workSheetRef.Cells(index, N1ColIndex).Value)
        elif N2RowIndex != 0:
            for index in range(N2RowIndex + nrLines, nrRows + 1):
                if workSheetRef.Cells(index, N2ColIndex).Value == None:
                    pass
                else:
                    list_ref.append(workSheetRef.Cells(index, N2ColIndex).Value)
        else:
            for index in range(N3RowIndex + nrLines, nrRows + 1):
                if workSheetRef.Cells(index, N3ColIndex).Value == None:
                    pass
                else:
                    list_ref.append(workSheetRef.Cells(index, N3ColIndex).Value)

        for element in list_eff:
            if element in list_ref:
                localisation = None
                pass
            else:
                localisation = ""
                break

        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)