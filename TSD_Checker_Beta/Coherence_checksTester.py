import TSD_Checker_V1_0
import inspect
from ExcelEdit import TestReturn as result
from ErrorMessages import errorMessagesDict as error


#Coherence checks requirements

def Test_02043_18_04939_COH_2000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        var = 0

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "mesures et commandes (Mesure Parametre et Test Actionneur) / Tests de cohérence".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            nrCols = refCellRange.Columns.Count
            localisation = list()
            list_table = list()
            list_measure = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == "N/A" or workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    list_table.append(workSheet.Cells(index, refColIndex).Value)

            if TSDApp.WorkbookStats.hasMeasure == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
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
                    check = True
                    break
            if list_table == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2005(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        codeColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Code défaut":
                    codeColIndex = cell.Column
                    codeRowIndex = cell.Row
                    break
            if codeColIndex != 0:
                break
        if codeColIndex == 0:
            var = 1

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Code défaut".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1


        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
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
                    check = True

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
                        check = True
                        break

            if localisation == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2006(ExcelApp, workBook, TSDApp, DOC8Name):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.famillyList == "[]":
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        DOC8 = ExcelApp.Workbooks.Open(DOC8Name)
        workSheetRef = DOC8.Sheets("sous familles Cesare 2018 08 30")

        workSheetRange = workSheetRef.UsedRange
        nrCols = workSheetRange.Columns.Count
        nrRows = workSheetRange.Rows.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == " Nom de la sous famille ":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1


        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
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
                   check = True


            if localisation == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        return check

def Test_02043_18_04939_COH_2007(ExcelApp, workBook, TSDApp, DOC14Name):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.famillyList == "[]":
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:

        DOC14 = ExcelApp.Workbooks.Open(DOC14Name)
        workSheetRef = DOC14.Sheets("Matrix")

        workSheetRange = workSheetRef.UsedRange
        nrCols = workSheetRange.Columns.Count
        nrRows = workSheetRange.Rows.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Data Trouble Code (DTC)":
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1


        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
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
                   check = True


            if localisation == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        return check

def Test_02043_18_04939_COH_2010(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        list_code = list()
        tempList = list()
        var = 0

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Code défaut":
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            nrCols = refCellRange.Columns.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == "NO DTC" or workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    list_table = dict()
                    list_table["value"] = workSheet.Cells(index, refColIndex).Value
                    list_table["localisation"] = workSheet.Cells(index, refColIndex)
                    tempList.append(dict(list_table))

            if TSDApp.WorkbookStats.hasCode == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
                codeColIndex = 0
                var = 0
                for cellRow in workSheet.Rows:
                    for cell in cellRow.Cells:
                        if str(cell.Value).casefold() == "Code défaut":
                            codeColIndex = cell.Column
                            codeRowIndex = cell.Row
                            break
                    if codeColIndex != 0:
                        break

                codeCellRange = workSheet.Cells(codeRowIndex, codeColIndex).MergeArea
                nrLines = codeCellRange.Rows.Count
                nrCols = codeCellRange.Columns.Count
                localisation = list()


                for index in range(codeRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                    if workSheet.Cells(index, codeColIndex).Value == None:
                        pass
                    else:
                        list_code.append(workSheet.Cells(index, codeColIndex).Value)

            for element in tempList:
                if element["value"] in list_code:
                    pass
                else:
                    localisation.append(element["localisation"])
                    check = True

            if localisation == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2020(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        refColIndex = 0
        var = 0

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Constituant défaillant détecté":
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1


        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            nrCols = refCellRange.Columns.Count
            localisation = list()
            list_table = dict()
            list_constituants = list()
            tempList = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    list_table["value"] = workSheet.Cells(index, refColIndex).Value
                    list_table["localisation"] = workSheet.Cells(index, refColIndex)
                    tempList.append(dict(list_table))

            if TSDApp.WorkbookStats.hasConstituants == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.Sheets(TSDApp.WorkbookStats.constituantsIndex)
                constituantsColIndex = 0
                var = 0
                for cellRow in workSheet.Rows:
                    for cell in cellRow.Cells:
                        if str(cell.Value).casefold() == "Noms":
                            constituantsColIndex = cell.Column
                            constituantsRowIndex = cell.Row
                            break
                    if constituantsColIndex != 0:
                        break

                constituantsCellRange = workSheet.Cells(constituantsRowIndex, constituantsColIndex).MergeArea
                nrLines = constituantsCellRange.Rows.Count
                localisation = list()

                for index in range(constituantsRowIndex + nrLines, TSDApp.WorkbookStats.constituantsLastRow):
                    if workSheet.Cells(index, constituantsColIndex).Value == None:
                        pass
                    else:
                        list_constituants.append(workSheet.Cells(index, constituantsColIndex).Value)

            for element in tempList:
                if element["value"] in list_constituants:
                    pass
                else:
                    localisation.append(element["localisation"])
                    check = True

            if localisation == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2030(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        var = 0

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Effet(s) client(s)":
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1


        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            nrCols = refCellRange.Columns.Count
            localisation = list()
            list_table = dict()
            list_eff = list()
            tempList = list()


            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    list_table["value"] = workSheet.Cells(index, refColIndex).Value
                    list_table["localisation"] = workSheet.Cells(index, refColIndex)
                    tempList.append(dict(list_table))

            if TSDApp.WorkbookStats.hasEffClients == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.Sheets(TSDApp.WorkbookStats.EffClientsIndex)
                workSheetRange = workSheet.UsedRange
                nrCols = workSheetRange.Columns.Count
                nrRows = workSheetRange.Rows.Count
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

                for row in workSheetRange.Rows:
                    flag = False
                    for valueTuple in row.Value:
                        for value in valueTuple:
                            if value != None:
                                flag = True
                    if flag == False:
                        TSDApp.WorkbookStats.effLastRow = row.Row
                        break

                for index in range(effRowIndex + nrLines, nrRows):
                    if workSheet.Cells(index, effColIndex).Value == None:
                        pass
                    else:
                        list_eff.append(workSheet.Cells(index, effColIndex).Value)

            for element in tempList:
                if element["value"] in list_eff:
                    pass
                else:
                    localisation.append(element["localisation"])
                    check = True

            if localisation == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2040(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        nrRows = workSheet.Rows.Count
        refColIndex = 0
        var = 0
        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "DIAGNOSTIC DEBARQUE":
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1


        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            nrCols = refCellRange.Columns.Count
            localisation = list()
            list_table = dict()
            list_diag = list()
            tempList = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == "N/A" or workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    list_table["value"] = workSheet.Cells(index, refColIndex).Value
                    list_table["localisation"] = workSheet.Cells(index, refColIndex)
                    tempList.append(dict(list_table))

            if TSDApp.WorkbookStats.hasDiagDeb == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagDebIndex)
                nrRows = workSheet.Rows.Count
                diagColIndex = 0
                var = 0
                for cellRow in workSheet.Rows:
                    for cell in cellRow.Cells:
                        if str(cell.Value).casefold() == "libellé (signification)":
                            diagColIndex = cellRow.index(cell) + 1
                            diagRowIndex = workSheetRange.Value.index(cellRow) + 1
                            break
                    if diagColIndex != 0:
                        break

                diagCellRange = workSheet.Cells(diagRowIndex, diagColIndex).MergeArea
                nrLines = diagCellRange.Rows.Count
                nrCols = diagCellRange.Columns.Count
                localisation = list()

                for index in range(diagRowIndex + nrLines, nrRows):
                    if workSheet.Cells(index, diagColIndex).Value == None:
                        pass
                    else:
                        list_diag.append(workSheet.Cells(index, diagColIndex).Value)

            for element in tempList:
                if element["value"] in list_diag:
                    pass
                else:
                    localisation.append(element["localisation"])
                    check = True

            if localisation == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2050(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        nrRows = workSheet.Rows.Count
        refColIndex = 0
        var = 0
        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Evenement(s) redouté(s) (ER)":
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            nrCols = refCellRange.Columns.Count
            localisation = list()
            list_table = dict()
            list_ER = list()
            tempList = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == "No DTC" or workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    list_table["value"] = workSheet.Cells(index, refColIndex).Value
                    list_table["localisation"] = workSheet.Cells(index, refColIndex)
                    tempList.append(dict(list_table))

            if TSDApp.WorkbookStats.hasER == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.Sheets(TSDApp.WorkbookStats.ERIndex)
                nrRows = workSheet.Rows.Count
                ERColIndex = 0
                var = 0
                for cellRow in workSheet.Rows:
                    for cell in cellRow.Cells:
                        if str(cell.Value).casefold() == "nom":
                            ERColIndex = cellRow.index(cell) + 1
                            ERRowIndex = workSheetRange.Value.index(cellRow) + 1
                            break
                    if ERColIndex != 0:
                        break

                ERCellRange = workSheet.Cells(ERRowIndex, ERColIndex).MergeArea
                nrLines = ERCellRange.Rows.Count
                localisation = list()
                for index in range(ERRowIndex + nrLines, nrRows):
                    if workSheet.Cells(index, ERColIndex).Value == None:
                        pass
                    else:
                        list_ER.append(workSheet.Cells(index, ERColIndex).Value)

            for element in tempList:
                if element["value"] in list_ER:
                    pass
                else:
                    localisation.append(element["localisation"])
                    check = True

            if localisation == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2060(ExcelApp, workBook, TSDApp, DOC7Name):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.EffClientsIndex)
        nrRows = workSheet.Rows
        effColIndex = 0
        var = 0
        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Noms":
                    effColIndex = cellRow.index(cell) + 1
                    effRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if effColIndex != 0:
                break
        if effColIndex == 0:
            var = 1

        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            effCellRange = workSheet.Cells(effRowIndex, effColIndex).MergeArea
            nrLines = effCellRange.Rows.Count
            localisation = list()

            list_eff = list()
            list_ref = list()

            for index in range(effRowIndex + nrLines, nrRows):
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
                    check = True
                    break

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2070(ExcelApp, workBook, TSDApp, DOC7Name):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCustEff == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.CustEffIndex)
        effColIndex = 0
        var = 0
        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Name":
                    effColIndex = cellRow.index(cell) + 1
                    effRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if effColIndex != 0:
                break
        if effColIndex == 0:
            var = 1

        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            effCellRange = workSheet.Cells(effRowIndex, effColIndex).MergeArea
            nrLines = effCellRange.Rows.Count
            localisation = list()
            firstCell = workSheet.Cells(effRowIndex + nrLines, 1)
            lastCell = workSheet.Cells(workSheetRange.Rows.Count, nrCols)
            workSheetRange = workSheet.Range(firstCell, lastCell)
            flag = False
            list_eff = list()
            list_ref = list()

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
                    check = True
                    break

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return  check

def Test_02043_18_04939_COH_2080(ExcelApp, workBook, TSDApp, DOC7Name):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
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
        if effColIndex == 0:
            var = 1

        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            effCellRange = workSheet.Cells(effRowIndex, effColIndex).MergeArea
            nrLines = effCellRange.Rows.Count
            localisation = list()
            firstCell = workSheet.Cells(effRowIndex + nrLines, 1)
            lastCell = workSheet.Cells(workSheetRange.Rows.Count, nrCols)
            workSheetRange = workSheet.Range(firstCell, lastCell)
            flag = False
            list_eff = list()
            list_ref = list()

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
                    check = True
                    break

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2091(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    temp = workBook.Sheets
    sheetNames = list()
    localisation = list()
    check = False
    for sheet in temp:
        sheetNames.append(sheet.Name.strip().casefold())

    for name in sheetNames:
        index = sheetNames.index(name) + 1
        workSheet = workBook.Sheets(index)
        workSheetRange = workSheet.UsedRange
        nrLines = workSheetRange.Rows.Count
        nrCols = workSheetRange.Columns.Count
        localisation = list()
        firtCell = workSheet.Cells(1, 1)
        lastCell = workSheet.Cells(nrLines, nrCols)
        workSheetRange = workSheet.Range(firtCell, lastCell)
        flag = False

        for row in workSheetRange.Rows:
            flag = False
            for valueTuple in row.Value:
                for value in valueTuple:
                    if value != None:
                        flag = True
            if flag == False:
                lastRow = row.Row
                break

            for rowIndex in range(1, nrLines):
                for colIndex in range(1, nrCols):
                    if workSheet.Cells(rowIndex, colIndex).Value == "?" or  workSheet.Cells(rowIndex, colIndex).Value == "tbd" or workSheet.Cells(rowIndex, colIndex).Value == "tbc":
                        localisation.append(workSheet.Cells(rowIndex, colIndex))
                        check = True

    if localisation == "[]":
        localisation = None

    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2100(ExcelApp, workBook, TSDApp, DOC8Name):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        list_famille = list()
        tempDict = list()
        var = 0
        localisation = list()

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "supporté par constituant (s)":
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1


        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    tempDict.append(workSheet.Cells(index, refColIndex).Value)


            DOC8 = ExcelApp.Workbooks.Open(DOC8Name)
            workSheetRef = DOC8.Sheets("sous familles Cesare 2018 08 30")

            workSheetRange = workSheetRef.UsedRange
            nrCols = workSheetRange.Columns.Count
            nrRows = workSheetRange.Rows.Count
            familleColIndex = 0

            for cellRow in workSheetRange.Value:
                for cell in cellRow:
                    if cell == " Nom de la sous famille ":
                        familleColIndex = cellRow.index(cell) + 1
                        familleRowIndex = workSheetRange.Value.index(cellRow) + 1
                        break
                if familleColIndex != 0:
                    break

            familleCellRange = workSheetRef.Cells(familleRowIndex, familleColIndex).MergeArea
            nrLines = familleCellRange.Rows.Count
            localisation = list()

            for index in range(familleRowIndex + nrLines, nrRows + 1):
                if workSheetRef.Cells(index, familleColIndex).Value == None:
                    pass
                else:
                    list_famille.append(workSheetRef.Cells(index, familleColIndex).Value)

            if len(tempDict) == 0:
                localisation = None
            else:
                for element in tempDict:
                    if element in list_famille:
                        pass
                    else:
                       localisation = ""
                       check = True

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2110(ExcelApp, workBook, TSDApp, DOC8Name):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)

        refColIndex = 0
        list_famille = list()
        tempDict = list()
        var = 0
        localisation = list()

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "supporté par constituant (s)":
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1


        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    tempDict.append(workSheet.Cells(index, refColIndex).Value)

            DOC8 = ExcelApp.Workbooks.Open(DOC8Name)
            workSheetRef = DOC8.Sheets("sous familles Cesare 2018 08 30")

            workSheetRange = workSheetRef.UsedRange
            nrCols = workSheetRange.Columns.Count
            nrRows = workSheetRange.Rows.Count
            familleColIndex = 0

            for cellRow in workSheetRange.Value:
                for cell in cellRow:
                    if cell == " Nom de la sous famille ":
                        familleColIndex = cellRow.index(cell) + 1
                        familleRowIndex = workSheetRange.Value.index(cellRow) + 1
                        break
                if familleColIndex != 0:
                    break

            familleCellRange = workSheetRef.Cells(familleRowIndex, familleColIndex).MergeArea
            nrLines = familleCellRange.Rows.Count
            localisation = list()

            for index in range(familleRowIndex + nrLines, nrRows + 1):
                if workSheetRef.Cells(index, familleColIndex).Value == None:
                    pass
                else:
                    list_famille.append(workSheetRef.Cells(index, familleColIndex).Value)

            if len(tempDict) == 0:
                localisation = None
            else:
                for element in tempDict:
                    if element in list_famille:
                        pass
                    else:
                       localisation = ""
                       check = True

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2120(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTechEff == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.TechEffIndex)
        nrCols = workSheet.Columns.Count
        nrRows = workSheet.Rows.Count
        refColIndex = 0
        list_amont = list()
        tempDict = list()
        var = 0
        localisation = list()

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Reference" or str(cell.Value).casefold() == "Référence":
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()
            flag = False

            for index in range(refRowIndex + nrLines, nrRows):
                if workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    tempDict.append(workSheet.Cells(index, refColIndex).Value)


            DOC5 = ExcelApp.Workbooks.Open(DOC5Name)
            workSheetRef = DOC5.Sheets("Effets techniques")
            workSheetRange = workSheetRef.UsedRange
            nrCols = workSheetRange.Columns.Count
            nrRows = workSheetRange.Rows.Count
            amontColIndex = 0

            for cellRow in workSheetRange.Value:
                for cell in cellRow:
                    if cell == "Référence amont":
                        amontColIndex = cellRow.index(cell) + 1
                        amontRowIndex = workSheetRange.Value.index(cellRow) + 1
                        break
                if amontColIndex != 0:
                    break

            amontCellRange = workSheetRef.Cells(amontRowIndex, amontColIndex).MergeArea
            nrLines = amontCellRange.Rows.Count
            flag = False
            for index in range(amontRowIndex + nrLines, nrRows + 1):
                if workSheetRef.Cells(index, amontColIndex).Value == None:
                    pass
                else:
                    list_amont.append(workSheetRef.Cells(index, amontColIndex).Value)

            for element in tempDict:
                if element in list_amont:
                    pass
                else:
                   localisation = ""
                   check = True

            if tempDict == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2130(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        var = 0
        localisation = list()

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Référence" or str(cell.Value) == "Reference":
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1


        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()
            list_table = list()
            list_effets = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    list_table.append(workSheet.Cells(index, refColIndex).Value)


            if TSDApp.WorkbookStats.hasTechEff == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.Sheets(TSDApp.WorkbookStats.TechEffIndex)
                workSheetRange = workSheet.UsedRange
                nrCols = workSheetRange.Columns.Count
                nrRows = workSheetRange.Rows.Count
                effColIndex = 0

                for cellRow in workSheetRange.Value:
                    for cell in cellRow:
                        if cell == "Référence amont":
                            effColIndex = cellRow.index(cell) + 1
                            effRowIndex = workSheetRange.Value.index(cellRow) + 1
                            break
                    if effColIndex != 0:
                        break

                if effColIndex != 0:

                    effCellRange = workSheet.Cells(effRowIndex, effColIndex).MergeArea
                    nrLines = effCellRange.Rows.Count
                    nrCols = effCellRange.Columns.Count
                    firstCell = workSheet.Cells(refRowIndex + nrLines, 1)
                    lastCell = workSheet.Cells(workSheetRange.Rows.Count, nrCols)
                    workSheetRange = workSheet.Range(firstCell, lastCell)
                    flag = False

                    for row in workSheetRange.Rows:
                        flag = False
                        for valueTuple in row.Value:
                            for value in valueTuple:
                                if value != None:
                                    flag = True
                        if flag == False:
                            TSDApp.WorkbookStats.TechEffLastRow = row.Row
                            break


                    for index in range(effRowIndex + nrLines, nrRows):
                        if workSheet.Cells(index, effColIndex).Value == None:
                            pass
                        else:
                            list_effets.append(workSheet.Cells(index, effColIndex).Value)

                    for element in list_table:
                        if element in list_effets:
                            localisation = None
                        else:
                            localisation = ""
                            check = True

                    if list_table == "[]":
                        localisation = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

                elif effColIndex == 0:
                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                    check =True
    return check

def Test_02043_18_04939_COH_2140(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        var = 0
        localisation = list()

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "code defauts induits":
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()
            list_table = list()
            list_effets = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    list_table.append(workSheet.Cells(index, refColIndex).Value)


            if TSDApp.WorkbookStats.hasCode == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
                effColIndex = 0

                for cellRow in workSheet.Rows:
                    for cell in cellRow.Cells:
                        if str(cell.Value).casefold() == "Code défaut":
                            effColIndex = cellRow.index(cell) + 1
                            effRowIndex = workSheetRange.Value.index(cellRow) + 1
                            break
                    if effColIndex != 0:
                        break

                if effColIndex != 0:

                    effCellRange = workSheet.Cells(effRowIndex, effColIndex).MergeArea
                    nrLines = effCellRange.Rows.Count
                    nrCols = effCellRange.Columns.Count


                    for index in range(effRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                        if workSheet.Cells(index, effColIndex).Value == None:
                            pass
                        else:
                            list_effets.append(workSheet.Cells(index, effColIndex).Value)

                    for element in list_table:
                        if element in list_effets or element == "N/A":
                            localisation = None
                        else:
                            localisation = ""
                            check = True

                    if list_table == "[]":
                        localisation = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

                elif effColIndex == 0:
                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                    check = True
    return check

def Test_02043_18_04939_COH_2150(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        var = 0
        localisation = list()

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Supporté par constituant(s)":
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()
            list_table = list()
            list_effets = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    list_table.append(workSheet.Cells(index, refColIndex).Value)


            if TSDApp.WorkbookStats.hasConstituants == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.Sheets(TSDApp.WorkbookStats.constituantsIndex)
                workSheetRange = workSheet.UsedRange
                nrCols = workSheetRange.Columns.Count
                nrRows = workSheetRange.Rows.Count
                effColIndex = 0

                for cellRow in workSheetRange.Value:
                    for cell in cellRow:
                        if cell == "Noms":
                            effColIndex = cellRow.index(cell) + 1
                            effRowIndex = workSheetRange.Value.index(cellRow) + 1
                            break
                    if effColIndex != 0:
                        break

                if effColIndex != 0:

                    effCellRange = workSheet.Cells(effRowIndex, effColIndex).MergeArea
                    nrLines = effCellRange.Rows.Count
                    nrCols = effCellRange.Columns.Count
                    firstCell = workSheet.Cells(refRowIndex + nrLines, 1)
                    lastCell = workSheet.Cells(workSheetRange.Rows.Count, nrCols)
                    workSheetRange = workSheet.Range(firstCell, lastCell)
                    flag = False

                    for row in workSheetRange.Rows:
                        flag = False
                        for valueTuple in row.Value:
                            for value in valueTuple:
                                if value != None:
                                    flag = True
                        if flag == False:
                            TSDApp.WorkbookStats.TechEffLastRow = row.Row
                            break


                    for index in range(effRowIndex + nrLines, nrRows):
                        if workSheet.Cells(index, effColIndex).Value == None:
                            pass
                        else:
                            list_effets.append(workSheet.Cells(index, effColIndex).Value)

                    for element in list_table:
                        if element in list_effets:
                            localisation = None
                        else:
                            localisation = ""
                            check = True

                    if list_table == "[]":
                        localisation = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

                elif effColIndex == 0:
                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                    check = True
    return check

def Test_02043_18_04939_COH_2160(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)

        refColIndex = 0
        var = 0
        localisation = list()

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Supporté par constituant(s)":
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()
            list_table = list()
            list_effets = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    list_table.append(workSheet.Cells(index, refColIndex).Value)

            if TSDApp.WorkbookStats.hasConstituants == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.Sheets(TSDApp.WorkbookStats.constituantsIndex)
                workSheetRange = workSheet.UsedRange
                nrCols = workSheetRange.Columns.Count
                nrRows = workSheetRange.Rows.Count
                effColIndex = 0

                for cellRow in workSheetRange.Value:
                    for cell in cellRow:
                        if cell == "Noms":
                            effColIndex = cellRow.index(cell) + 1
                            effRowIndex = workSheetRange.Value.index(cellRow) + 1
                            break
                    if effColIndex != 0:
                        break

                if effColIndex != 0:

                    effCellRange = workSheet.Cells(effRowIndex, effColIndex).MergeArea
                    nrLines = effCellRange.Rows.Count
                    nrCols = effCellRange.Columns.Count
                    firstCell = workSheet.Cells(refRowIndex + nrLines, 1)
                    lastCell = workSheet.Cells(workSheetRange.Rows.Count, nrCols)
                    workSheetRange = workSheet.Range(firstCell, lastCell)
                    flag = False

                    for row in workSheetRange.Rows:
                        flag = False
                        for valueTuple in row.Value:
                            for value in valueTuple:
                                if value != None:
                                    flag = True
                        if flag == False:
                            TSDApp.WorkbookStats.TechEffLastRow = row.Row
                            break


                    for index in range(effRowIndex + nrLines, nrRows):
                        if workSheet.Cells(index, effColIndex).Value == None:
                            pass
                        else:
                            list_effets.append(workSheet.Cells(index, effColIndex).Value)

                    for element in list_table:
                        if element in list_effets:
                            localisation = None
                        else:
                            localisation = ""
                            check = True

                    if list_table == "[]":
                        localisation = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

                elif effColIndex == 0:
                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                    check = True
    return check

def Test_02043_18_04939_COH_2170(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DataCodesIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        nrRows = workSheetRange.Rows.Count
        refColIndex = 0
        var = 0
        localisation = list()

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Stored by the ECU":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()
            list_table = list()
            list_effets = list()

            for index in range(refRowIndex + nrLines, nrRows):
                if workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    list_table.append(workSheet.Cells(index, refColIndex).Value)


            if TSDApp.WorkbookStats.hasConstituants == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.Sheets(TSDApp.WorkbookStats.constituantsIndex)
                workSheetRange = workSheet.UsedRange
                nrCols = workSheetRange.Columns.Count
                nrRows = workSheetRange.Rows.Count
                effColIndex = 0

                for cellRow in workSheetRange.Value:
                    for cell in cellRow:
                        if cell == "Name":
                            effColIndex = cellRow.index(cell) + 1
                            effRowIndex = workSheetRange.Value.index(cellRow) + 1
                            break
                    if effColIndex != 0:
                        break

                if effColIndex != 0:

                    effCellRange = workSheet.Cells(effRowIndex, effColIndex).MergeArea
                    nrLines = effCellRange.Rows.Count
                    nrCols = effCellRange.Columns.Count
                    firstCell = workSheet.Cells(refRowIndex + nrLines, 1)
                    lastCell = workSheet.Cells(workSheetRange.Rows.Count, nrCols)
                    workSheetRange = workSheet.Range(firstCell, lastCell)
                    flag = False

                    for row in workSheetRange.Rows:
                        flag = False
                        for valueTuple in row.Value:
                            for value in valueTuple:
                                if value != None:
                                    flag = True
                        if flag == False:
                            TSDApp.WorkbookStats.TechEffLastRow = row.Row
                            break


                    for index in range(effRowIndex + nrLines, nrRows):
                        if workSheet.Cells(index, effColIndex).Value == None:
                            pass
                        else:
                            list_effets.append(workSheet.Cells(index, effColIndex).Value)

                    for element in list_table:
                        if element in list_effets:
                            localisation = None
                        else:
                            localisation = ""
                            check = True

                    if list_table == "[]":
                        localisation = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

                elif effColIndex == 0:
                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                    check = True
    return check

def Test_02043_18_04939_COH_2180(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReadDataIOIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        nrRows = workSheetRange.Rows.Count
        refColIndex = 0
        var = 0
        localisation = list()

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Stored by the ECU":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()
            list_table = list()
            list_effets = list()

            for index in range(refRowIndex + nrLines, nrRows):
                if workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    list_table.append(workSheet.Cells(index, refColIndex).Value)

            if TSDApp.WorkbookStats.hasConstituants == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.Sheets(TSDApp.WorkbookStats.constituantsIndex)
                workSheetRange = workSheet.UsedRange
                nrCols = workSheetRange.Columns.Count
                nrRows = workSheetRange.Rows.Count
                effColIndex = 0

                for cellRow in workSheetRange.Value:
                    for cell in cellRow:
                        if cell == "Name":
                            effColIndex = cellRow.index(cell) + 1
                            effRowIndex = workSheetRange.Value.index(cellRow) + 1
                            break
                    if effColIndex != 0:
                        break

                if effColIndex != 0:

                    effCellRange = workSheet.Cells(effRowIndex, effColIndex).MergeArea
                    nrLines = effCellRange.Rows.Count
                    nrCols = effCellRange.Columns.Count
                    firstCell = workSheet.Cells(refRowIndex + nrLines, 1)
                    lastCell = workSheet.Cells(workSheetRange.Rows.Count, nrCols)
                    workSheetRange = workSheet.Range(firstCell, lastCell)
                    flag = False

                    for row in workSheetRange.Rows:
                        flag = False
                        for valueTuple in row.Value:
                            for value in valueTuple:
                                if value != None:
                                    flag = True
                        if flag == False:
                            TSDApp.WorkbookStats.TechEffLastRow = row.Row
                            break


                    for index in range(effRowIndex + nrLines, nrRows):
                        if workSheet.Cells(index, effColIndex).Value == None:
                            pass
                        else:
                            list_effets.append(workSheet.Cells(index, effColIndex).Value)

                    for element in list_table:
                        if element in list_effets:
                            localisation = None
                        else:
                            localisation = ""
                            check = True

                    if list_table == "[]":
                        localisation = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

                elif effColIndex == 0:
                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                    check = True
    return check

def Test_02043_18_04939_COH_2190(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        var = 0
        localisation = list()

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "situation de vie":
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()
            list_table = list()
            list_effets = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    list_table.append(workSheet.Cells(index, refColIndex).Value)


            if TSDApp.WorkbookStats.hasSitDeVie == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.Sheets(TSDApp.WorkbookStats.SitDeVieIndex)
                workSheetRange = workSheet.UsedRange
                nrCols = workSheetRange.Columns.Count
                nrRows = workSheetRange.Rows.Count
                effColIndex = 0

                for cellRow in workSheetRange.Value:
                    for cell in cellRow:
                        if cell == "situation de vie":
                            effColIndex = cellRow.index(cell) + 1
                            effRowIndex = workSheetRange.Value.index(cellRow) + 1
                            break
                    if effColIndex != 0:
                        break

                if effColIndex != 0:

                    effCellRange = workSheet.Cells(effRowIndex, effColIndex).MergeArea
                    nrLines = effCellRange.Rows.Count
                    nrCols = effCellRange.Columns.Count
                    firstCell = workSheet.Cells(refRowIndex + nrLines, 1)
                    lastCell = workSheet.Cells(workSheetRange.Rows.Count, nrCols)
                    workSheetRange = workSheet.Range(firstCell, lastCell)
                    flag = False

                    for row in workSheetRange.Rows:
                        flag = False
                        for valueTuple in row.Value:
                            for value in valueTuple:
                                if value != None:
                                    flag = True
                        if flag == False:
                            TSDApp.WorkbookStats.TechEffLastRow = row.Row
                            break


                    for index in range(effRowIndex + nrLines, nrRows):
                        if workSheet.Cells(index, effColIndex).Value == None:
                            pass
                        else:
                            list_effets.append(workSheet.Cells(index, effColIndex).Value)

                    for element in list_table:
                        if element in list_effets or element == "N/A":
                            localisation = None
                        else:
                            localisation = ""
                            check = True

                    if list_table == "[]":
                        localisation = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

                elif effColIndex == 0:
                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                    check = True
    return check

def Test_02043_18_04939_COH_2200(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        var = 0
        localisation = list()

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Situation":
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()
            list_table = list()
            list_effets = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    list_table.append(workSheet.Cells(index, refColIndex).Value)


            if TSDApp.WorkbookStats.hasSitDeVie == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.Sheets(TSDApp.WorkbookStats.SitDeVieIndex)
                workSheetRange = workSheet.UsedRange
                nrCols = workSheetRange.Columns.Count
                nrRows = workSheetRange.Rows.Count
                effColIndex = 0

                for cellRow in workSheetRange.Value:
                    for cell in cellRow:
                        if cell == "Description":
                            effColIndex = cellRow.index(cell) + 1
                            effRowIndex = workSheetRange.Value.index(cellRow) + 1
                            break
                    if effColIndex != 0:
                        break

                if effColIndex != 0:

                    effCellRange = workSheet.Cells(effRowIndex, effColIndex).MergeArea
                    nrLines = effCellRange.Rows.Count
                    nrCols = effCellRange.Columns.Count
                    firstCell = workSheet.Cells(refRowIndex + nrLines, 1)
                    lastCell = workSheet.Cells(workSheetRange.Rows.Count, nrCols)
                    workSheetRange = workSheet.Range(firstCell, lastCell)
                    flag = False

                    for row in workSheetRange.Rows:
                        flag = False
                        for valueTuple in row.Value:
                            for value in valueTuple:
                                if value != None:
                                    flag = True
                        if flag == False:
                            TSDApp.WorkbookStats.TechEffLastRow = row.Row
                            break


                    for index in range(effRowIndex + nrLines, nrRows):
                        if workSheet.Cells(index, effColIndex).Value == None:
                            pass
                        else:
                            list_effets.append(workSheet.Cells(index, effColIndex).Value)

                    for element in list_table:
                        if element in list_effets or element == "N/A":
                            localisation = None
                        else:
                            localisation = ""
                            check = True

                    if list_table == "[]":
                        localisation = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

                elif effColIndex == 0:
                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                    check = True
    return check

def Test_02043_18_04939_COH_2210(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        var = 0
        localisation = list()

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Diagnostic débarqué":
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()
            list_table = list()
            list_effets = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    list_table.append(workSheet.Cells(index, refColIndex).Value)


            if TSDApp.WorkbookStats.hasDiagDeb == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagDebIndex)
                workSheetRange = workSheet.UsedRange
                nrCols = workSheetRange.Columns.Count
                nrRows = workSheetRange.Rows.Count
                effColIndex = 0

                for cellRow in workSheetRange.Value:
                    for cell in cellRow:
                        if cell == "libellé (signification)":
                            effColIndex = cellRow.index(cell) + 1
                            effRowIndex = workSheetRange.Value.index(cellRow) + 1
                            break
                    if effColIndex != 0:
                        break

                if effColIndex != 0:

                    effCellRange = workSheet.Cells(effRowIndex, effColIndex).MergeArea
                    nrLines = effCellRange.Rows.Count
                    nrCols = effCellRange.Columns.Count
                    firstCell = workSheet.Cells(refRowIndex + nrLines, 1)
                    lastCell = workSheet.Cells(workSheetRange.Rows.Count, nrCols)
                    workSheetRange = workSheet.Range(firstCell, lastCell)
                    flag = False

                    for row in workSheetRange.Rows:
                        flag = False
                        for valueTuple in row.Value:
                            for value in valueTuple:
                                if value != None:
                                    flag = True
                        if flag == False:
                            TSDApp.WorkbookStats.TechEffLastRow = row.Row
                            break


                    for index in range(effRowIndex + nrLines, nrRows):
                        if workSheet.Cells(index, effColIndex).Value == None:
                            pass
                        else:
                            list_effets.append(workSheet.Cells(index, effColIndex).Value)

                    for element in list_table:
                        if element in list_effets or element == "N/A":
                            localisation = None
                        else:
                            localisation = ""
                            check = True

                    if list_table == "[]":
                        localisation = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

                elif effColIndex == 0:
                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                    check = True
    return check

def Test_02043_18_04939_COH_2220(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        var = 0
        localisation = list()

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Non-embedded diagnosis":
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()
            list_table = list()
            list_effets = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    list_table.append(workSheet.Cells(index, refColIndex).Value)


            if TSDApp.WorkbookStats.hasNotEmbDiag == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.Sheets(TSDApp.WorkbookStats.NotEmbDiagIndex)
                workSheetRange = workSheet.UsedRange
                nrCols = workSheetRange.Columns.Count
                nrRows = workSheetRange.Rows.Count
                effColIndex = 0

                for cellRow in workSheetRange.Value:
                    for cell in cellRow:
                        if cell == "Label":
                            effColIndex = cellRow.index(cell) + 1
                            effRowIndex = workSheetRange.Value.index(cellRow) + 1
                            break
                    if effColIndex != 0:
                        break

                if effColIndex != 0:

                    effCellRange = workSheet.Cells(effRowIndex, effColIndex).MergeArea
                    nrLines = effCellRange.Rows.Count
                    nrCols = effCellRange.Columns.Count
                    firstCell = workSheet.Cells(refRowIndex + nrLines, 1)
                    lastCell = workSheet.Cells(workSheetRange.Rows.Count, nrCols)
                    workSheetRange = workSheet.Range(firstCell, lastCell)
                    flag = False

                    for row in workSheetRange.Rows:
                        flag = False
                        for valueTuple in row.Value:
                            for value in valueTuple:
                                if value != None:
                                    flag = True
                        if flag == False:
                            TSDApp.WorkbookStats.TechEffLastRow = row.Row
                            break


                    for index in range(effRowIndex + nrLines, nrRows):
                        if workSheet.Cells(index, effColIndex).Value == None:
                            pass
                        else:
                            list_effets.append(workSheet.Cells(index, effColIndex).Value)

                    for element in list_table:
                        if element in list_effets or element == "N/A":
                            localisation = None
                        else:
                            localisation = ""
                            check = True

                    if list_table == "[]":
                        localisation = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

                elif effColIndex == 0:
                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                    check = True
    return check

