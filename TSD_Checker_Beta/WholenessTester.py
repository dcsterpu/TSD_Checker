import TSD_Checker_V1_0
import inspect
from ExcelEdit import TestReturn as result
from ErrorMessages import errorMessagesDict as error



def Test_02043_18_04939_WHOLENESS_1000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        #workSheetRange = workSheet.UsedRange
        nrRows = workSheet.Rows.Count
        nrCols = workSheet.Columns.Count
        #nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0
        ok = 0
        col_range = 0
        lastCol = 0
        lastRow = 0
        tmp = 0

        for cellRow in workSheet.Rows:
            col_range = 0
            for cell in cellRow.Cells:
                if ok == 0:
                    if str(cell.Value).casefold() == "Référence".casefold().strip() or str(cell.Value).casefold().strip() == "Reference".casefold():
                        refColIndex = cell.Column
                        refRowIndex = cell.Row
                        indexCol = 1
                        col_range = 1
                    if col_range == 1:
                        if cell.Borders(8).LineStyle != -4142 and cell != None:
                            indexCol += 1
                            pass
                        else:
                           lastCol = cell.Column
                           tmp = 1
                           ok = 1
                           break

                if tmp != 0:
                    ok = 1
                    if col_range == 0:
                        if cell.Borders(9).LineStyle != -4142:
                            pass
                        else:
                           lastRow = cell.Row
                           break


        if refColIndex == 0:
            var = 1

        '''for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if cell.Borders(8).LineStyle != -4142:
                    pass
                else:
                   pass'''

        if var == 0:
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
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check  = True
    return check

def Test_02043_18_04939_WHOLENESS_1001(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Version":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1010(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Référence" or cell == "Reference":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
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
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1011(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Version":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1020(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Référence":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
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
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1021(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Version":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1030(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagDebIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Référence":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
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
                    TSDApp.WorkbookStats.DiagDebLastRow = row.Row
                    break

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.DiagDebLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1031(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagDebIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Version":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.DiagDebLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1040(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMDD == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.MDDIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0
        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Référence":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
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
                    TSDApp.WorkbookStats.MDDLastRow = row.Row
                    break

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.MDDLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1041(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMDD == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.MDDIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0
        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Version":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.MDDLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1050(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        list_code = list()
        list_table = list()
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Applicabilité projet" or cell == "Project applicability":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            nrCols = refCellRange.Columns.Count
            localisation = list()


            for index in range(refColIndex, refColIndex + nrCols):
                if workSheet.Cells(refRowIndex + nrLines, refColIndex).Value == None:
                    pass
                else:
                    list_table.append(workSheet.Cells(refRowIndex + nrLines, index).Value)
                    check = True

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
                        if cell == "Applicabilité projet" or cell == "Project applicability":
                            codeColIndex = cellRow.index(cell) + 1
                            codeRowIndex = workSheetRange.Value.index(cellRow) + 1
                            break
                    if codeColIndex != 0:
                        break

                codeCellRange = workSheet.Cells(codeRowIndex, codeColIndex).MergeArea
                nrLines = codeCellRange.Rows.Count
                nrCols = codeCellRange.Columns.Count
                localisation = list()

                for index in range(codeColIndex, codeColIndex + nrCols):
                    if workSheet.Cells(codeRowIndex + nrLines, codeColIndex).Value == None:
                        pass
                    else:
                        list_code.append(workSheet.Cells(codeRowIndex + nrLines, index).Value)

            for element in list_table:
                if element in list_code:
                    localisation = None
                else:
                    localisation = ""
                    check = True
                    break


            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1055(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        list_measure = list()
        list_table = list()
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Applicabilité projet" or cell == "Project applicability":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 1:
            result(TSDApp.DOC9Dict["Test_02043_18_04939_WHOLENESS _1055"][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            nrCols = refCellRange.Columns.Count
            localisation = list()

            for index in range(refColIndex, refColIndex + nrCols):
                if workSheet.Cells(refRowIndex + nrLines, refColIndex).Value == None:
                    pass
                else:
                    list_table.append(workSheet.Cells(refRowIndex + nrLines, index).Value)

            if TSDApp.WorkbookStats.hasMeasure == False:
                result(TSDApp.DOC9Dict["Test_02043_18_04939_WHOLENESS _1055"][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
                workSheetRange = workSheet.UsedRange
                nrCols = workSheetRange.Columns.Count
                measureColIndex = 0
                var = 0
                for cellRow in workSheetRange.Value:
                    for cell in cellRow:
                        if cell == "Applicabilité projet" or cell == "Project applicability":
                            measureColIndex = cellRow.index(cell) + 1
                            measureRowIndex = workSheetRange.Value.index(cellRow) + 1
                            break
                    if measureColIndex != 0:
                        break

                measureCellRange = workSheet.Cells(measureRowIndex, measureColIndex).MergeArea
                nrLines = measureCellRange.Rows.Count
                nrCols = measureCellRange.Columns.Count
                localisation = list()


                for index in range(measureColIndex, measureColIndex + nrCols):
                    if workSheet.Cells(measureRowIndex + nrLines, measureColIndex).Value == None:
                        pass
                    else:
                        list_measure.append(workSheet.Cells(measureRowIndex + nrLines, index).Value)

            for element in list_table:
                if element in list_measure:
                    localisation = None
                else:
                    localisation = ""
                    check = True
                    break

            result(TSDApp.DOC9Dict["Test_02043_18_04939_WHOLENESS _1055"][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return check

def Test_02043_18_04939_WHOLENESS_1060(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Applicabilité projet" or cell == "Project applicability":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            nrCols = refCellRange.Columns.Count
            localisation = list()
            list_table = list()

            for index in range(refColIndex, refColIndex + nrCols):
                if workSheet.Cells(refRowIndex + nrLines, index).Value == "NA" or workSheet.Cells(refRowIndex + nrLines, index).Value == "X":
                    pass
                else:
                    localisation.append(workSheet.Cells(refRowIndex + nrLines, index))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1061(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Applicabilité projet" or cell == "Project applicability":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            nrCols = refCellRange.Columns.Count
            localisation = list()
            list_table = list()

            for index in range(refColIndex, refColIndex + nrCols):
                if workSheet.Cells(refRowIndex + nrLines, index).Value == "NA" or workSheet.Cells(refRowIndex + nrLines, index).Value == "X":
                    pass
                else :
                    localisation.append(workSheet.Cells(refRowIndex + nrLines, index))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return  check

def Test_02043_18_04939_WHOLENESS_1062(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check =True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Applicabilité projet" or cell == "Project applicability":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            nrCols = refCellRange.Columns.Count
            localisation = list()
            list_table = list()

            for index in range(refColIndex, refColIndex + nrCols):
                if workSheet.Cells(refRowIndex + nrLines, index).Value == "NA" or workSheet.Cells(refRowIndex + nrLines, index).Value == "X":
                    pass
                else :
                    localisation.append(workSheet.Cells(refRowIndex + nrLines, index))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1070(workBook, TSDApp):

    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
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
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == "NO DTC":
                    pass
                else:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1080(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
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
            elif refColIndex == 0:
                var = 1
                break
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1090(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "supporté par constituant (s)":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check =True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1100(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "supporté par constituant (s)":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1110(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "libellé (signification)":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1120(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Description de la strategie pour détecter le défaut":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1130(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Seuil de détection  /  valeur  du défaut ":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1140(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Temps de confirmation du défaut":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1150(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Description de la strategie de disparition du défaut / Procedure à effectuer pour vérifier la disparition du défaut":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1160(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Mode dégradé":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1170(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Voyant":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1180(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
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
            elif refColIndex == 0:
                var = 1
                break
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1190(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Défaillance constituant":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1200(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Situation de vie client":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check =True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1210(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
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
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check =True
    return check

def Test_02043_18_04939_WHOLENESS_1220(workBook, TSDApp):
    check = False
    testName = inspect.currentframe().f_code.co_name
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
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
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1230(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        refColIndex = 0
        var = 0

        for cellRow in workSheetRange.Value:
            for cell in cellRow:
                if cell == "Défaillance constituant":
                    refColIndex = cellRow.index(cell) + 1
                    refRowIndex = workSheetRange.Value.index(cellRow) + 1
                    break
            if refColIndex != 0:
                break
            elif refColIndex == 0:
                var = 1
                break

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if str(localisation) == "[]":
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check
