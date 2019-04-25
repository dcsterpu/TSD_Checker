import TSD_Checker_V3_1
import inspect
from ExcelEdit import TestReturn as result
from ErrorMessages import errorMessagesDict as error



def Test_02043_18_04939_WHOLENESS_1000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)

        if TSDApp.WorkbookStats.tableRefColIndex > 0:
            refCellRange = workSheet.Cells(TSDApp.WorkbookStats.tableRefRowIndex,TSDApp.WorkbookStats.tableRefColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            localisation = list()

            for index in range(TSDApp.WorkbookStats.tableRefRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, TSDApp.WorkbookStats.tableRefColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, TSDApp.WorkbookStats.tableRefColIndex))
                    check = True
            if not localisation:
                localisation = None
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check  = True
    return check

def Test_02043_18_04939_WHOLENESS_1001(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
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
                if str(cell.Value).casefold() == "Version".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = []

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1010(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)

        if TSDApp.WorkbookStats.codeRefColIndex > 0:
            refCellRange = workSheet.Cells(TSDApp.WorkbookStats.codeRefRowIndex, TSDApp.WorkbookStats.codeRefColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(TSDApp.WorkbookStats.codeRefRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, TSDApp.WorkbookStats.codeRefRowIndex).Value == None:
                    localisation.append(workSheet.Cells(index, TSDApp.WorkbookStats.codeRefColIndex))
                    check = True
            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1011(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        var = 0

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Version".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1020(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)

        if TSDApp.WorkbookStats.measureRefColIndex > 0:
            refCellRange = workSheet.Cells(TSDApp.WorkbookStats.measureRefRowIndex, TSDApp.WorkbookStats.measureRefColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = []

            for index in range(TSDApp.WorkbookStats.measureRefRowIndex + nrLines, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.Cells(index, TSDApp.WorkbookStats.measureRefColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, TSDApp.WorkbookStats.measureRefColIndex))
                    check = True
            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1021(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        var = 0

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Version".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = []

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1030(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagDebIndex)

        if TSDApp.WorkbookStats.DiagDebRefColIndex > 0:
            refCellRange = workSheet.Cells(TSDApp.WorkbookStats.DiagDebRefColIndex, TSDApp.WorkbookStats.DiagDebRefRowIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = []

            for index in range(TSDApp.WorkbookStats.DiagDebRefRowIndex + nrLines, TSDApp.WorkbookStats.DiagDebLastRow+1):
                if workSheet.Cells(index, TSDApp.WorkbookStats.DiagDebRefColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, TSDApp.WorkbookStats.DiagDebRefColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1031(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagDebIndex)
        refColIndex = 0
        var = 0

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Version".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

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
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMDD == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.MDDIndex)

        if TSDApp.WorkbookStats.MDDRefColIndex > 0:
            refCellRange = workSheet.Cells(TSDApp.WorkbookStats.MDDRefRowIndex, TSDApp.WorkbookStats.MDDRefColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(TSDApp.WorkbookStats.MDDRefRowIndex + nrLines, TSDApp.WorkbookStats.MDDLastRow):
                if workSheet.Cells(index, TSDApp.WorkbookStats.MDDRefColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, TSDApp.WorkbookStats.MDDRefColIndex))
                    check = True
            if not localisation:
                localisation = None
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1041(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMDD == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.MDDIndex)
        refColIndex = 0
        var = 0

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Version".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.MDDLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if not localisation:
                localisation = None
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1050(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        list_code = list()
        list_table = list()
        var = 0

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Applicabilité projet".casefold().strip() or str(cell.Value).casefold() == "Project applicability".casefold().strip():
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
                codeColIndex = 0
                var = 0
                for cellRow in workSheet.Rows:
                    for cell in cellRow.Cells:
                        if str(cell.Value).casefold() == "Applicabilité projet".casefold().strip() or str(cell.Value).casefold() == "Project applicability".casefold().strip():
                            codeColIndex = cell.Column
                            codeRowIndex = cell.Row
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
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        list_measure = list()
        list_table = list()
        var = 0

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Applicabilité projet".casefold().strip() or str(cell.Value).casefold() == "Project applicability".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

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
                measureColIndex = 0
                var = 0
                for cellRow in workSheet.Rows:
                    for cell in cellRow.Cells:
                        if str(cell.Value).casefold() == "Applicabilité projet".casefold().strip() or str(cell.Value).casefold() == "Project applicability".casefold().strip():
                            measureColIndex = cell.Column
                            measureRowIndex = cell.Row
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
    print(testName)
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
                if str(cell.Value).casefold() == "Applicabilité projet".casefold().strip() or str(cell.Value).casefold() == "Project applicability".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

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
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        var = 0

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Applicabilité projet".casefold().strip() or str(cell.Value).casefold() == "Project applicability".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

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
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check =True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        var = 0

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Applicabilité projet".casefold().strip() or str(cell.Value).casefold() == "Project applicability".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

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
    print(testName)
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
                if str(cell.Value).casefold() == "Code défaut".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = []

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == "NO DTC":
                    pass
                else:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1080(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        var = 0

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

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = []

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1090(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        var = 0

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "supporté par constituant (s)".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check =True
            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1100(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        var = 0

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "supporté par constituant (s)".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1110(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        var = 0

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "libellé (signification)".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1120(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        var = 0

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold() == "Description de la strategie pour détecter le défaut".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1130(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        var = 0

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold().strip() == "Seuil de détection  /  valeur  du défaut".casefold():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1140(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        var = 0

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(
                        cell.Value).casefold().strip() == "Temps de confirmation du défaut".casefold():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1150(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        var = 0

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(
                        cell.Value).casefold() == "Description de la strategie de disparition du défaut / Procedure à effectuer pour vérifier la disparition du défaut".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1160(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        var = 0

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(
                        cell.Value).casefold() == "Mode dégradé".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1170(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        var = 0

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(
                        cell.Value).casefold() == "Voyant".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1180(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
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
                if str(cell.Value).casefold() == "Constituant défaillant détecté".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1190(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
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
                if str(cell.Value).casefold() == "Défaillance constituant".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1200(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
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
                if str(cell.Value).casefold() == "Situation de vie client".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check =True
            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1210(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        var = 0
        refColIndex = 0
        refRowIndex = 0

        for cellRow in workSheet.Rows:
            for cell in cellRow.Cells:
                if str(cell.Value).casefold().strip() == "Effet(s) client(s)".casefold():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = []

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return check

def Test_02043_18_04939_WHOLENESS_1220(workBook, TSDApp):
    check = False
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        var = 0

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

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if not localisation:
                localisation = None
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1230(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
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
                if str(cell.Value).casefold() == "Défaillance constituant".casefold().strip():
                    refColIndex = cell.Column
                    refRowIndex = cell.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1240(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasInfGen == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.InfGenIndex)
        workSheetRange = workSheet.UsedRange
        refColIndex = 0
        refRowIndex = 0
        refColIndex1 = 0
        refRowIndex1 = 0
        var = 0

        for row in workSheetRange:
            for cel in row:
                if "Liste de diffusion".casefold() in str(cel.Value).casefold().strip() or "Mailing list".casefold() in str(cel.Value).casefold().strip():
                    refColIndex = cel.Column
                    refRowIndex = cel.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            var = 0
            for row in workSheetRange:
                for cel in row:
                    if "Diffusion à :".casefold() in str(cel.Value).casefold().strip():
                        refColIndex1 = cel.Column
                        refRowIndex1 = cel.Row
                        break
                if refColIndex1 != 0:
                    break
            if refColIndex1 == 0:
                var = 1

        localisation = []
        if workSheet.Cells(refRowIndex1+1, refColIndex1).Value is not None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], None, workBook, TSDApp)
            check = True
        else:
            localisation.append(workSheet.Cells(refRowIndex1+1, refColIndex1))
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook,TSDApp)
    return check