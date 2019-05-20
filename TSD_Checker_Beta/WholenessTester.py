import TSD_Checker_V3_1_sans_limites
import inspect
from ExcelEdit import TestReturn as result
import win32timezone
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

            for index in range(TSDApp.WorkbookStats.tableRefRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, TSDApp.WorkbookStats.tableRefColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, TSDApp.WorkbookStats.tableRefColIndex))
                    check = True
            if not localisation:
                localisation = None
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Version".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = []

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))
                    check = True
            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        elif var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisation, workBook, TSDApp)
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

            for index in range(TSDApp.WorkbookStats.codeRefRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow + 1):
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Version".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow + 1):
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

            for index in range(TSDApp.WorkbookStats.measureRefRowIndex + nrLines, TSDApp.WorkbookStats.measureLastRow + 1):
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Version".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = []

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.measureLastRow + 1):
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

            for index in range(TSDApp.WorkbookStats.DiagDebRefRowIndex + nrLines, TSDApp.WorkbookStats.DiagDebLastRow + 1):
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DiagDebLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Version".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.DiagDebLastRow + 1):
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

            for index in range(TSDApp.WorkbookStats.MDDRefRowIndex + nrLines, TSDApp.WorkbookStats.MDDLastRow + 1):
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.MDDLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Version".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.MDDLastRow + 1):
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
        refRowIndex = 0
        list_code = list()
        list_table = list()
        var = 0
        localisation = list()

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Applicabilité projet".casefold() or str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Project applicability".casefold():
                    refColIndex = index2
                    refRowIndex = index1
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
                for index1 in range(1, 15):
                    for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                        if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Applicabilité projet".casefold() or str(workSheet.Cells(index1,index2).Value).casefold().strip() == "Project applicability".casefold():
                            codeColIndex = index2
                            codeRowIndex = index1
                            break
                    if codeColIndex != 0:
                        break

                codeCellRange = workSheet.Cells(codeRowIndex, codeColIndex).MergeArea
                nrLines = codeCellRange.Rows.Count
                nrCols = codeCellRange.Columns.Count


                for index in range(codeColIndex, codeColIndex + nrCols):
                    if workSheet.Cells(codeRowIndex + nrLines, codeColIndex).Value == None:
                        pass
                    else:
                        list_code.append(workSheet.Cells(codeRowIndex + nrLines, index).Value)

        flag = True
        if len(list_table) > 0 and len(list_code) > 0:
            for index in range(1,len(list_table) + 1):
                if list_table[index] == list_code[index]:
                    pass
                else:
                    flag = False
                    break

        if flag == True:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisation, workBook, TSDApp)
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook,TSDApp)
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
        refRowIndex = 0
        list_measure = list()
        list_table = list()
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Applicabilité projet".casefold() or str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Project applicability".casefold():
                    refColIndex = index2
                    refRowIndex = index1
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
                for index1 in range(1, 15):
                    for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                        if str(workSheet.Cells(index1,
                                               index2).Value).casefold().strip() == "Applicabilité projet".casefold() or str(
                                workSheet.Cells(index1,
                                                index2).Value).casefold().strip() == "Project applicability".casefold():
                            measureColIndex = index2
                            measureRowIndex = index1
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1,
                                       index2).Value).casefold().strip() == "Applicabilité projet".casefold() or str(
                        workSheet.Cells(index1, index2).Value).casefold().strip() == "Project applicability".casefold():
                    refColIndex = index2
                    refRowIndex = index1
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
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisation, workBook, TSDApp)
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Applicabilité projet".casefold() or str(
                        workSheet.Cells(index1, index2).Value).casefold().strip() == "Project applicability".casefold():
                    refColIndex = index2
                    refRowIndex = index1
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1,
                                       index2).Value).casefold().strip() == "Applicabilité projet".casefold() or str(
                        workSheet.Cells(index1, index2).Value).casefold().strip() == "Project applicability".casefold():
                    refColIndex = index2
                    refRowIndex = index1
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Code défaut".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = []

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Code défaut".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = []

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow + 1):
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "supporté par constituant (s)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow + 1):
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "supporté par constituant (s)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.measureLastRow + 1):
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "libellé (signification)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow + 1):
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Description de la strategie pour détecter le défaut".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow + 1):
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Seuil de détection  /  valeur  du défaut".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow + 1):
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Temps de confirmation du défaut".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow + 1):
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Description de la strategie de disparition du défaut / Procedure à effectuer pour vérifier la disparition du défaut".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow + 1):
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Mode dégradé".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow + 1):
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Voyant".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow + 1):
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Constituant défaillant détecté".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Défaillance constituant".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Situation de vie client".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Effet(s) client(s)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = list()
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count


            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value == None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisation, workBook, TSDApp)
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Code défaut".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Défaillance constituant".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
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
        for row in workSheetRange.Rows:
            for cel in row.Cells:
                if "Diffusion à :".casefold() in str(cel.Value).casefold().strip() or "E-mail to :".casefold() in str(cel.Value).casefold().strip():
                    refColIndex1 = cel.Column
                    refRowIndex1 = cel.Row
                    break
            if refColIndex1 != 0:
                break
        if refColIndex1 == 0:
            var = 1

        localisation = []

        if var == 0:
            if workSheet.Cells(refRowIndex1+1, refColIndex1).Value is not None:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], None, workBook, TSDApp)
                check = True
            else:
                localisation.append(workSheet.Cells(refRowIndex1+1, refColIndex1))
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook,TSDApp)
        else:
            localisation.append(workSheet.Cells(refRowIndex1, refColIndex1))
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook,TSDApp)
    return check

def Test_02043_18_04939_WHOLENESS_1300(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        localisation = []
        refCellRange = workSheet.Cells(TSDApp.WorkbookStats.tableRefRowIndex, TSDApp.WorkbookStats.tableRefColIndex).MergeArea
        nrLines = refCellRange.Rows.Count

        for index in range(nrLines + TSDApp.WorkbookStats.tableRefRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
            if workSheet.Cells(index, TSDApp.WorkbookStats.tableRefColIndex).Value is None:
                localisation.append(workSheet.Cells(index, TSDApp.WorkbookStats.tableRefColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1301(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Version".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1302(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "To diagnose".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1303(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Supplier system".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1304(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Logical flow".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1305(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Physical flow".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1306(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Client system".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1307(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Type of connection".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1308(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Type".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1309(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Logical failure mode".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1310(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Physical failure mode".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1311(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Wiring harness cause".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1312(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Other cause".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1313(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Operation situation / Scenario".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1314(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "system effect".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1315(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1 ):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Customer effect".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1316(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Comment".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1317(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Feared event".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1318(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Reference".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1319(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Severity".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1320(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Level".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1321(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "target".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1322(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Safety measure (G4) / Functional diagnostic(G3,G2,G1)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1323(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Type of failure".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1324(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Degraded mode /Safe state".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1325(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "lead time".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1326(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Efficiency".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1327(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "recovering mode".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1328(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Requirement N° to the Design Document".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1329(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Requirement N° from Design document".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1330(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "research time allocated to the system (in minutes)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1331(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol+ 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "HMI\n(Indicators/messages)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1332(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "High level test".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1333(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Diagnosis needs".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1334(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Comments".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1350(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DiagNeedsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Reference".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1351(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DiagNeedsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Version".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1352(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DiagNeedsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Label".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1353(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DiagNeedsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Description".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1354(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DiagNeedsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Situation during which the diagnosis is active".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1355(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DiagNeedsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Technical Effect covers by the need".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1356(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DiagNeedsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Diversity".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1357(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DiagNeedsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Allocated to the system".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1358(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DiagNeedsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Upstream requirements".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1359(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DiagNeedsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Taken into account".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1360(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DiagNeedsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "comment".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1361(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DiagNeedsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Project applicability".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1400(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.EffClientsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Name".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.EffClientsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1401(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.EffClientsLastRow + 15):
            for index2 in range(1, TSDApp.WorkbookStats.EffClientsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Taken into account".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.EffClientsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1402(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.EffClientsLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.EffClientsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Diagnosticability synthesis".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.EffClientsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1403(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.EffClientsLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.EffClientsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Comments".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.EffClientsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1430(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasReqTech == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReqTechIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.ReqTechLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.ReqTechLastcol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Reference".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ReqTechLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1431(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasReqTech == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReqTechIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.ReqTechLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.ReqTechLastcol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "version".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ReqTechLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check
def Test_02043_18_04939_WHOLENESS_1432(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasReqTech == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReqTechIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.ReqTechLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.ReqTechLastcol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Description".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ReqTechLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1433(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasReqTech == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReqTechIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.ReqTechLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.ReqTechLastcol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "technical effect".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ReqTechLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1434(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasReqTech == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReqTechIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.ReqTechLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.ReqTechLastcol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Allocated to".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ReqTechLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1435(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasReqTech == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReqTechIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.ReqTechLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.ReqTechLastcol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Tracability with the TSD".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ReqTechLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1450(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.FearedEventLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.FearedEventLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Description".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.FearedEventLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1451(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.FearedEventLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.FearedEventLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Reference".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.FearedEventLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1452(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.FearedEventLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.FearedEventLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Severity".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.FearedEventLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1453(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.FearedEventLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.FearedEventLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Level".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.FearedEventLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1454(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.FearedEventLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.FearedEventLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Taken into account".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.FearedEventLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1455(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.FearedEventLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.FearedEventLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Justification for not taking into account the dread Event".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.FearedEventLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1456(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.FearedEventLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.FearedEventLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Commentaire".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.FearedEventLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1500(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasSystem == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.SystemIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.SystemLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.SystemLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Description".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.SystemLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1501(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasSystem == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.SystemIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.SystemLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.SystemLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Taken into account".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.SystemLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1550(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasOpSit == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.OpSitIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.OpSitLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.OpSitLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Description".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.OpSitLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1551(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasOpSit == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.OpSitIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.OpSitLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.OpSitLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Taken into account".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.OpSitLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1552(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasOpSit == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.OpSitIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.OpSitLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.OpSitLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Comments".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.OpSitLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1600(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        localisation = []
        refCellRange = workSheet.Cells(TSDApp.WorkbookStats.tableRefRowIndex, TSDApp.WorkbookStats.tableRefColIndex).MergeArea
        nrLines = refCellRange.Rows.Count

        for index in range(nrLines + TSDApp.WorkbookStats.tableRefRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
            if workSheet.Cells(index, TSDApp.WorkbookStats.tableRefColIndex).Value is None:
                localisation.append(workSheet.Cells(index, TSDApp.WorkbookStats.tableRefColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1601(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Version".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1602(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Réf doc".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1603(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Variante/\noption".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1604(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Version de soft (MOTEUR / BSI,...)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1605(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "sous Fonction de conception incriminée".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1606(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Groupe de constituant".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1607(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Constituant défaillant détecté".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1608(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Flux fonctionnel".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1609(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Défaillance logique".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1610(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Défaillance constituant".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1611(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "PPM réparties".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1612(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "poids".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1613(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Situation de vie client".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1614(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Situation de vie détaillée".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1615(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "lien vers autre TSD".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1616(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Effet(s) client(s)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1617(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Evenement(s) redouté(s) (ER)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1618(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Voyant(s) ou \nmessage(s)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1619(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Code défaut".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1620(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Code défauts induits".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1621(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "mesures et commandes (Mesure Parametre et Test Actionneur) / Tests de cohérence".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1622(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Critère de décision".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1623(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "DIAGNOSTIC DEBARQUE".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1624(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Critère de decision".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1625(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Action sur constituant incriminé".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1626(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Statut réunion DSP-DRD".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1627(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Action a réaliser / Commentaires".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1628(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Référence AMDEC".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1629(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Justification de la modification".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1630(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Validation".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1631(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Controle Usine".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1632(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Pris en compte dans logigramme".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1650(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        localisation = []
        refCellRange = workSheet.Cells(TSDApp.WorkbookStats.codeRefRowIndex, TSDApp.WorkbookStats.codeRefColIndex).MergeArea
        nrLines = refCellRange.Rows.Count

        for index in range(nrLines + TSDApp.WorkbookStats.codeRefRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
            if workSheet.Cells(index, TSDApp.WorkbookStats.codeRefColIndex).Value is None:
                localisation.append(workSheet.Cells(index, TSDApp.WorkbookStats.codeRefColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1651(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Version".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1652(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Code défaut".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1653(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "libellé (signification)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1654(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Flux Fonctionnel".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1655(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Description de la strategie pour détecter le défaut".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1656(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Seuil de détection  /  valeur  du défaut".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1657(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Temps de confirmation du défaut".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1658(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Description de la strategie de disparition du défaut / Procedure à effectuer pour vérifier la disparition du défaut".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1659(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Situation de vie véhicule pour faire remonter le code défaut".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1660(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Mode dégradé".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1661(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Taux de remonté du code défaut".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1662(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Voyant".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1663(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Accès scantool".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1664(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Groupe de contextes associés".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1684(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Diversité".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1685(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Applicabilité usine".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1686(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "condition d'applicabilité en usine".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1687(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "supporté par constituant (s)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1688(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "se référer au document spécifiant DRD : (réf & version)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1689(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Référence amont".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1690(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Version de la référence amont".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1691(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Pris en compte".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1692(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Justification de la modification".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1693(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Validation".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1700(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        localisation = []
        refCellRange = workSheet.Cells(TSDApp.WorkbookStats.measureRefRowIndex, TSDApp.WorkbookStats.measureRefColIndex).MergeArea
        nrLines = refCellRange.Rows.Count

        for index in range(nrLines + TSDApp.WorkbookStats.measureRefRowIndex, TSDApp.WorkbookStats.measureLastRow + 1):
            if workSheet.Cells(index, TSDApp.WorkbookStats.measureRefColIndex).Value is None:
                localisation.append(workSheet.Cells(index, TSDApp.WorkbookStats.measureRefColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1701(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Version".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.measureLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1702(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Type (choix par menu)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.measureLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1703(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "libellé (signification)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.measureLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1704(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Description".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.measureLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1705(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Situation pendant laquelle la mesure ou commande est utilisable".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.measureLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1706(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Statut".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.measureLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1707(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Taux de fiabilité du test (50%, 100%)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.measureLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1708(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Flux fonctionnel".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.measureLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1709(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Uniquement \npour O Control\nlecture \nsortie effective /commande".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.measureLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check


def Test_02043_18_04939_WHOLENESS_1710(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Diversité".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.measureLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check


def Test_02043_18_04939_WHOLENESS_1711(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Applicabilité usine".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.measureLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1712(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "condition d'applicabilité en usine".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.measureLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1713(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "supporté par constituant (s)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.measureLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1714(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "se référer au document spécifiant DRD : (réf & version)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.measureLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1715(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Référence amont".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.measureLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1716(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Version de la référence amont".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.measureLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1717(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Pris en compte".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.measureLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1718(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Justification de la modification".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.measureLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1719(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Validation".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.measureLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1750(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagDebIndex)
        localisation = []
        refCellRange = workSheet.Cells(TSDApp.WorkbookStats.DiagDebRefRowIndex, TSDApp.WorkbookStats.DiagDebRefColIndex).MergeArea
        nrLines = refCellRange.Rows.Count

        for index in range(nrLines + TSDApp.WorkbookStats.DiagDebRefRowIndex, TSDApp.WorkbookStats.DiagDebLastRow + 1):
            if workSheet.Cells(index, TSDApp.WorkbookStats.DiagDebRefColIndex).Value is None:
                localisation.append(workSheet.Cells(index, TSDApp.WorkbookStats.DiagDebRefColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1751(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagDebIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DiagDebLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Version".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DiagDebLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1752(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagDebIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DiagDebLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "libellé (signification)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DiagDebLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1753(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagDebIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DiagDebLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Description".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DiagDebLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1754(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagDebIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DiagDebLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Taux de fiabilité du test (50%, 100%)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DiagDebLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1755(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagDebIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DiagDebLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Applicabilité Usine".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DiagDebLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1756(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagDebIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DiagDebLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "se référer au document spécifiant : (réf & version)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DiagDebLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1757(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagDebIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DiagDebLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Pris en compte".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DiagDebLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1758(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagDebIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DiagDebLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Justification de la modification".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DiagDebLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1759(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagDebIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DiagDebLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Validation".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DiagDebLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1800(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.EffClientsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Noms".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.EffClientsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1801(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.EffClientsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Pris en compte".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.EffClientsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1802(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.EffClientsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Synthèse de la diagnosticabilité".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.EffClientsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1803(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.EffClientsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Justification de la modification".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.EffClientsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1810(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasER == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ERIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.ERLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "nom".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ERLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1811(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasER == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ERIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.ERLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "désignation".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ERLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1812(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasER == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ERIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.ERLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Gravité".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ERLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1813(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasER == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ERIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.ERLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Pris en compte".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ERLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1814(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasER == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ERIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.ERLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Justification de non prise en compte de l'ER".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ERLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1815(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasER == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ERIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.ERLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Justification de la modification".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ERLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1820(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasConstituants == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.constituantsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.constituantsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Noms".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.constituantsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1821(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasConstituants == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.constituantsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.constituantsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Description".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.constituantsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1822(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasConstituants == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.constituantsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.constituantsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Taux de défaillance (en ppm)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.constituantsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1823(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasConstituants == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.constituantsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.constituantsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Découpage PSA".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.constituantsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1824(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasConstituants == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.constituantsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.constituantsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Pris en compte".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.constituantsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1825(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasConstituants == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.constituantsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.constituantsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Justification de la modification".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.constituantsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1830(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasSitDeVie == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.SitDeVieIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.SitDeVieLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Situations de vie".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.SitDeVieLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1831(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasSitDeVie == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.SitDeVieIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.SitDeVieLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Justification de la modification".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.SitDeVieLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1840(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMDD == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.MDDIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.MDDLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Modes dégradés:".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.MDDLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1841(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasMDD == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.MDDIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.MDDLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Justification de la modification".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.MDDLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1900(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        localisation = []
        refCellRange = workSheet.Cells(TSDApp.WorkbookStats.tableRefRowIndex, TSDApp.WorkbookStats.tableRefColIndex).MergeArea
        nrLines = refCellRange.Rows.Count

        for index in range(nrLines + TSDApp.WorkbookStats.tableRefRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
            if workSheet.Cells(index, TSDApp.WorkbookStats.tableRefColIndex).Value is None:
                localisation.append(workSheet.Cells(index, TSDApp.WorkbookStats.tableRefColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1901(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Version".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1902(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Document of reference".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1903(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Variant/\noption".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1904(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Sub-function of the system incriminated".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1905(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Module / Group of parts".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1906(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Defective part".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1907(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Contribution to fonctionnality".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1908(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Logical failure mode".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1909(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Physical failure mode".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1910(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Weight".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1911(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Situation".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1912(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Detailed situation".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check


def Test_02043_18_04939_WHOLENESS_1913(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Link to another DST".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1914(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Technical effect".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1915(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Customer effect".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1916(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Feared events".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1917(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Degraded mode".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1918(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "HMI\n(Indicator lights/messages)".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1919(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Data Trouble code".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1920(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow  + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Mislead Data trouble code".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1921(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Read data or I/O control".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1922(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "decision criterion".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1923(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Non-embedded diagnosis".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1924(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "decision criterion".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1925(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Action on the incriminated part".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1926(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol  + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "to do list / Comments".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1927(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol  + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "FMEA reference".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1950(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        localisation = []
        refCellRange = workSheet.Cells(TSDApp.WorkbookStats.tableRefRowIndex, TSDApp.WorkbookStats.tableRefColIndex).MergeArea
        nrLines = refCellRange.Rows.Count

        for index in range(nrLines + TSDApp.WorkbookStats.codeRefRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
            if workSheet.Cells(index, TSDApp.WorkbookStats.codeRefColIndex).Value is None:
                localisation.append(workSheet.Cells(index, TSDApp.WorkbookStats.codeRefColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1951(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Version".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1952(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Data trouble code".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1953(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Label".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1954(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Description of the qualification conditions".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1955(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Detection threshold".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1956(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Qualification time".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1957(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Description of the dequalification conditions / Operation to do to check if the defect disappeared".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1958(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Conditions of the diagnostic activation".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1959(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Degraded mode".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1960(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Failure detection rate".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1961(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Indicateur light".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1962(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Visibility of the failure with the Scantool".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1963(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Freeze Frame Class".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1964(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Diversity".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1965(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Stored by the ECU".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1966(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Upstream requirements".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1967(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Taken into account".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1968(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "projet X".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_1969(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Projet Y".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.ReadDataIOLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Reference".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2001(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.ReadDataIOLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Version".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2002(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.ReadDataIOLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Type of diagnosis".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2003(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.ReadDataIOLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Label".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2004(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.ReadDataIOLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Description".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2005(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.ReadDataIOLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Conditions of the diagnostic activation".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2006(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.ReadDataIOLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Status".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2007(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.ReadDataIOLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Diversity".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2008(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.ReadDataIOLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Stored by the ECU".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2009(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.ReadDataIOLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Upstream requirements".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2010(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.ReadDataIOLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Taken into account".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check


def Test_02043_18_04939_WHOLENESS_2011(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.ReadDataIOLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "project X".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.ReadDataIOLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2050(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasNotEmbDiag == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.NotEmbDiagIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.NotEmbDiagLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.NotEmbDiagLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Reference".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.NotEmbDiagLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2051(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasNotEmbDiag == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.NotEmbDiagIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.NotEmbDiagLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.NotEmbDiagLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Version".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.NotEmbDiagLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2052(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasNotEmbDiag == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.NotEmbDiagIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.NotEmbDiagLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.NotEmbDiagLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Label".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.NotEmbDiagLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2053(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasNotEmbDiag == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.NotEmbDiagIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.NotEmbDiagLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.NotEmbDiagLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Description".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.NotEmbDiagLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2054(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasNotEmbDiag == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.NotEmbDiagIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.NotEmbDiagLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.NotEmbDiagLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Upstream requirements".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.NotEmbDiagLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2055(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasNotEmbDiag == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.NotEmbDiagIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.NotEmbDiagLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.NotEmbDiagLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Taken into account".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.NotEmbDiagLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2056(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasNotEmbDiag == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.NotEmbDiagIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.NotEmbDiagLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.NotEmbDiagLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "projet X".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.NotEmbDiagLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2060(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTechEff == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.TechEffIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.TechEffLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.TechEffLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Name".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.TechEffLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2061(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTechEff == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.TechEffIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.TechEffLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.TechEffLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Taken into account".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.TechEffLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2062(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasTechEff == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.TechEffIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.TechEffLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.TechEffLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Upstream requirements".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.TechEffLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2070(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.EffClientsLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.EffClientsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Name".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.EffClientsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2071(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.EffClientsLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.EffClientsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Taken into account".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.EffClientsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2072(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.EffClientsLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.EffClientsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Diagnosticability synthesis".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.EffClientsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2080(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.FearedEventLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.FearedEventLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Description".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.FearedEventLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2081(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.FearedEventLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.FearedEventLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Reference".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.FearedEventLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2082(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.FearedEventLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.FearedEventLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Severity".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.FearedEventLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2083(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.FearedEventLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.FearedEventLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Taken into account".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.FearedEventLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2084(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.FearedEventLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.FearedEventLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Justification for not taking into account the dread Event".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.FearedEventLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2090(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasParts == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.PartsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.PartsLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.PartsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Name".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.PartsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2091(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasParts == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.PartsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.PartsLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.PartsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Description".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.PartsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2092(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasParts == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.PartsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.PartsLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.PartsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Taken into account".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.PartsLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2100(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasVariant == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.VariantIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.VariantLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.VariantLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Name".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.VariantLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2101(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasVariant == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.VariantIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.VariantLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.VariantLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Description".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.VariantLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2102(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasVariant == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.VariantIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.VariantLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.VariantLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Taken into account".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.VariantLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2110(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasSituation == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.SituationIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.SituationLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.SituationLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Description".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.SituationLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2111(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasSituation == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.SituationIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.SituationLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.SituationLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Taken into account".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.SituationLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2112(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasSituation == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.SituationIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.SituationLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.SituationLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Comments".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.SituationLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2120(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDegradedMode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DegradedModeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.DegradedModeLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.DegradedModeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Modes dégradés:".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DegradedModeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def Test_02043_18_04939_WHOLENESS_2121(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    check = False
    if TSDApp.WorkbookStats.hasDegradedMode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DegradedModeIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0

        for index1 in range(1, TSDApp.WorkbookStats.DegradedModeLastRow + 1):
            for index2 in range(1, TSDApp.WorkbookStats.DegradedModeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Taken into account:".casefold():
                    refColIndex = index2
                    refRowIndex = index1
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        localisation = []
        if var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            for index in range(nrLines + refRowIndex, TSDApp.WorkbookStats.DegradedModeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    localisation.append(workSheet.Cells(index, refColIndex))

        if not localisation:
            localisation = None

        if localisation is None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check