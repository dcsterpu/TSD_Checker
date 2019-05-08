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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.measureLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.DiagDebLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.DiagDebLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.MDDLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.MDDLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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
                for index1 in range(1, TSDApp.WorkbookStats.codeLastRow):
                    for index2 in range(1, TSDApp.WorkbookStats.codeLastCol):
                        if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Applicabilité projet".casefold() or str(workSheet.Cells(index1,index2).Value).casefold().strip() == "Project applicability".casefold():
                            codeColIndex = index2
                            codeRowIndex = index1
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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
                for index1 in range(1, TSDApp.WorkbookStats.measureLastRow):
                    for index2 in range(1, TSDApp.WorkbookStats.measureLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.measureLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.measureLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.codeLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for row in workSheetRange.Rows:
            for cel in row.Cells:
                if "Liste de diffusion".casefold() in str(cel.Value).casefold().strip() or "Mailing list (the taking part)".casefold() in str(cel.Value).casefold().strip():
                    refColIndex = cel.Column
                    refRowIndex = cel.Row
                    break
            if refColIndex != 0:
                break
        if refColIndex == 0:
            var = 1

        if var == 0:
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
        if workSheet.Cells(refRowIndex1+1, refColIndex1).Value is not None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], None, workBook, TSDApp)
            check = True
        else:
            localisation.append(workSheet.Cells(refRowIndex1+1, refColIndex1))
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.tableLastRow):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol):
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

        for index1 in range(1, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
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

        for index1 in range(1, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
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

        for index1 in range(1, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
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

        for index1 in range(1, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
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

        for index1 in range(1, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
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

        for index1 in range(1, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
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

        for index1 in range(1, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
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

        for index1 in range(1, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
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

        for index1 in range(1, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
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

        for index1 in range(1, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
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

        for index1 in range(1, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
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

        for index1 in range(1, TSDApp.WorkbookStats.DiagNeedsLastRow + 1):
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
