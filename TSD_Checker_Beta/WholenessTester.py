import TSD_Checker_V6_8
import inspect
from ExcelEdit import TestReturn as result
from ExcelEdit import TestReturnName as show
from ErrorMessages import errorMessagesDict as error



def Test_02043_18_04939_WHOLENESS_1000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Référence".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, refColIndex).value == None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("tableau", index, refColIndex))
                    check = True
            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook, TSDApp)
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Version".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, refColIndex).value == None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("tableau", index, refColIndex))
                    check = True
            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
        else:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Référence".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, refColIndex).value == None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("codes défauts", index, refColIndex))
                    check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,TSDApp)
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Version".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, refColIndex).value == None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("codes défauts", index, refColIndex))
                    check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
        else:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "Référence".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("mesures et commandes", index, refColIndex))
                    check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "Version".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("mesures et commandes", index, refColIndex))
                    check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
        else:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagDebIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagDebLastCol):
            if str(workSheet.cell(TSDApp.diagDebHeaderRow, index).value).casefold().strip() == "Référence".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagDebFirstInfoRow, TSDApp.WorkbookStats.DiagDebLastRow):
                if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("Diagnostic débarqués", index, refColIndex))
                    check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1031(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagDebIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagDebLastCol):
            if str(workSheet.cell(TSDApp.diagDebHeaderRow, index).value).casefold().strip() == "Version".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagDebFirstInfoRow, TSDApp.WorkbookStats.DiagDebLastRow):
                if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("Diagnostic débarqués", index, refColIndex))
                    check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
        else:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.MDDIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.MDDLastCol):
            if str(workSheet.cell(TSDApp.listeMDDHeaderRow, index).value).casefold().strip() == "Référence".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.listeMDDFirstInfoRow, TSDApp.WorkbookStats.MDDLastRow):
                if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("Diagnostic débarqués", index, refColIndex))
                    check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.MDDIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.MDDLastCol):
            if str(workSheet.cell(TSDApp.listeMDDHeaderRow, index).value).casefold().strip() == "Version".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.listeMDDFirstInfoRow, TSDApp.WorkbookStats.MDDLastRow):
                if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("Diagnostic débarqués", index, refColIndex))
                    check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
        else:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        name = []
        contor = 0
        table_project = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow - 1, index).value).casefold().strip() == "Applicabilité projet".casefold() or str(workSheet.cell(TSDApp.tableHeaderRow - 1, index).value).casefold().strip() == "Project applicability".casefold():
                refColIndex = index
                break


        if refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif refColIndex != -1:
            for index in range(refColIndex, TSDApp.WorkbookStats.tableLastCol):
                if workSheet.cell(TSDApp.tableHeaderRow + 1, index).value is not None and workSheet.cell(TSDApp.tableHeaderRow + 1, index).value != "":
                    table_project.append(workSheet.cell(TSDApp.tableHeaderRow + 1, index).value.strip())

            if TSDApp.WorkbookStats.hasCode == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
                codeColIndex = -1

                for index in range(0, TSDApp.WorkbookStats.codeLastCol):
                    if str(workSheet.cell(TSDApp.codeHeaderRow - 1,index).value).casefold().strip() == "Applicabilité projet".casefold() or str(workSheet.cell(TSDApp.codeHeaderRow - 1,index).value).casefold().strip() == "Project applicability".casefold():
                        codeColIndex = index
                        break

                if codeColIndex != -1:
                    code_project = []
                    for index in range(codeColIndex, TSDApp.WorkbookStats.codeLastCol):
                        if workSheet.cell(TSDApp.codeHeaderRow, index).value is not None and workSheet.cell(TSDApp.codeHeaderRow, index).value != "":
                            code_project.append(workSheet.cell(TSDApp.codeHeaderRow, index).value.strip())

                    if len(table_project) == len(code_project):
                        for project in table_project:
                            if project in code_project:
                                contor += 1
                                break

                    if contor == len(table_project):
                        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook,TSDApp)
                        check = True
                    else:
                        name.append("Different projects!")
                        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook,TSDApp)
                else:
                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    return check
    # testName = inspect.currentframe().f_code.co_name
    # print(testName)
    # check = False
    # if TSDApp.WorkbookStats.hasTable == False:
    #     result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    #     check = True
    # else:
    #     workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
    #     refColIndex = 0
    #     refRowIndex = 0
    #     list_code = list()
    #     list_table = list()
    #     var = 0
    #     name = []
    #     contor = 0
    #     table_project = []
    #
    #     for index1 in range(1, 15):
    #         for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
    #             if str(workSheet.cell(index1, index2).value).casefold().strip() == "Applicabilité projet".casefold() or str(workSheet.cell(index1, index2).value).casefold().strip() == "Project applicability".casefold():
    #                 refColIndex = index2
    #                 refRowIndex = index1
    #                 break
    #         if refColIndex != 0:
    #             break
    #     if refColIndex == 0:
    #         var = 1
    #
    #     if var == 1:
    #         result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    #         check = True
    #     elif var == 0:
    #         refCellRange = workSheet.cell(refRowIndex, refColIndex).MergeArea
    #         nrLines = refCellRange.Rows.Count
    #         nrCols = refCellRange.Columns.Count
    #
    #         for index in range(refColIndex, TSDApp.WorkbookStats.tableLastCol + 1):
    #             if workSheet.cell(refRowIndex + nrLines, index).Borders(8).LineStyle != -4142 and workSheet.cell(refRowIndex + nrLines, index).value != None:
    #                 table_project.append(workSheet.cell(refRowIndex + nrLines, index).value.strip())
    #
    #         if TSDApp.WorkbookStats.hasCode == False:
    #             result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    #             check = True
    #         else:
    #             workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
    #             codeColIndex = 0
    #             var = 0
    #             for index1 in range(1, 15):
    #                 for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
    #                     if str(workSheet.cell(index1, index2).value).casefold().strip() == "Applicabilité projet".casefold() or str(workSheet.cell(index1,index2).value).casefold().strip() == "Project applicability".casefold():
    #                         codeColIndex = index2
    #                         codeRowIndex = index1
    #                         break
    #                 if codeColIndex != 0:
    #                     break
    #
    #             codeCellRange = workSheet.cell(codeRowIndex, codeColIndex).MergeArea
    #             nrLines = codeCellRange.Rows.Count
    #             nrCols = codeCellRange.Columns.Count
    #
    #             code_project = []
    #             for index in range(codeColIndex, TSDApp.WorkbookStats.codeLastCol + 1):
    #                 if workSheet.cell(codeRowIndex + nrLines, index).Borders(8).LineStyle != -4142 and workSheet.cell(codeRowIndex + nrLines, index).value != None:
    #                     code_project.append(workSheet.cell(codeRowIndex + nrLines, index).value.strip())
    #
    #
    #             if len(table_project) == len(code_project):
    #                 for project in table_project:
    #                     if project in code_project:
    #                         contor += 1
    #                         break
    #
    #     if contor == len(table_project):
    #         show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    #         check = True
    #     else:
    #         name.append("Different projects!")
    #         show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook,TSDApp)
    # return check

def Test_02043_18_04939_WHOLENESS_1055(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        name = []
        contor = 0
        table_project = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow - 1,
                                  index).value).casefold().strip() == "Applicabilité projet".casefold() or str(
                    workSheet.cell(TSDApp.tableHeaderRow - 1,
                                   index).value).casefold().strip() == "Project applicability".casefold():
                refColIndex = index
                break

        if refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif refColIndex != -1:
            for index in range(refColIndex, TSDApp.WorkbookStats.tableLastCol):
                if workSheet.cell(TSDApp.tableHeaderRow + 1, index).value is not None and workSheet.cell(
                        TSDApp.tableHeaderRow + 1, index).value != "":
                    table_project.append(workSheet.cell(TSDApp.tableHeaderRow + 1, index).value.strip())

            if TSDApp.WorkbookStats.hasMeasure == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
                measureColIndex = -1

                for index in range(0, TSDApp.WorkbookStats.measureLastCol):
                    if str(workSheet.cell(TSDApp.measureHeaderRow - 1,
                                          index).value).casefold().strip() == "Applicabilité projet".casefold() or str(
                            workSheet.cell(TSDApp.measureHeaderRow - 1,
                                           index).value).casefold().strip() == "Project applicability".casefold():
                        measureColIndex = index
                        break

                if measureColIndex != -1:
                    measure_project = []
                    for index in range(measureColIndex, TSDApp.WorkbookStats.measureLastCol):
                        if workSheet.cell(TSDApp.measureHeaderRow, index).value != None and workSheet.cell(
                                TSDApp.measureHeaderRow, index).value != "":
                            measure_project.append(workSheet.cell(TSDApp.measureHeaderRow, index).value.strip())

                    if len(table_project) == len(measure_project):
                        for project in table_project:
                            if project in measure_project:
                                contor += 1
                                break

                    if contor == len(table_project):
                        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook,
                             TSDApp)
                        check = True
                    else:
                        name.append("Different projects!")
                        show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook,
                             TSDApp)
                else:
                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    return check
    # testName = inspect.currentframe().f_code.co_name
    # print(testName)
    # check = False
    # if TSDApp.WorkbookStats.hasTable == False:
    #     result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    #     check = True
    # else:
    #     workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
    #     refColIndex = 0
    #     refRowIndex = 0
    #     list_code = list()
    #     list_table = list()
    #     var = 0
    #     name = []
    #     contor = -1
    #     table_project = []
    #
    #     for index1 in range(1, 15):
    #         for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
    #             if str(workSheet.cell(index1,
    #                                    index2).value).casefold().strip() == "Applicabilité projet".casefold() or str(
    #                     workSheet.cell(index1, index2).value).casefold().strip() == "Project applicability".casefold():
    #                 refColIndex = index2
    #                 refRowIndex = index1
    #                 break
    #         if refColIndex != 0:
    #             break
    #     if refColIndex == 0:
    #         var = 1
    #
    #     if var == 1:
    #         result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    #         check = True
    #     elif var == 0:
    #         refCellRange = workSheet.cell(refRowIndex, refColIndex).MergeArea
    #         nrLines = refCellRange.Rows.Count
    #         nrCols = refCellRange.Columns.Count
    #
    #         for index in range(refColIndex, TSDApp.WorkbookStats.tableLastCol + 1):
    #             if workSheet.cell(refRowIndex + nrLines, index).Borders(8).LineStyle != -4142 and workSheet.cell(refRowIndex + nrLines, index).value != None:
    #                 table_project.append(workSheet.cell(refRowIndex + nrLines, index).value.strip())
    #
    #         if TSDApp.WorkbookStats.hasMeasure == False:
    #             result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    #             check = True
    #         else:
    #             workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
    #             codeColIndex = 0
    #             var = 0
    #             for index1 in range(1, 15):
    #                 for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
    #                     if str(workSheet.cell(index1, index2).value).casefold().strip() == "Applicabilité projet".casefold() or str(workSheet.cell(index1, index2).value).casefold().strip() == "Project applicability".casefold():
    #                         measureColIndex = index2
    #                         measureRowIndex = index1
    #                         break
    #                 if codeColIndex != 0:
    #                     break
    #
    #             codeCellRange = workSheet.cell(measureRowIndex, measureColIndex).MergeArea
    #             nrLines = codeCellRange.Rows.Count
    #             nrCols = codeCellRange.Columns.Count
    #
    #             measure_project = []
    #             for index in range(measureColIndex, TSDApp.WorkbookStats.codeLastCol + 1):
    #                 if workSheet.cell(measureRowIndex + nrLines, index).Borders(8).LineStyle != -4142 and workSheet.cell(measureRowIndex + nrLines, index).value != None:
    #                     measure_project.append(workSheet.cell(measureRowIndex + nrLines, index).value.strip())
    #
    #             contor = 0
    #             if len(table_project) == len(measure_project):
    #                 for project in table_project:
    #                     if project in measure_project:
    #                         contor += 1
    #                         break
    #
    #     if contor == len(table_project):
    #         show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    #         check = True
    #     else:
    #         name.append("Different projects!")
    #         show(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], name, workBook, TSDApp)
    # return check

def Test_02043_18_04939_WHOLENESS_1060(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = list()

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow - 1, index).value).casefold().strip() == "Applicabilité projet".casefold() or str(workSheet.cell(TSDApp.tableHeaderRow - 1,index).value).casefold().strip() == "Project applicability".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            contor = 0
            for index in range(refColIndex, TSDApp.WorkbookStats.tableLastCol):
                if workSheet.cell(TSDApp.tableHeaderRow + 1, index).value is not None and workSheet.cell(TSDApp.tableHeaderRow + 1, index).value != "":
                    contor = contor + 1

            values = ["x", "X", "NA", "n/a"]
            if contor != 0:
                for index1 in range(refColIndex, refColIndex + contor):
                    for index2 in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                        if workSheet.cell(index2,0).value is not None or workSheet.cell(index2, 0).value != "":
                            try:
                                if workSheet.cell(index2, index1).value is None:
                                    localisations.append(("tableau",index2, index1))
                                elif workSheet.cell(index2, index1).value.strip() in values:
                                    pass
                            except:
                                localisations.append(("tableau",index2, index1))

        if not localisations:
            localisations = None

        if localisations is not None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
        else:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = list()

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow - 1,
                                  index).value).casefold().strip() == "Applicabilité projet".casefold() or str(
                    workSheet.cell(TSDApp.codeHeaderRow - 1,
                                   index).value).casefold().strip() == "Project applicability".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            contor = 0
            for index in range(refColIndex, TSDApp.WorkbookStats.codeLastCol):
                if workSheet.cell(TSDApp.codeHeaderRow, index).value is not None and workSheet.cell(
                        TSDApp.codeHeaderRow, index).value != "":
                    contor = contor + 1

            values = ["x", "X", "NA", "n/a"]
            if contor != 0:
                for index1 in range(refColIndex, refColIndex + contor):
                    for index2 in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                        if workSheet.cell(index2, 0).value is not None or workSheet.cell(index2, 0).value != "":
                            try:
                                if workSheet.cell(index2, index1).value is None:
                                    localisations.append(("codes défauts", index2, index1))
                                elif workSheet.cell(index2, index1).value.strip() in values:
                                    pass
                            except:
                                localisations.append(("codes défauts", index2, index1))

        if not localisations:
            localisations = None

        if localisations is not None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1062(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = list()

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow - 1,
                                  index).value).casefold().strip() == "Applicabilité projet".casefold() or str(
                workSheet.cell(TSDApp.measureHeaderRow - 1,
                               index).value).casefold().strip() == "Project applicability".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            contor = 0
            for index in range(refColIndex, TSDApp.WorkbookStats.measureLastCol):
                if workSheet.cell(TSDApp.measureHeaderRow, index).value != None and workSheet.cell(
                        TSDApp.measureHeaderRow, index).value != "":
                    contor = contor + 1

            values = ["x", "X", "NA", "n/a"]
            if contor != 0:
                for index1 in range(refColIndex, refColIndex + contor):
                    for index2 in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                        if workSheet.cell(index2, 0).value is not None or workSheet.cell(index2, 0).value != "":
                            try:
                                if workSheet.cell(index2, index1).value is None:
                                    localisations.append(("mesures et commandes", index2, index1))
                                elif workSheet.cell(index2, index1).value.strip() in values:
                                    pass
                            except:
                                localisations.append(("mesures et commandes", index2, index1))

        if not localisations:
            localisations = None

        if localisations is not None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
        else:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        var = 0
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Code défaut".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau",index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Code défaut".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            localisations = []
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, refColIndex).value == None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("codes défauts",index, refColIndex))
                    check = True
            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook, TSDApp)
        elif refColIndex == -1:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "supporté par constituant (s)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            localisations = []
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, refColIndex).value == None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("codes défauts", index, refColIndex))
                    check = True
            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
        elif refColIndex == -1:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "supporté par constituant (s)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            localisations = []
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, refColIndex).value == None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("mesures et commandes", index, refColIndex))
                    check = True
            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
        elif refColIndex == -1:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow,index).value).casefold().strip() == "libellé (signification)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            localisations = []
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, refColIndex).value == None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("codes défauts", index, refColIndex))
                    check = True
            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
        elif refColIndex == -1:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow,index).value).casefold().strip() == "Description de la strategie pour détecter le défaut".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            localisations = []
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, refColIndex).value == None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("codes défauts", index, refColIndex))
                    check = True
            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
        elif refColIndex == -1:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow,index).value).casefold().strip() == "Seuil de détection  /  valeur  du défaut".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            localisations = []
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, refColIndex).value == None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("codes défauts", index, refColIndex))
                    check = True
            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
        elif refColIndex == -1:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow,
                                  index).value).casefold().strip() == "Temps de confirmation du défaut".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            localisations = []
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, refColIndex).value == None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("codes défauts", index, refColIndex))
                    check = True
            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
        elif refColIndex == -1:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow,
                                  index).value).casefold().strip() == "Description de la strategie de disparition du défaut / Procedure à effectuer pour vérifier la disparition du défaut".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            localisations = []
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, refColIndex).value == None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("codes défauts", index, refColIndex))
                    check = True
            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
        elif refColIndex == -1:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow,
                                  index).value).casefold().strip() == "Mode dégradé".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            localisations = []
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, refColIndex).value == None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("codes défauts", index, refColIndex))
                    check = True
            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
        elif refColIndex == -1:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow,
                                  index).value).casefold().strip() == "Voyant".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            localisations = []
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, refColIndex).value == None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("codes défauts", index, refColIndex))
                    check = True
            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
        elif refColIndex == -1:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        var = 0
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Constituant défaillant détecté".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("tableau", index, refColIndex))
                    check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        var = 0
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Défaillance constituant".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("tableau", index, refColIndex))
                    check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        var = 0
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Situation de vie client".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("tableau", index, refColIndex))
                    check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        var = 0
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Effet(s) client(s)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("tableau", index, refColIndex))
                    check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check


def Test_02043_18_04939_WHOLENESS_1220(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        var = 0
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Code défaut".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("tableau", index, refColIndex))
                    check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        var = 0
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Défaillance constituant".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                    localisations.append(("tableau", index, refColIndex))
                    check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
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
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.InfGenIndex)
        # workSheetRange = workSheet.UsedRange
        refColIndex1 = -1
        refRowIndex1 = -1

        var = 0
        for index1 in range(0, workSheet.nrows):
            for index2 in range(0, workSheet.ncols):
                if "Diffusion à :".casefold() in str(workSheet.cell(index1,index2).value).casefold().strip() or "E-mail to :".casefold() in str(workSheet.cell(index1,index2).value).casefold().strip():
                    refColIndex1 = index2
                    refRowIndex1 = index1
                    break
            if refColIndex1 != -1 and refRowIndex1 != -1:
                break
        if refColIndex1 == -1:
            var = 1

        localisations = []

        if var == 0:
            if workSheet.cell(refRowIndex1+1, refColIndex1).value is not None:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                localisations.append(("Informations Générales",refRowIndex1+1, refColIndex1))
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,TSDApp)
        else:
            localisations.append(("Informations Générales",refRowIndex1+1, refColIndex1))
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,TSDApp)
    return check

def Test_02043_18_04939_WHOLENESS_1300(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,index).value).casefold().strip() == "Reference".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1301(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Version".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1302(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "To diagnose".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1303(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Supplier system".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1304(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Logical flow".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1305(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Physical flow".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1306(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Client system".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1307(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Type of connection".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1308(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Type".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1309(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Logical failure mode".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1310(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Physical failure mode".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1311(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Wiring harness cause".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1312(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Other cause".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1313(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Operation situation / Scenario".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1314(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "system effect".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1315(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Customer effect".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1316(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Comment".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1317(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Feared event".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1318(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Reference".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1319(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Severity".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1320(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Level".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1321(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "target".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1322(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Safety measure (G4) / Functional diagnostic(G3,G2,G1)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1323(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Type of failure".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1324(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Degraded mode /Safe state".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1325(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "lead time".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1326(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Efficiency".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1327(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "recovering mode".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1328(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Requirement N° to the Design Document".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1329(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Requirement N° from Design document".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1330(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "research time allocated to the system (in minutes)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1331(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "HMI\n(Indicators/messages)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1332(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "High level test".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1333(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Diagnosis needs".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1334(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,
                                  index).value).casefold().strip() == "Comments".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Table", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1350(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagNeedsLastCol):
            if str(workSheet.cell(TSDApp.diagNeedsHeaderRow,
                                  index).value).casefold().strip() == "Reference".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagNeedsFirstInfoRow, TSDApp.WorkbookStats.DiagNeedsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Diagnostic Needs", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1351(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagNeedsLastCol):
            if str(workSheet.cell(TSDApp.diagNeedsHeaderRow,
                                  index).value).casefold().strip() == "Version".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagNeedsFirstInfoRow, TSDApp.WorkbookStats.DiagNeedsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Diagnostic Needs", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1352(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagNeedsLastCol):
            if str(workSheet.cell(TSDApp.diagNeedsHeaderRow,
                                  index).value).casefold().strip() == "Label".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagNeedsFirstInfoRow, TSDApp.WorkbookStats.DiagNeedsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Diagnostic Needs", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1353(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagNeedsLastCol):
            if str(workSheet.cell(TSDApp.diagNeedsHeaderRow,
                                  index).value).casefold().strip() == "Description".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagNeedsFirstInfoRow, TSDApp.WorkbookStats.DiagNeedsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Diagnostic Needs", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1354(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagNeedsLastCol):
            if str(workSheet.cell(TSDApp.diagNeedsHeaderRow,
                                  index).value).casefold().strip() == "Situation during which the diagnosis is active".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagNeedsFirstInfoRow, TSDApp.WorkbookStats.DiagNeedsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Diagnostic Needs", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1355(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagNeedsLastCol):
            if str(workSheet.cell(TSDApp.diagNeedsHeaderRow,
                                  index).value).casefold().strip() == "Technical Effect covers by the need".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagNeedsFirstInfoRow, TSDApp.WorkbookStats.DiagNeedsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Diagnostic Needs", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1356(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagNeedsLastCol):
            if str(workSheet.cell(TSDApp.diagNeedsHeaderRow,
                                  index).value).casefold().strip() == "Diversity".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagNeedsFirstInfoRow, TSDApp.WorkbookStats.DiagNeedsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Diagnostic Needs", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1357(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagNeedsLastCol):
            if str(workSheet.cell(TSDApp.diagNeedsHeaderRow,
                                  index).value).casefold().strip() == "Allocated to the system".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagNeedsFirstInfoRow, TSDApp.WorkbookStats.DiagNeedsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Diagnostic Needs", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1358(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagNeedsLastCol):
            if str(workSheet.cell(TSDApp.diagNeedsHeaderRow,
                                  index).value).casefold().strip() == "Upstream requirements".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagNeedsFirstInfoRow, TSDApp.WorkbookStats.DiagNeedsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Diagnostic Needs", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1359(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagNeedsLastCol):
            if str(workSheet.cell(TSDApp.diagNeedsHeaderRow,
                                  index).value).casefold().strip() == "Taken into account".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagNeedsFirstInfoRow, TSDApp.WorkbookStats.DiagNeedsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Diagnostic Needs", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1360(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagNeedsLastCol):
            if str(workSheet.cell(TSDApp.diagNeedsHeaderRow,
                                  index).value).casefold().strip() == "comment".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagNeedsFirstInfoRow, TSDApp.WorkbookStats.DiagNeedsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Diagnostic Needs", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1361(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagNeeds == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagNeedsLastCol):
            if str(workSheet.cell(TSDApp.diagNeedsHeaderRow,
                                  index).value).casefold().strip() == "Project applicability".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagNeedsFirstInfoRow, TSDApp.WorkbookStats.DiagNeedsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Diagnostic Needs", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1400(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.EffClientsLastCol):
            if str(workSheet.cell(TSDApp.effClientsHeaderRow,index).value).casefold().strip() == "Name".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.effClientsFirstInfoRow, TSDApp.WorkbookStats.EffClientsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Customer Effects", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1401(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.EffClientsLastCol):
            if str(workSheet.cell(TSDApp.effClientsHeaderRow, index).value).casefold().strip() == "Taken into account".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.effClientsFirstInfoRow, TSDApp.WorkbookStats.EffClientsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Customer Effects", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1402(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.EffClientsLastCol):
            if str(workSheet.cell(TSDApp.effClientsHeaderRow, index).value).casefold().strip() == "Diagnosticability synthesis".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.effClientsFirstInfoRow, TSDApp.WorkbookStats.EffClientsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Customer Effects", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1403(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.EffClientsLastCol):
            if str(workSheet.cell(TSDApp.effClientsHeaderRow, index).value).casefold().strip() == "Comments".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.effClientsFirstInfoRow, TSDApp.WorkbookStats.EffClientsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Customer Effects", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1430(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasReqTech == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ReqTechIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ReqTechLastCol):
            if str(workSheet.cell(TSDApp.reqTechHeaderRow, index).value).casefold().strip() == "Reference".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.reqTechFirstInfoRow, TSDApp.WorkbookStats.ReqTechLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Req. of tech. effects", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1431(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasReqTech == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ReqTechIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ReqTechLastCol):
            if str(workSheet.cell(TSDApp.reqTechHeaderRow, index).value).casefold().strip() == "version".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.reqTechFirstInfoRow, TSDApp.WorkbookStats.ReqTechLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Req. of tech. effects", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1432(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasReqTech == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ReqTechIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ReqTechLastCol):
            if str(workSheet.cell(TSDApp.reqTechHeaderRow, index).value).casefold().strip() == "Description".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.reqTechFirstInfoRow, TSDApp.WorkbookStats.ReqTechLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Req. of tech. effects", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1433(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasReqTech == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ReqTechIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ReqTechLastCol):
            if str(workSheet.cell(TSDApp.reqTechHeaderRow, index).value).casefold().strip() == "technical effect".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.reqTechFirstInfoRow, TSDApp.WorkbookStats.ReqTechLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Req. of tech. effects", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1434(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasReqTech == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ReqTechIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ReqTechLastCol):
            if str(workSheet.cell(TSDApp.reqTechHeaderRow, index).value).casefold().strip() == "Allocated to".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.reqTechFirstInfoRow, TSDApp.WorkbookStats.ReqTechLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Req. of tech. effects", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1435(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasReqTech == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ReqTechIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ReqTechLastCol):
            if str(workSheet.cell(TSDApp.reqTechHeaderRow, index).value).casefold().strip() == "Tracability with the TSD".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.reqTechFirstInfoRow, TSDApp.WorkbookStats.ReqTechLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Req. of tech. effects", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1450(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.FearedEventLastCol):
            if str(workSheet.cell(TSDApp.fearedEventHeaderRow,
                                  index).value).casefold().strip() == "Description".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.fearedEventFirstInfoRow, TSDApp.WorkbookStats.FearedEventLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("feared events", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1451(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.FearedEventLastCol):
            if str(workSheet.cell(TSDApp.fearedEventHeaderRow,
                                  index).value).casefold().strip() == "Reference".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.fearedEventFirstInfoRow, TSDApp.WorkbookStats.FearedEventLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("feared events", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1452(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.FearedEventLastCol):
            if str(workSheet.cell(TSDApp.fearedEventHeaderRow,
                                  index).value).casefold().strip() == "Severity".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.fearedEventFirstInfoRow, TSDApp.WorkbookStats.FearedEventLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("feared events", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1453(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.FearedEventLastCol):
            if str(workSheet.cell(TSDApp.fearedEventHeaderRow,
                                  index).value).casefold().strip() == "Level".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.fearedEventFirstInfoRow, TSDApp.WorkbookStats.FearedEventLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("feared events", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1454(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.FearedEventLastCol):
            if str(workSheet.cell(TSDApp.fearedEventHeaderRow,
                                  index).value).casefold().strip() == "Taken into account".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.fearedEventFirstInfoRow, TSDApp.WorkbookStats.FearedEventLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("feared events", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1455(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.FearedEventLastCol):
            if str(workSheet.cell(TSDApp.fearedEventHeaderRow,
                                  index).value).casefold().strip() == "Justification for not taking into account the dread Event".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.fearedEventFirstInfoRow, TSDApp.WorkbookStats.FearedEventLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("feared events", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1456(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.FearedEventLastCol):
            if str(workSheet.cell(TSDApp.fearedEventHeaderRow,
                                  index).value).casefold().strip() == "Commentaire".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.fearedEventFirstInfoRow, TSDApp.WorkbookStats.FearedEventLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("feared events", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1500(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasSystem == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.SystemIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.SystemLastCol):
            if str(workSheet.cell(TSDApp.systemHeaderRow,
                                  index).value).casefold().strip() == "Description".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.systemFirstInfoRow, TSDApp.WorkbookStats.SystemLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("System", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1501(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasSystem == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.SystemIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.SystemLastCol):
            if str(workSheet.cell(TSDApp.systemHeaderRow,
                                  index).value).casefold().strip() == "Taken into account".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.systemFirstInfoRow, TSDApp.WorkbookStats.SystemLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("System", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1550(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasOpSit == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.OpSitIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.OpSitLastCol):
            if str(workSheet.cell(TSDApp.opSitHeaderRow,
                                  index).value).casefold().strip() == "Description".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.opSitFirstInfoRow, TSDApp.WorkbookStats.OpSitLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Operation situation", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1551(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasOpSit == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.OpSitIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.OpSitLastCol):
            if str(workSheet.cell(TSDApp.opSitHeaderRow,
                                  index).value).casefold().strip() == "Taken into account".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.opSitFirstInfoRow, TSDApp.WorkbookStats.OpSitLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Operation situation", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1552(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasOpSit == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.OpSitIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.OpSitLastCol):
            if str(workSheet.cell(TSDApp.opSitHeaderRow,
                                  index).value).casefold().strip() == "Comments".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.opSitFirstInfoRow, TSDApp.WorkbookStats.OpSitLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Operation situation", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1600(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,index).value).casefold().strip() == "Référence".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check


def Test_02043_18_04939_WHOLENESS_1601(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Version".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1602(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Réf doc".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1603(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Variante/\noption".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1604(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Version de soft (MOTEUR / BSI,...)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1605(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "sous Fonction de conception incriminée".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1606(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Groupe de constituant".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1607(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Constituant défaillant détecté".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1608(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Flux fonctionnel".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1609(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Défaillance logique".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1610(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Défaillance constituant".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1611(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow + 1, index).value).casefold().strip() == "PPM réparties".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1612(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow + 1, index).value).casefold().strip() == "poids".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1613(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Situation de vie client".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1614(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Situation de vie détaillée".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1615(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "lien vers autre TSD".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1616(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Effet(s) client(s)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1617(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Evenement(s) redouté(s) (ER)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1618(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Voyant(s) ou \nmessage(s) ".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1619(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Code défaut".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1620(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Code défauts induits".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1621(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "mesures et commandes (Mesure Parametre et Test Actionneur) / Tests de cohérence".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1622(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Critère de décision".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1623(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "DIAGNOSTIC DEBARQUE".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1624(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Critère de decision".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1625(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Action sur constituant incriminé".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1626(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Statut réunion DSP-DRD".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1627(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Action a réaliser / Commentaires".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1628(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Référence AMDEC".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1629(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Justification de la modification".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1630(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Validation".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1631(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow - 1, index).value).casefold().strip() == "Controle Usine".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1632(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow - 1, index).value).casefold().strip() == "Pris en compte dans logigramme".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1650(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Référence".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1651(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Version".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1652(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Code défaut".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1653(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "libellé (signification)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1654(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Flux Fonctionnel".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1655(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Description de la strategie pour détecter le défaut".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1656(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Seuil de détection  /  valeur  du défaut".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1657(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Temps de confirmation du défaut".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1658(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Description de la strategie de disparition du défaut / Procedure à effectuer pour vérifier la disparition du défaut".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1659(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Situation de vie véhicule pour faire remonter le code défaut".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1660(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Mode dégradé".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1661(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Taux de remonté du code défaut".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1662(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Voyant".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1663(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Accès scantool".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1664(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Groupe de contextes associés".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1684(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Diversité".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1685(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Applicabilité usine".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1686(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "condition d'applicabilité en usine".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1687(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "supporté par constituant (s)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1688(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "se référer au document spécifiant DRD : (réf & version)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1689(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Référence amont".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1690(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Version de la référence amont".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1691(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Pris en compte".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1692(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Justification de la modification".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1693(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Validation".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("codes défauts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1700(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "Référence".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("mesures et commandes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1701(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "Version".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("mesures et commandes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1702(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "Type (choix par menu)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("mesures et commandes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1703(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "libellé (signification)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("mesures et commandes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1704(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "Description".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("mesures et commandes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1705(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "Situation pendant laquelle la mesure ou commande est utilisable".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("mesures et commandes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1706(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "Statut".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("mesures et commandes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1707(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "Taux de fiabilité du test (50%, 100%)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("mesures et commandes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1708(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "Flux fonctionnel".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("mesures et commandes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1709(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "Uniquement \npour O Control\nlecture \nsortie effective /commande".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("mesures et commandes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check


def Test_02043_18_04939_WHOLENESS_1710(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "Diversité".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("mesures et commandes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check


def Test_02043_18_04939_WHOLENESS_1711(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "Applicabilité usine".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("mesures et commandes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1712(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "condition d'applicabilité en usine".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("mesures et commandes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1713(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "supporté par constituant (s)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("mesures et commandes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1714(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "se référer au document spécifiant DRD : (réf & version)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("mesures et commandes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1715(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "Référence amont".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("mesures et commandes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1716(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "Version de la référence amont".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("mesures et commandes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1717(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "Pris en compte".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("mesures et commandes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1718(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "Justification de la modification".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("mesures et commandes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1719(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "Validation".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("mesures et commandes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1750(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagDebIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagDebLastCol):
            if str(workSheet.cell(TSDApp.diagDebHeaderRow, index).value).casefold().strip() == "Référence".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagDebFirstInfoRow, TSDApp.WorkbookStats.DiagDebLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Diagnostic débarqués", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1751(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagDebIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagDebLastCol):
            if str(workSheet.cell(TSDApp.diagDebHeaderRow, index).value).casefold().strip() == "Version".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagDebFirstInfoRow, TSDApp.WorkbookStats.DiagDebLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Diagnostic débarqués", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1752(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagDebIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagDebLastCol):
            if str(workSheet.cell(TSDApp.diagDebHeaderRow, index).value).casefold().strip() == "libellé (signification)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagDebFirstInfoRow, TSDApp.WorkbookStats.DiagDebLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Diagnostic débarqués", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1753(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagDebIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagDebLastCol):
            if str(workSheet.cell(TSDApp.diagDebHeaderRow, index).value).casefold().strip() == "Description".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagDebFirstInfoRow, TSDApp.WorkbookStats.DiagDebLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Diagnostic débarqués", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1754(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagDebIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagDebLastCol):
            if str(workSheet.cell(TSDApp.diagDebHeaderRow, index).value).casefold().strip() == "Taux de fiabilité du test (50%, 100%)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagDebFirstInfoRow, TSDApp.WorkbookStats.DiagDebLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Diagnostic débarqués", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1755(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagDebIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagDebLastCol):
            if str(workSheet.cell(TSDApp.diagDebHeaderRow, index).value).casefold().strip() == "Applicabilité Usine".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagDebFirstInfoRow, TSDApp.WorkbookStats.DiagDebLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Diagnostic débarqués", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1756(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagDebIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagDebLastCol):
            if str(workSheet.cell(TSDApp.diagDebHeaderRow, index).value).casefold().strip() == "se référer au document spécifiant : (réf & version)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagDebFirstInfoRow, TSDApp.WorkbookStats.DiagDebLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Diagnostic débarqués", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1757(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagDebIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagDebLastCol):
            if str(workSheet.cell(TSDApp.diagDebHeaderRow, index).value).casefold().strip() == "Pris en compte".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagDebFirstInfoRow, TSDApp.WorkbookStats.DiagDebLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Diagnostic débarqués", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1758(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagDebIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagDebLastCol):
            if str(workSheet.cell(TSDApp.diagDebHeaderRow, index).value).casefold().strip() == "Justification de la modification".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagDebFirstInfoRow, TSDApp.WorkbookStats.DiagDebLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Diagnostic débarqués", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1759(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDiagDeb == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagDebIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DiagDebLastCol):
            if str(workSheet.cell(TSDApp.diagDebHeaderRow, index).value).casefold().strip() == "Validation".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.diagDebFirstInfoRow, TSDApp.WorkbookStats.DiagDebLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Diagnostic débarqués", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1800(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.EffClientsLastCol):
            if str(workSheet.cell(TSDApp.effClientsHeaderRow, index).value).casefold().strip() == "Noms".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.effClientsFirstInfoRow, TSDApp.WorkbookStats.EffClientsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Effets clients", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1801(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.EffClientsLastCol):
            if str(workSheet.cell(TSDApp.effClientsHeaderRow,
                                  index).value).casefold().strip() == "Pris en compte".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.effClientsFirstInfoRow, TSDApp.WorkbookStats.EffClientsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Effets clients", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1802(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.EffClientsLastCol):
            if str(workSheet.cell(TSDApp.effClientsHeaderRow,
                                  index).value).casefold().strip() == "Synthèse de la diagnosticabilité".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.effClientsFirstInfoRow, TSDApp.WorkbookStats.EffClientsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Effets clients", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1803(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.EffClientsLastCol):
            if str(workSheet.cell(TSDApp.effClientsHeaderRow,
                                  index).value).casefold().strip() == "Justification de la modification".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.effClientsFirstInfoRow, TSDApp.WorkbookStats.EffClientsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Effets clients", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1810(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasER == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ERIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ERLastCol):
            if str(workSheet.cell(TSDApp.ERHeaderRow,
                                  index).value).casefold().strip() == "nom".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.ERFirstInfoRow, TSDApp.WorkbookStats.ERLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("ER", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1811(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasER == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ERIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ERLastCol):
            if str(workSheet.cell(TSDApp.ERHeaderRow,
                                  index).value).casefold().strip() == "désignation".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.ERFirstInfoRow, TSDApp.WorkbookStats.ERLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("ER", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1812(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasER == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ERIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ERLastCol):
            if str(workSheet.cell(TSDApp.ERHeaderRow,
                                  index).value).casefold().strip() == "Gravité".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.ERFirstInfoRow, TSDApp.WorkbookStats.ERLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("ER", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1813(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasER == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ERIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ERLastCol):
            if str(workSheet.cell(TSDApp.ERHeaderRow,
                                  index).value).casefold().strip() == "Pris en compte".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.ERFirstInfoRow, TSDApp.WorkbookStats.ERLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("ER", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1814(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasER == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ERIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ERLastCol):
            if str(workSheet.cell(TSDApp.ERHeaderRow,
                                  index).value).casefold().strip() == "Justification de non prise en compte de l'ER".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.ERFirstInfoRow, TSDApp.WorkbookStats.ERLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("ER", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1815(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasER == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ERIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ERLastCol):
            if str(workSheet.cell(TSDApp.ERHeaderRow,
                                  index).value).casefold().strip() == "Justification de la modification".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.ERFirstInfoRow, TSDApp.WorkbookStats.ERLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("ER", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1820(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasConstituants == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.constituantsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.constituantsLastCol):
            if str(workSheet.cell(TSDApp.constituantsHeaderRow,index).value).casefold().strip() == "Noms".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.constituantsFirstInfoRow, TSDApp.WorkbookStats.constituantsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Constituants", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1821(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasConstituants == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.constituantsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.constituantsLastCol):
            if str(workSheet.cell(TSDApp.constituantsHeaderRow, index).value).casefold().strip() == "Description".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.constituantsFirstInfoRow, TSDApp.WorkbookStats.constituantsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Constituants", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1822(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasConstituants == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.constituantsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.constituantsLastCol):
            if str(workSheet.cell(TSDApp.constituantsHeaderRow, index).value).casefold().strip() == "Taux de défaillance (en ppm)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.constituantsFirstInfoRow, TSDApp.WorkbookStats.constituantsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Constituants", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1823(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasConstituants == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.constituantsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.constituantsLastCol):
            if str(workSheet.cell(TSDApp.constituantsHeaderRow, index).value).casefold().strip() == "Découpage PSA".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.constituantsFirstInfoRow, TSDApp.WorkbookStats.constituantsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Constituants", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1824(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasConstituants == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.constituantsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.constituantsLastCol):
            if str(workSheet.cell(TSDApp.constituantsHeaderRow, index).value).casefold().strip() == "Pris en compte".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.constituantsFirstInfoRow, TSDApp.WorkbookStats.constituantsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Constituants", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1825(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasConstituants == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.constituantsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.constituantsLastCol):
            if str(workSheet.cell(TSDApp.constituantsHeaderRow, index).value).casefold().strip() == "Justification de la modification".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.constituantsFirstInfoRow, TSDApp.WorkbookStats.constituantsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Constituants", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1830(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasSitDeVie == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.SitDeVieIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.SitDeVieLastCol):
            if str(workSheet.cell(TSDApp.sitDeVieHeaderRow,index).value).casefold().strip() == "Situations de vie".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.sitDeVieFirstInfoRow, TSDApp.WorkbookStats.SitDeVieLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("situations de vie", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1831(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasSitDeVie == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.SitDeVieIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.SitDeVieLastCol):
            if str(workSheet.cell(TSDApp.sitDeVieHeaderRow,
                                  index).value).casefold().strip() == "Justification de la modification".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.sitDeVieFirstInfoRow, TSDApp.WorkbookStats.SitDeVieLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("situations de vie", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1840(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMDD == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.MDDIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.MDDLastCol):
            if str(workSheet.cell(TSDApp.listeMDDHeaderRow,
                                  index).value).casefold().strip() == "Modes dégradés:".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.listeMDDFirstInfoRow, TSDApp.WorkbookStats.MDDLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Liste MDD", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1841(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMDD == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.MDDIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.MDDLastCol):
            if str(workSheet.cell(TSDApp.listeMDDHeaderRow,
                                  index).value).casefold().strip() == "Justification de la modification".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.listeMDDFirstInfoRow, TSDApp.WorkbookStats.MDDLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Liste MDD", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1900(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Reference".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1901(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Version".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1902(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Document of reference".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1903(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Variant/\noption".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1904(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Sub-function of the system incriminated".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1905(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Module / Group of parts".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1906(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Defective part".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1907(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Contribution to fonctionnality".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1908(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Logical failure mode".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1909(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Physical failure mode".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1910(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Weight".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1911(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Situation".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1912(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Detailed situation".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check


def Test_02043_18_04939_WHOLENESS_1913(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Link to another DST".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1914(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Technical effect".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1915(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Customer effect".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1916(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Feared events".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1917(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Degraded mode".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1918(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "HMI\n(Indicator lights/messages)".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1919(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Data Trouble code".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1920(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Mislead Data trouble code".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1921(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Read data or I/O control".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1922(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "decision criterion".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1923(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Non-embedded diagnosis".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1924(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "decision criterion".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1925(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Action on the incriminated part".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1926(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "to do list / Comments".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1927(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "FMEA reference".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("tableau", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1950(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DataCodesIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DataCodesLastCol):
            if str(workSheet.cell(TSDApp.dataCodesHeaderRow,index).value).casefold().strip() == "Reference".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.dataCodesFirstInfoRow, TSDApp.WorkbookStats.DataCodesLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Data trouble codes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check


def Test_02043_18_04939_WHOLENESS_1951(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DataCodesIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DataCodesLastCol):
            if str(workSheet.cell(TSDApp.dataCodesHeaderRow, index).value).casefold().strip() == "Version".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.dataCodesFirstInfoRow, TSDApp.WorkbookStats.DataCodesLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Data trouble codes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1952(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DataCodesIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DataCodesLastCol):
            if str(workSheet.cell(TSDApp.dataCodesHeaderRow, index).value).casefold().strip() == "Data trouble code".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.dataCodesFirstInfoRow, TSDApp.WorkbookStats.DataCodesLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Data trouble codes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1953(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DataCodesIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DataCodesLastCol):
            if str(workSheet.cell(TSDApp.dataCodesHeaderRow, index).value).casefold().strip() == "Label".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.dataCodesFirstInfoRow, TSDApp.WorkbookStats.DataCodesLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Data trouble codes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1954(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DataCodesIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DataCodesLastCol):
            if str(workSheet.cell(TSDApp.dataCodesHeaderRow, index).value).casefold().strip() == "Description of the qualification conditions".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.dataCodesFirstInfoRow, TSDApp.WorkbookStats.DataCodesLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Data trouble codes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1955(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DataCodesIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DataCodesLastCol):
            if str(workSheet.cell(TSDApp.dataCodesHeaderRow, index).value).casefold().strip() == "Detection threshold".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.dataCodesFirstInfoRow, TSDApp.WorkbookStats.DataCodesLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Data trouble codes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1956(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DataCodesIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DataCodesLastCol):
            if str(workSheet.cell(TSDApp.dataCodesHeaderRow, index).value).casefold().strip() == "Qualification time".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.dataCodesFirstInfoRow, TSDApp.WorkbookStats.DataCodesLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Data trouble codes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1957(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DataCodesIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DataCodesLastCol):
            if str(workSheet.cell(TSDApp.dataCodesHeaderRow, index).value).casefold().strip() == "Description of the dequalification conditions / Operation to do to check if the defect disappeared".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.dataCodesFirstInfoRow, TSDApp.WorkbookStats.DataCodesLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Data trouble codes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1958(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DataCodesIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DataCodesLastCol):
            if str(workSheet.cell(TSDApp.dataCodesHeaderRow, index).value).casefold().strip() == "Conditions of the diagnostic activation".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.dataCodesFirstInfoRow, TSDApp.WorkbookStats.DataCodesLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Data trouble codes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1959(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DataCodesIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DataCodesLastCol):
            if str(workSheet.cell(TSDApp.dataCodesHeaderRow, index).value).casefold().strip() == "Degraded mode".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.dataCodesFirstInfoRow, TSDApp.WorkbookStats.DataCodesLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Data trouble codes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1960(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DataCodesIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DataCodesLastCol):
            if str(workSheet.cell(TSDApp.dataCodesHeaderRow, index).value).casefold().strip() == "Failure detection rate".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.dataCodesFirstInfoRow, TSDApp.WorkbookStats.DataCodesLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Data trouble codes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1961(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DataCodesIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DataCodesLastCol):
            if str(workSheet.cell(TSDApp.dataCodesHeaderRow, index).value).casefold().strip() == "Indicateur light".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.dataCodesFirstInfoRow, TSDApp.WorkbookStats.DataCodesLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Data trouble codes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1962(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DataCodesIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DataCodesLastCol):
            if str(workSheet.cell(TSDApp.dataCodesHeaderRow, index).value).casefold().strip() == "Visibility of the failure with the Scantool".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.dataCodesFirstInfoRow, TSDApp.WorkbookStats.DataCodesLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Data trouble codes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1963(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DataCodesIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DataCodesLastCol):
            if str(workSheet.cell(TSDApp.dataCodesHeaderRow, index).value).casefold().strip() == "Freeze Frame Class".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.dataCodesFirstInfoRow, TSDApp.WorkbookStats.DataCodesLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Data trouble codes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1964(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DataCodesIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DataCodesLastCol):
            if str(workSheet.cell(TSDApp.dataCodesHeaderRow, index).value).casefold().strip() == "Diversity".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.dataCodesFirstInfoRow, TSDApp.WorkbookStats.DataCodesLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Data trouble codes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1965(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DataCodesIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DataCodesLastCol):
            if str(workSheet.cell(TSDApp.dataCodesHeaderRow, index).value).casefold().strip() == "Stored by the ECU".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.dataCodesFirstInfoRow, TSDApp.WorkbookStats.DataCodesLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Data trouble codes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1966(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DataCodesIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DataCodesLastCol):
            if str(workSheet.cell(TSDApp.dataCodesHeaderRow, index).value).casefold().strip() == "Upstream requirements".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.dataCodesFirstInfoRow, TSDApp.WorkbookStats.DataCodesLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Data trouble codes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1967(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DataCodesIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DataCodesLastCol):
            if str(workSheet.cell(TSDApp.dataCodesHeaderRow, index).value).casefold().strip() == "Taken into account".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.dataCodesFirstInfoRow, TSDApp.WorkbookStats.DataCodesLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Data trouble codes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1968(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DataCodesIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DataCodesLastCol):
            if str(workSheet.cell(TSDApp.dataCodesHeaderRow, index).value).casefold().strip() == "projet X".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.dataCodesFirstInfoRow, TSDApp.WorkbookStats.DataCodesLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Data trouble codes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_1969(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DataCodesIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DataCodesLastCol):
            if str(workSheet.cell(TSDApp.dataCodesHeaderRow, index).value).casefold().strip() == "Projet Y".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.dataCodesFirstInfoRow, TSDApp.WorkbookStats.DataCodesLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Data trouble codes", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ReadDataIOLastCol):
            if str(workSheet.cell(TSDApp.readDataIOHeaderRow, index).value).casefold().strip() == "Reference".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.readDataIOFirstInfoRow, TSDApp.WorkbookStats.ReadDataIOLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Read data and IO control", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2001(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ReadDataIOLastCol):
            if str(workSheet.cell(TSDApp.readDataIOHeaderRow, index).value).casefold().strip() == "Version".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.readDataIOFirstInfoRow, TSDApp.WorkbookStats.ReadDataIOLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Read data and IO control", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2002(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ReadDataIOLastCol):
            if str(workSheet.cell(TSDApp.readDataIOHeaderRow, index).value).casefold().strip() == "Type of diagnosis".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.readDataIOFirstInfoRow, TSDApp.WorkbookStats.ReadDataIOLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Read data and IO control", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2003(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ReadDataIOLastCol):
            if str(workSheet.cell(TSDApp.readDataIOHeaderRow, index).value).casefold().strip() == "Label".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.readDataIOFirstInfoRow, TSDApp.WorkbookStats.ReadDataIOLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Read data and IO control", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2004(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ReadDataIOLastCol):
            if str(workSheet.cell(TSDApp.readDataIOHeaderRow, index).value).casefold().strip() == "Description".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.readDataIOFirstInfoRow, TSDApp.WorkbookStats.ReadDataIOLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Read data and IO control", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2005(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ReadDataIOLastCol):
            if str(workSheet.cell(TSDApp.readDataIOHeaderRow, index).value).casefold().strip() == "Conditions of the diagnostic activation".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.readDataIOFirstInfoRow, TSDApp.WorkbookStats.ReadDataIOLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Read data and IO control", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2006(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ReadDataIOLastCol):
            if str(workSheet.cell(TSDApp.readDataIOHeaderRow, index).value).casefold().strip() == "Status".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.readDataIOFirstInfoRow, TSDApp.WorkbookStats.ReadDataIOLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Read data and IO control", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2007(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ReadDataIOLastCol):
            if str(workSheet.cell(TSDApp.readDataIOHeaderRow, index).value).casefold().strip() == "Diversity".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.readDataIOFirstInfoRow, TSDApp.WorkbookStats.ReadDataIOLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Read data and IO control", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2008(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ReadDataIOLastCol):
            if str(workSheet.cell(TSDApp.readDataIOHeaderRow, index).value).casefold().strip() == "Stored by the ECU".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.readDataIOFirstInfoRow, TSDApp.WorkbookStats.ReadDataIOLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Read data and IO control", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2009(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ReadDataIOLastCol):
            if str(workSheet.cell(TSDApp.readDataIOHeaderRow, index).value).casefold().strip() == "Upstream requirements".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.readDataIOFirstInfoRow, TSDApp.WorkbookStats.ReadDataIOLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Read data and IO control", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2010(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ReadDataIOLastCol):
            if str(workSheet.cell(TSDApp.readDataIOHeaderRow, index).value).casefold().strip() == "Taken into account".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.readDataIOFirstInfoRow, TSDApp.WorkbookStats.ReadDataIOLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Read data and IO control", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check


def Test_02043_18_04939_WHOLENESS_2011(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.ReadDataIOLastCol):
            if str(workSheet.cell(TSDApp.readDataIOHeaderRow, index).value).casefold().strip() == "projet X".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.readDataIOFirstInfoRow, TSDApp.WorkbookStats.ReadDataIOLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Read data and IO control", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2050(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasNotEmbDiag == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.NotEmbDiagIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.NotEmbDiagLastCol):
            if str(workSheet.cell(TSDApp.notEmbDiagHeaderRow, index).value).casefold().strip() == "Reference".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.notEmbDiagFirstInfoRow, TSDApp.WorkbookStats.NotEmbDiagLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Not embedded diagnosis", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2051(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasNotEmbDiag == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.NotEmbDiagIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.NotEmbDiagLastCol):
            if str(workSheet.cell(TSDApp.notEmbDiagHeaderRow,
                                  index).value).casefold().strip() == "Version".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.notEmbDiagFirstInfoRow, TSDApp.WorkbookStats.NotEmbDiagLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Not embedded diagnosis", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2052(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasNotEmbDiag == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.NotEmbDiagIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.NotEmbDiagLastCol):
            if str(workSheet.cell(TSDApp.notEmbDiagHeaderRow,
                                  index).value).casefold().strip() == "Label".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.notEmbDiagFirstInfoRow, TSDApp.WorkbookStats.NotEmbDiagLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Not embedded diagnosis", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2053(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasNotEmbDiag == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.NotEmbDiagIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.NotEmbDiagLastCol):
            if str(workSheet.cell(TSDApp.notEmbDiagHeaderRow,
                                  index).value).casefold().strip() == "Description".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.notEmbDiagFirstInfoRow, TSDApp.WorkbookStats.NotEmbDiagLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Not embedded diagnosis", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2054(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasNotEmbDiag == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.NotEmbDiagIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.NotEmbDiagLastCol):
            if str(workSheet.cell(TSDApp.notEmbDiagHeaderRow,
                                  index).value).casefold().strip() == "Upstream requirements".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.notEmbDiagFirstInfoRow, TSDApp.WorkbookStats.NotEmbDiagLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Not embedded diagnosis", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2055(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasNotEmbDiag == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.NotEmbDiagIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.NotEmbDiagLastCol):
            if str(workSheet.cell(TSDApp.notEmbDiagHeaderRow,
                                  index).value).casefold().strip() == "Taken into account".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.notEmbDiagFirstInfoRow, TSDApp.WorkbookStats.NotEmbDiagLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Not embedded diagnosis", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2056(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasNotEmbDiag == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.NotEmbDiagIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.NotEmbDiagLastCol):
            if str(workSheet.cell(TSDApp.notEmbDiagHeaderRow,
                                  index).value).casefold().strip() == "projet X".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.notEmbDiagFirstInfoRow, TSDApp.WorkbookStats.NotEmbDiagLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Not embedded diagnosis", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2060(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTechEff == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.TechEffIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.TechEffLastCol):
            if str(workSheet.cell(TSDApp.techEffHeaderRow,
                                  index).value).casefold().strip() == "Name".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.techEffFirstInfoRow, TSDApp.WorkbookStats.TechEffLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Technical effect", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2061(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTechEff == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.TechEffIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.TechEffLastCol):
            if str(workSheet.cell(TSDApp.techEffHeaderRow,
                                  index).value).casefold().strip() == "Taken into account".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.techEffFirstInfoRow, TSDApp.WorkbookStats.TechEffLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Technical effect", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2062(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTechEff == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.TechEffIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.TechEffLastCol):
            if str(workSheet.cell(TSDApp.techEffHeaderRow,
                                  index).value).casefold().strip() == "Upstream requirements".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.techEffFirstInfoRow, TSDApp.WorkbookStats.TechEffLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Technical effect", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2070(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.EffClientsLastCol):
            if str(workSheet.cell(TSDApp.effClientsHeaderRow,
                                  index).value).casefold().strip() == "Name".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.effClientsFirstInfoRow, TSDApp.WorkbookStats.EffClientsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Customer effect", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2071(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.EffClientsLastCol):
            if str(workSheet.cell(TSDApp.effClientsHeaderRow,
                                  index).value).casefold().strip() == "Taken into account".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.effClientsFirstInfoRow, TSDApp.WorkbookStats.EffClientsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Customer effect", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check


def Test_02043_18_04939_WHOLENESS_2072(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.EffClientsLastCol):
            if str(workSheet.cell(TSDApp.effClientsHeaderRow,
                                  index).value).casefold().strip() == "Diagnosticability synthesis".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.effClientsFirstInfoRow, TSDApp.WorkbookStats.EffClientsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Customer effect", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check


def Test_02043_18_04939_WHOLENESS_2080(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.FearedEventLastCol):
            if str(workSheet.cell(TSDApp.fearedEventHeaderRow,
                                  index).value).casefold().strip() == "Description".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.fearedEventFirstInfoRow, TSDApp.WorkbookStats.FearedEventLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Feared events", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check


def Test_02043_18_04939_WHOLENESS_2081(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.FearedEventLastCol):
            if str(workSheet.cell(TSDApp.fearedEventHeaderRow,
                                  index).value).casefold().strip() == "Reference".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.fearedEventFirstInfoRow, TSDApp.WorkbookStats.FearedEventLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Feared events", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2082(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.FearedEventLastCol):
            if str(workSheet.cell(TSDApp.fearedEventHeaderRow,
                                  index).value).casefold().strip() == "Severity".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.fearedEventFirstInfoRow, TSDApp.WorkbookStats.FearedEventLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Feared events", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2083(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.FearedEventLastCol):
            if str(workSheet.cell(TSDApp.fearedEventHeaderRow,
                                  index).value).casefold().strip() == "Taken into account".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.fearedEventFirstInfoRow, TSDApp.WorkbookStats.FearedEventLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Feared events", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2084(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasFearedEvent == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.FearedEventIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.FearedEventLastCol):
            if str(workSheet.cell(TSDApp.fearedEventHeaderRow,index).value).casefold().strip() == "Justification for not taking into account the dread Event".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.fearedEventFirstInfoRow, TSDApp.WorkbookStats.FearedEventLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Feared events", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2090(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasParts == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.PartsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.PartsLastCol):
            if str(workSheet.cell(TSDApp.partsHeaderRow,
                                  index).value).casefold().strip() == "Name".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.partsFirstInfoRow, TSDApp.WorkbookStats.PartsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Parts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2091(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasParts == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.PartsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.PartsLastCol):
            if str(workSheet.cell(TSDApp.partsHeaderRow,
                                  index).value).casefold().strip() == "Description".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.partsFirstInfoRow, TSDApp.WorkbookStats.PartsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Parts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2092(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasParts == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.PartsIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.PartsLastCol):
            if str(workSheet.cell(TSDApp.partsHeaderRow,
                                  index).value).casefold().strip() == "Taken into account".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.partsFirstInfoRow, TSDApp.WorkbookStats.PartsLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Parts", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2100(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasVariant == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.VariantIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.VariantLastCol):
            if str(workSheet.cell(TSDApp.variantHeaderRow,
                                  index).value).casefold().strip() == "Name".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.variantFirstInfoRow, TSDApp.WorkbookStats.VariantLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Variant", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2101(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasVariant == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.VariantIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.VariantLastCol):
            if str(workSheet.cell(TSDApp.variantHeaderRow,
                                  index).value).casefold().strip() == "Description".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.variantFirstInfoRow, TSDApp.WorkbookStats.VariantLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Variant", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2102(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasVariant == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.VariantIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.VariantLastCol):
            if str(workSheet.cell(TSDApp.variantHeaderRow,
                                  index).value).casefold().strip() == "Taken into account".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.variantFirstInfoRow, TSDApp.WorkbookStats.VariantLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Variant", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2110(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasSituation == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.SituationIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.SituationLastCol):
            if str(workSheet.cell(TSDApp.situationHeaderRow,
                                  index).value).casefold().strip() == "Description".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.situationFirstInfoRow, TSDApp.WorkbookStats.SituationLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Situation", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2111(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasSituation == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.SituationIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.SituationLastCol):
            if str(workSheet.cell(TSDApp.situationHeaderRow,
                                  index).value).casefold().strip() == "Taken into account".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.situationFirstInfoRow, TSDApp.WorkbookStats.SituationLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Situation", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2112(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasSituation == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.SituationIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.SituationLastCol):
            if str(workSheet.cell(TSDApp.situationHeaderRow,
                                  index).value).casefold().strip() == "Comments".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.situationFirstInfoRow, TSDApp.WorkbookStats.SituationLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Situation", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2120(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDegradedMode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DegradedModeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DegradedModeLastCol):
            if str(workSheet.cell(TSDApp.degradedModeHeaderRow,
                                  index).value).casefold().strip() == "Modes dégradés:".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.degradedModeFirstInfoRow, TSDApp.WorkbookStats.DegradedModeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Degraded mode", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_WHOLENESS_2121(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDegradedMode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DegradedModeIndex)
        refColIndex = -1
        localisations = []

        for index in range(0, TSDApp.WorkbookStats.DegradedModeLastCol):
            if str(workSheet.cell(TSDApp.degradedModeHeaderRow,
                                  index).value).casefold().strip() == "Taken into account".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            for index in range(TSDApp.degradedModeFirstInfoRow, TSDApp.WorkbookStats.DegradedModeLastRow):
                if workSheet.cell(index, 0).value is not None or workSheet.cell(index, 0).value != "":
                    if workSheet.cell(index, refColIndex).value is None or workSheet.cell(index, refColIndex).value == "":
                        localisations.append(("Degraded mode", index, refColIndex))
                        check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)

        elif refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check
