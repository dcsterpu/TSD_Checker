import TSD_Checker_V5_0
import inspect
from ExcelEdit import TestReturn as result
from ErrorMessages import errorMessagesDict as error
import xlrd

#Coherence checks requirements

def Test_02043_18_04939_COH_2000(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "mesures et commandes (Mesure Parametre et Test Actionneur) / Tests de cohérence".casefold():
                refColIndex = index
                break

        if refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif refColIndex != -1:
            localisations = list()
            list_table = list()
            list_measure = list()

            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, refColIndex).value == "N/A" or workSheet.cell(index, refColIndex).value == "":
                    pass
                else:
                    dict = {}
                    dict["value"] = workSheet.cell(index, refColIndex).value
                    dict["row"] = index
                    dict["col"] = refColIndex
                    list_table.append(dict)

            if TSDApp.WorkbookStats.hasMeasure == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
                measureColIndex = -1

                for index in range(0, TSDApp.WorkbookStats.measureLastCol):
                    if str(workSheet.cell(TSDApp.measureHeaderRow,index).value).casefold().strip() == "libellé (signification)".casefold():
                        measureColIndex = index
                        break

                for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                    if workSheet.cell(index, measureColIndex).value == "":
                        pass
                    else:
                        list_measure.append(workSheet.cell(index, measureColIndex).value)

            for element in list_table:
                if element["value"] in list_measure:
                    pass
                else:
                    localisations.append(("tableau", element["row"], element["col"]))

            if not localisations:
                localisations = None

            if localisations is None:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisations, workBook,TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook, TSDApp)

    return check

def is_ascii(s):
    return all(ord(c) < 128 for c in s)

def Test_02043_18_04939_COH_2001(workBook, TSDApp, DOC8List):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow,index).value).casefold().strip() == "mesures et commandes (Mesure Parametre et Test Actionneur) / Tests de cohérence".casefold():
                refColIndex = index
                break

        if refColIndex != -1:
            localisations = []
            contor = 0

            for index in range(TSDApp.tableHeaderRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, refColIndex).value == "":
                    pass
                else:
                    try:
                        cel = workSheet.cell(index, refColIndex).value.split("-")
                        if len(cel) == 2:

                            check1 = False
                            if len(cel[1]) == 4:
                                check1 = True

                            check2 = True
                            mystring = cel[0]
                            for char in mystring:
                                if not (is_ascii(char)):
                                    check2 = False
                                    break
                            if check1 == True and check2 == True:
                                if cel[0] in DOC8List:
                                    contor = contor + 1
                        else:
                            localisations.append(("tableau",index, refColIndex))

                    except:
                        localisations.append(("tableau",index, refColIndex))
            if not localisations:
                localisations = None
            if contor == TSDApp.WorkbookStats.tableLastRow - TSDApp.tableHeaderRow - 1:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisations, workBook,TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,TSDApp)
    return check

def Test_02043_18_04939_COH_2002(workBook, TSDApp, DOC8List):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)

        codeColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Code défaut".casefold():
                codeColIndex = index
                break

        if codeColIndex != -1:
            localisations = []

            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                try:
                    cel = workSheet.cell(index, codeColIndex).value.split("-")
                    if len(cel) == 2:
                        pass
                    elif cel[0] not in DOC8List:
                        localisations.append(("codes défauts",index, codeColIndex))
                except:
                    localisations.append(("codes défauts",index, codeColIndex))
            if not localisations:
                localisations = None
            if not localisations:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisations, workBook,TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,TSDApp)
    return check

def Test_02043_18_04939_COH_2005(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        codeColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Code défaut".casefold():
                codeColIndex = index
                break

        if codeColIndex != -1:
            localisations = []

            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                try:
                    cel = workSheet.cell(index, codeColIndex).value.split("-")
                    if len(cel) == 2:
                        if cel[0].isascii() and cel[1][0].isalpha() and len(cel[1]) == 5:
                            try:
                                int(cel[1][1:], 16)
                            except:
                                localisations.append(("codes défauts",index,codeColIndex))
                    else:
                        if len(cel) == 3:
                            a = 3
                except:
                    localisations.append(("codes défauts",index,codeColIndex))

            if not localisations:
                localisations = None

            if not localisations:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisations, workBook,
                       TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                       TSDApp)
    return check

def Test_02043_18_04939_COH_2006(workBook, TSDApp, DOC8List):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        codeColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Code défaut".casefold():
                codeColIndex = index
                break

        if codeColIndex != -1:
            localisations = []

            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow + 1):
                try:
                    cel = workSheet.cell(index, codeColIndex).value.split("-")
                    if len(cel) == 2:
                        pass
                    else:
                        if cel[0] not in DOC8List:
                            localisations.append(("codes défauts",index, codeColIndex))
                except:
                    localisations.append(("codes défauts",index, codeColIndex))

            if not localisations:
                localisations = None
            if not localisations:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisations, workBook,TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,TSDApp)
    return check

def Test_02043_18_04939_COH_2007(ExcelApp, workBook, TSDApp, DOC14Name):
    # testName = inspect.currentframe().f_code.co_name
    # print(testName)
    # check = False
    # if TSDApp.WorkbookStats.famillyList == "[]":
    #     result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    #     check = True
    # else:
    #
    #     DOC14 = xlrd.open_workbook(DOC14Name, on_demand=True)
    #     workSheetRef = DOC14.sheet_by_name("Matrix")
    #
    #     nrCols = workSheetRef.ncols
    #     nrRows = workSheetRef.nrows
    #     refColIndex = -1
    #     refRowIndex = -1
    #     var = 0
    #
    #     for index1 in range(0, nrRows):
    #         for index2 in range(0, nrCols):
    #             if str(workSheetRef.cell(index1, index2).value).casefold().strip() == "Data Trouble Code (DTC)".casefold():
    #                 refColIndex = index2
    #                 refRowIndex = index1
    #                 break
    #         if refColIndex != - 1 and refRowIndex != -1:
    #             break
    #
    #     if refColIndex == -1 or refRowIndex == -1:
    #         var = 1
    #
    #     if var == 1:
    #         result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    #         check = True
    #     elif var == 0:
    #         localisations = []
    #         flag = False
    #         list_ref  = []
    #
    #
    #         for index in range(refRowIndex + 1, nrRows):
    #             if workSheetRef.cell(index, refColIndex).value == None:
    #                 pass
    #             else:
    #                 list_ref.append(workSheetRef.cell(index, refColIndex).value)
    #
    #         for element in TSDApp.WorkbookStats.famillyList:
    #             if element["codenr"] in list_ref:
    #                 pass
    #             else:
    #                localisation.append(element["localisation"])
    #                check = True
    #
    #
    #         if not localisation:
    #             localisation = None
    #
    #         result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    #     return check


    testName = inspect.currentframe().f_code.co_name
    print(testName)
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

        for cellRow in workSheetRange.Rows:
            for cell in cellRow.cell:
                if str(cell.value).casefold().strip() == "Data Trouble Code (DTC)".casefold():
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
            refCellRange = workSheetRef.cell(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = []
            flag = False
            list_ref = []

            for index in range(refRowIndex + nrLines, nrRows + 1):
                if workSheetRef.cell(index, refColIndex).value == "":
                    pass
                else:
                    list_ref.append(workSheetRef.cell(index, refColIndex).value)

            for element in TSDApp.WorkbookStats.famillyList:
                if element["codenr"] in list_ref:
                    pass
                else:
                    localisation.append(element["localisation"])
                    check = True

            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook,
                   TSDApp)
        return check

def Test_02043_18_04939_COH_2010(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1
        list_code = list()
        tempList = list()
        var = 0

        for index in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Code défaut".casefold():
                refColIndex = index
                break

        if refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif refColIndex != -1:
            localisations = list()

            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, refColIndex).value == "NO DTC" or workSheet.cell(index, refColIndex).value == "":
                    pass
                else:
                    list_table = dict()
                    list_table["value"] = workSheet.cell(index, refColIndex).value
                    list_table["row"] = index
                    list_table["col"] = refColIndex
                    tempList.append(dict(list_table))

            if TSDApp.WorkbookStats.hasCode == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
                codeColIndex = -1

                for index in range(0, TSDApp.WorkbookStats.codeLastCol):
                    if str(workSheet.cell(TSDApp.codeHeaderRow,index).value).casefold().strip() == "Code défaut".casefold():
                        codeColIndex = index
                        break

                for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                    if workSheet.cell(index, codeColIndex).value == "":
                        pass
                    else:
                        list_code.append(workSheet.cell(index, codeColIndex).value.strip())

            for element in tempList:
                if ',' in element["value"]:
                    elem = element["value"].split(",")
                    for i in elem:
                        if i.strip() in list_code:
                            pass
                        else:
                            localisations.append(("tableau",element["row"],element["col"]))
                            check = True
                else:
                    if ';' in element["value"]:
                        elem = element["value"].split(";")
                        for i in elem:
                            if i.strip() in list_code:
                                pass
                            else:
                                localisations.append(("tableau", element["row"], element["col"]))
                                check = True
                    else:
                        if element["value"].strip() in list_code or element["value"] in list_code:
                            pass
                        else:
                            localisations.append(("tableau",element["row"],element["col"]))

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2020(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Constituant défaillant détecté".casefold():
                refColIndex = index
                break

        if refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif refColIndex != -1:

            localisations = list()
            list_table = dict()
            list_constituants = list()
            tempList = list()

            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, refColIndex).value == "":
                    pass
                else:
                    list_table["value"] = workSheet.cell(index, refColIndex).value
                    list_table["row"] = index
                    list_table["col"] = refColIndex
                    tempList.append(dict(list_table))

            if TSDApp.WorkbookStats.hasConstituants == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.constituantsIndex)
                constituantsColIndex = -1

                for index in range(0, TSDApp.WorkbookStats.constituantsLastCol ):
                    if str(workSheet.cell(TSDApp.constituantsHeaderRow, index).value).casefold().strip() == "Noms".casefold():
                        constituantsColIndex = index
                        break

                for index in range(TSDApp.constituantsFirstInfoRow, TSDApp.WorkbookStats.constituantsLastRow):
                    if workSheet.cell(index, constituantsColIndex).value == "":
                        pass
                    else:
                        list_constituants.append(workSheet.cell(index, constituantsColIndex).value)

            for element in tempList:
                if element["value"] in list_constituants:
                    pass
                else:
                    localisations.append(("tableau",element["row"],element["col"]))
                    check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2030(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Effet(s) client(s)".casefold():
                refColIndex = index
                break

        if refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif refColIndex != -1:
            localisations = list()
            list_table = dict()
            list_eff = list()
            tempList = list()


            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, refColIndex).value == "":
                    pass
                else:
                    list_table["value"] = workSheet.cell(index, refColIndex).value
                    list_table["row"] = index
                    list_table["col"] = refColIndex
                    tempList.append(dict(list_table))

            if TSDApp.WorkbookStats.hasEffClients == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.EffClientsIndex)
                effColIndex = -1

                for index in range(0, TSDApp.WorkbookStats.EffClientsLastCol):
                    if str(workSheet.cell(TSDApp.effClientsHeaderRow, index).value).casefold().strip() == "Noms".casefold():
                        effColIndex = index
                        break

                for index in range(TSDApp.effClientsFirstInfoRow, TSDApp.WorkbookStats.EffClientsLastRow):
                    if workSheet.cell(index, effColIndex).value == "":
                        pass
                    else:
                        list_eff.append(workSheet.cell(index, effColIndex).value.strip())

            for element in tempList:
                if element["value"].strip() in list_eff:
                    pass
                else:
                    localisations.append(("tableau", element["row"], element["col"]))
                    check = True

            if  not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2040(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "DIAGNOSTIC DEBARQUE".casefold():
                refColIndex = index
                break

        if refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif refColIndex != -1:
            localisations = list()
            list_table = dict()
            list_diag = list()
            tempList = list()

            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, refColIndex).value == "N/A" or workSheet.cell(index, refColIndex).value == "":
                    pass
                else:
                    list_table["value"] = workSheet.cell(index, refColIndex).value
                    list_table["row"] = index
                    list_table["col"] = refColIndex
                    tempList.append(dict(list_table))

            if TSDApp.WorkbookStats.hasDiagDeb == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagDebIndex)
                diagColIndex = -1

                for index in range(1, TSDApp.WorkbookStats.DiagDebLastCol + 1):
                    if str(workSheet.cell(TSDApp.diagDebHeaderRow, index).value).casefold().strip() == "libellé (signification)".casefold():
                        diagColIndex = index
                        break

                for index in range(TSDApp.diagDebFirstInfoRow, TSDApp.WorkbookStats.DiagDebLastRow):
                    if workSheet.cell(index, diagColIndex).value == "":
                        pass
                    else:
                        list_diag.append(workSheet.cell(index, diagColIndex).value)

            for element in tempList:
                if element["value"] in list_diag:
                    pass
                else:
                    localisations.append(("tableau",element["row"],element["col"]))
                    check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2050(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Evenement(s) redouté(s) (ER)".casefold():
                refColIndex = index
                break

        if refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif refColIndex != -1:

            localisations = list()
            list_table = dict()
            list_ER = list()
            tempList = list()

            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, refColIndex).value == "No DTC" or workSheet.cell(index, refColIndex).value == "":
                    pass
                else:
                    list_table["value"] = workSheet.cell(index, refColIndex).value
                    list_table["row"] = index
                    list_table["col"] = refColIndex
                    tempList.append(dict(list_table))

            if TSDApp.WorkbookStats.hasER == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ERIndex)
                ERColIndex = -1

                for index in range(0, TSDApp.WorkbookStats.ERLastCol):
                    if str(workSheet.cell(TSDApp.ERHeaderRow, index).value).casefold().strip() == "nom".casefold():
                        ERColIndex = index
                        break

                for index in range(TSDApp.ERFirstInfoRow, TSDApp.WorkbookStats.ERLastRow):
                    if workSheet.cell(index, ERColIndex).value == "":
                        pass
                    else:
                        list_ER.append(workSheet.cell(index, ERColIndex).value)

            for element in tempList:
                if element["value"] in list_ER:
                    pass
                else:
                    localisations.append(("tableau",element["row"],element["col"]))
                    check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2060(ExcelApp, workBook, TSDApp, DOC7Name):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.EffClientsIndex)

        effColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.EffClientsLastCol):
            if str(workSheet.cell(TSDApp.effClientsHeaderRow, index).value).casefold().strip() == "Noms".casefold():
                effColIndex = index
                break

        if effColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif effColIndex != -1:
            localisations = list()

            list_eff = list()
            list_ref = list()

            for index in range(TSDApp.effClientsFirstInfoRow, TSDApp.WorkbookStats.EffClientsLastRow):
                if workSheet.cell(index, effColIndex).value == "":
                    pass
                else:
                    dict = {}
                    dict["value"] = workSheet.cell(index, effColIndex).value
                    dict["row"] = index
                    dict["col"] = effColIndex
                    list_eff.append(dict)

            DOC7 = xlrd.open_workbook(DOC7Name, on_demand=True)
            if workBook.sheet_by_name("Effets clients"):
                workSheetRef = DOC7.sheet_by_name("FR")
            elif workBook.sheet_by_name("Customer Effects") or workBook.sheet_by_name("Customer Effect"):
                workSheetRef = DOC7.sheet_by_name("GB")


            nrCols = workSheetRef.ncols
            nrRows = workSheetRef.nrows
            N1ColIndex = -1
            N2ColIndex = -1
            N2ColIndex = -1
            col = 0
            for index1 in range(0, nrRows):
                for index2 in range(0, nrCols):
                    if str(workSheetRef.cell(index1, index2).value).strip() == "Libellé N1":
                        N1ColIndex = index2
                        N1RowIndex = index1
                        col += 1
                    if str(workSheetRef.cell(index1, index2).value).strip() == "Libellé N2":
                        N2ColIndex = index2
                        N2RowIndex = index1
                        col += 1
                    if str(workSheetRef.cell(index1, index2).value).strip() == "Libellé N3":
                        N3ColIndex = index2
                        N3RowIndex = index1
                        col += 1
                    if col == 3:
                        break
                if col == 3:
                    break


            if N1RowIndex != -1:
                for index in range(N1RowIndex + 1, nrRows):
                    if workSheetRef.cell(index, N1ColIndex).value == "":
                        pass
                    else:
                        list_ref.append(workSheetRef.cell(index, N1ColIndex).value)
            if N2RowIndex != -1:
                for index in range(N2RowIndex + 1, nrRows):
                    if workSheetRef.cell(index, N2ColIndex).value == "":
                        pass
                    else:
                        list_ref.append(workSheetRef.cell(index, N2ColIndex).value)
            if N3RowIndex != -1:
                for index in range(N3RowIndex + 1, nrRows):
                    if workSheetRef.cell(index, N3ColIndex).value == "":
                        pass
                    else:
                        list_ref.append(workSheetRef.cell(index, N3ColIndex).value)

            for element in list_eff:
                if element["value"] in list_ref:
                    pass
                else:
                    localisations.append(("Effets clients", element["row"],element["col"]))
                    check = True

            if not localisations:
                localisations = None

        if localisations is not None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

    # testName = inspect.currentframe().f_code.co_name
    # print(testName)
    # check = False
    # if TSDApp.WorkbookStats.hasEffClients == False:
    #     result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    #     check = True
    # else:
    #     workSheet = workBook.Sheets(TSDApp.WorkbookStats.EffClientsIndex)
    #     nrRows = workSheet.Rows
    #     effColIndex = 0
    #     var = 0
    #     for index1 in range(1, 15):
    #         for index2 in range(1, TSDApp.WorkbookStats.EffClientsLastCol + 1):
    #             if str(workSheet.cell(index1, index2).value).casefold().strip() == "Noms".casefold():
    #                 effColIndex = index2
    #                 effRowIndex = index1
    #                 break
    #         if effColIndex != 0:
    #             break
    #     if effColIndex == 0:
    #         var = 1
    #
    #     if var == 1:
    #         result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    #         check = True
    #     elif var == 0:
    #         effCellRange = workSheet.cell(effRowIndex, effColIndex).MergeArea
    #         nrLines = effCellRange.Rows.Count
    #         localisation = list()
    #
    #         list_eff = list()
    #         list_ref = list()
    #
    #         for index in range(effRowIndex + nrLines, TSDApp.WorkbookStats.EffClientsLastRow + 1):
    #             if workSheet.cell(index, effColIndex).value == None:
    #                 pass
    #             else:
    #                 dict = {}
    #                 dict["value"] = workSheet.cell(index, effColIndex).value
    #                 dict["localisation"] = workSheet.cell(index, effColIndex)
    #                 list_eff.append(dict)
    #
    #         DOC7 = ExcelApp.Workbooks.Open(DOC7Name)
    #         if workBook.Sheets("Effets clients"):
    #             workSheetRef = DOC7.Sheets("FR")
    #         elif workBook.Sheets("Customer Effects") or workBook.Sheets("Customer Effect"):
    #             workSheetRef = DOC7.Sheets("GB")
    #
    #         workSheetRange = workSheetRef.UsedRange
    #         nrCols = workSheetRange.Columns.Count
    #         nrRows = workSheetRange.Rows.Count
    #         N1ColIndex = 0
    #         N2ColIndex = 0
    #         N2ColIndex = 0
    #         col = 0
    #         for cellRow in workSheetRange.value:
    #             for cell in cellRow:
    #                 if cell == "Libellé N1":
    #                     N1ColIndex = cellRow.index(cell) + 1
    #                     N1RowIndex = workSheetRange.value.index(cellRow) + 1
    #                     col += 1
    #                 if cell == "Libellé N2":
    #                     N2ColIndex = cellRow.index(cell) + 1
    #                     N2RowIndex = workSheetRange.value.index(cellRow) + 1
    #                     col += 1
    #                 if cell == "Libellé N3":
    #                     N3ColIndex = cellRow.index(cell) + 1
    #                     N3RowIndex = workSheetRange.value.index(cellRow) + 1
    #                     col += 1
    #                 if col == 3:
    #                     break
    #             if col == 3:
    #                 break
    #
    #         try:
    #             refCellRange = workSheetRef.cell(N1RowIndex, N1ColIndex).MergeArea
    #         except:
    #             try:
    #                 refCellRange = workSheetRef.cell(N2RowIndex, N2ColIndex).MergeArea
    #             except:
    #                 refCellRange = workSheetRef.cell(N3RowIndex, N3ColIndex).MergeArea
    #
    #
    #         nrLines = refCellRange.Rows.Count
    #         localisation = list()
    #         flag = False
    #
    #         if N1RowIndex != 0:
    #             for index in range(N1RowIndex + nrLines, nrRows + 1):
    #                 if workSheetRef.cell(index, N1ColIndex).value == None:
    #                     pass
    #                 else:
    #                     list_ref.append(workSheetRef.cell(index, N1ColIndex).value)
    #         if N2RowIndex != 0:
    #             for index in range(N2RowIndex + nrLines, nrRows + 1):
    #                 if workSheetRef.cell(index, N2ColIndex).value == None:
    #                     pass
    #                 else:
    #                     list_ref.append(workSheetRef.cell(index, N2ColIndex).value)
    #         if N3RowIndex != 0:
    #             for index in range(N3RowIndex + nrLines, nrRows + 1):
    #                 if workSheetRef.cell(index, N3ColIndex).value == None:
    #                     pass
    #                 else:
    #                     list_ref.append(workSheetRef.cell(index, N3ColIndex).value)
    #
    #         for element in list_eff:
    #             if element["value"] in list_ref:
    #                 pass
    #             else:
    #                 localisation.append(element["localisation"])
    #                 check = True
    #
    #         if not localisation:
    #             localisation = None
    #
    #     if localisation is not None:
    #         result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    #     else:
    #         result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    #         check = True
    # return check

def Test_02043_18_04939_COH_2070(ExcelApp, workBook, TSDApp, DOC7Name):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.EffClientsIndex)

        effColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.EffClientsLastCol):
            if str(workSheet.cell(TSDApp.effClientsHeaderRow, index).value).casefold().strip() == "Name".casefold():
                effColIndex = index
                break

        if effColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif effColIndex != -1:
            localisations = list()

            list_eff = list()
            list_ref = list()

            for index in range(TSDApp.effClientsFirstInfoRow, TSDApp.WorkbookStats.EffClientsLastRow):
                if workSheet.cell(index, effColIndex).value == "":
                    pass
                else:
                    dict = {}
                    dict["value"] = workSheet.cell(index, effColIndex).value
                    dict["row"] = index
                    dict["col"] = effColIndex
                    list_eff.append(dict)

            DOC7 = xlrd.open_workbook(DOC7Name, on_demand=True)
            if workBook.sheet_by_name("Effets clients"):
                workSheetRef = DOC7.sheet_by_name("FR")
            elif workBook.sheet_by_name("Customer Effects") or workBook.sheet_by_name("Customer Effect"):
                workSheetRef = DOC7.sheet_by_name("GB")

            nrCols = workSheetRef.ncols
            nrRows = workSheetRef.nrows
            N1ColIndex = -1
            N2ColIndex = -1
            N2ColIndex = -1
            col = 0
            for index1 in range(0, nrRows):
                for index2 in range(0, nrCols):
                    if str(workSheetRef.cell(index1, index2).value).strip() == "Libellé N1":
                        N1ColIndex = index2
                        N1RowIndex = index1
                        col += 1
                    if str(workSheetRef.cell(index1, index2).value).strip() == "Libellé N2":
                        N2ColIndex = index2
                        N2RowIndex = index1
                        col += 1
                    if str(workSheetRef.cell(index1, index2).value).strip() == "Libellé N3":
                        N3ColIndex = index2
                        N3RowIndex = index1
                        col += 1
                    if col == 3:
                        break
                if col == 3:
                    break

            if N1RowIndex != -1:
                for index in range(N1RowIndex + 1, nrRows):
                    if workSheetRef.cell(index, N1ColIndex).value == "":
                        pass
                    else:
                        list_ref.append(workSheetRef.cell(index, N1ColIndex).value)
            if N2RowIndex != -1:
                for index in range(N2RowIndex + 1, nrRows):
                    if workSheetRef.cell(index, N2ColIndex).value == "":
                        pass
                    else:
                        list_ref.append(workSheetRef.cell(index, N2ColIndex).value)
            if N3RowIndex != -1:
                for index in range(N3RowIndex + 1, nrRows):
                    if workSheetRef.cell(index, N3ColIndex).value == "":
                        pass
                    else:
                        list_ref.append(workSheetRef.cell(index, N3ColIndex).value)

            for element in list_eff:
                if element["value"] in list_ref:
                    pass
                else:
                    localisations.append(("Customer Effects", element["row"], element["col"]))
                    check = True

            if not localisations:
                localisations = None

        if localisations is not None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_COH_2080(ExcelApp, workBook, TSDApp, DOC7Name):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.EffClientsIndex)

        effColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.EffClientsLastCol):
            if str(workSheet.cell(TSDApp.effClientsHeaderRow, index).value).casefold().strip() == "Noms".casefold():
                effColIndex = index
                break

        if effColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif effColIndex != -1:
            localisations = list()

            list_eff = list()
            list_ref = list()

            for index in range(TSDApp.effClientsFirstInfoRow, TSDApp.WorkbookStats.EffClientsLastRow):
                if workSheet.cell(index, effColIndex).value == "":
                    pass
                else:
                    dict = {}
                    dict["value"] = workSheet.cell(index, effColIndex).value
                    dict["row"] = index
                    dict["col"] = effColIndex
                    list_eff.append(dict)

            DOC7 = xlrd.open_workbook(DOC7Name, on_demand=True)
            if workBook.sheet_by_name("Effets clients"):
                workSheetRef = DOC7.sheet_by_name("FR")
            elif workBook.sheet_by_name("Customer Effects") or workBook.sheet_by_name("Customer Effect"):
                workSheetRef = DOC7.sheet_by_name("GB")

            nrCols = workSheetRef.ncols
            nrRows = workSheetRef.nrows
            N1ColIndex = -1
            N2ColIndex = -1
            N2ColIndex = -1
            col = 0
            for index1 in range(0, nrRows):
                for index2 in range(0, nrCols):
                    if str(workSheetRef.cell(index1, index2).value).strip() == "Libellé N1":
                        N1ColIndex = index2
                        N1RowIndex = index1
                        col += 1
                    if str(workSheetRef.cell(index1, index2).value).strip() == "Libellé N2":
                        N2ColIndex = index2
                        N2RowIndex = index1
                        col += 1
                    if str(workSheetRef.cell(index1, index2).value).strip() == "Libellé N3":
                        N3ColIndex = index2
                        N3RowIndex = index1
                        col += 1
                    if col == 3:
                        break
                if col == 3:
                    break

            if N1RowIndex != -1:
                for index in range(N1RowIndex + 1, nrRows):
                    if workSheetRef.cell(index, N1ColIndex).value == "":
                        pass
                    else:
                        list_ref.append(workSheetRef.cell(index, N1ColIndex).value)
            if N2RowIndex != -1:
                for index in range(N2RowIndex + 1, nrRows):
                    if workSheetRef.cell(index, N2ColIndex).value == "":
                        pass
                    else:
                        list_ref.append(workSheetRef.cell(index, N2ColIndex).value)
            if N3RowIndex != -1:
                for index in range(N3RowIndex + 1, nrRows):
                    if workSheetRef.cell(index, N3ColIndex).value == "":
                        pass
                    else:
                        list_ref.append(workSheetRef.cell(index, N3ColIndex).value)

            for element in list_eff:
                if element["value"] in list_ref:
                    pass
                else:
                    localisations.append(("Effets clients", element["row"], element["col"]))
                    check = True

            if not localisations:
                localisations = None

        if localisations is not None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                   TSDApp)
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_COH_2091(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
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
        firtCell = workSheet.cell(1, 1)
        lastCell = workSheet.cell(nrLines, nrCols)
        workSheetRange = workSheet.Range(firtCell, lastCell)
        flag = False

        for row in workSheetRange.Rows:
            flag = False
            for valueTuple in row.value:
                for value in valueTuple:
                    if value != None:
                        flag = True
            if flag == False:
                lastRow = row.Row
                break

            for rowIndex in range(1, nrLines):
                for colIndex in range(1, nrCols):
                    if workSheet.cell(rowIndex, colIndex).value == "?" or  workSheet.cell(rowIndex, colIndex).value == "tbd" or workSheet.cell(rowIndex, colIndex).value == "tbc":
                        localisation.append(workSheet.cell(rowIndex, colIndex))
                        check = True

    if localisation == "[]":
        localisation = None

    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2100(workBook, TSDApp, DOC8List):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
        localisations = []
        codeColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.codeLastCol):
            if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "supporté par constituant (s)".casefold():
                codeColIndex = index
                break

        if codeColIndex != -1:
            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                cel = workSheet.cell(index, codeColIndex).value
                if cel not in DOC8List:
                    localisations.append(("codes défauts", index, codeColIndex))

            if not localisations:
                localisations = None
            if not localisations:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisations, workBook,TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,TSDApp)
    return check

def Test_02043_18_04939_COH_2110(workBook, TSDApp, DOC8List):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.measureIndex)
        ColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.measureLastCol):
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "supporté par constituant (s)".casefold():
                ColIndex = index
                break

        if ColIndex != -1:
            localisations = []

            for index in range(TSDApp.measureFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                cel = workSheet.cell(index, ColIndex).value
                if cel is "":
                    pass
                else:
                    if cel not in DOC8List:
                        localisations.append(("mesures et commandes",index, ColIndex))

            if not localisations:
                localisations = None
            if not localisations:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisations, workBook,TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,TSDApp)
    return check

def Test_02043_18_04939_COH_2120(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasReqTech == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ReqTechIndex)
        refColIndex = -1
        list_amont = list()
        tempDict = list()
        localisations = list()

        for index in range (0, TSDApp.WorkbookStats.ReqTechLastCol):
            if str(workSheet.cell(TSDApp.reqTechHeaderRow, index).value).casefold().strip() == "Reference".casefold():
                refColIndex = index
                break

        if refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            for index in range(TSDApp.reqTechFirstInfoRow, TSDApp.WorkbookStats.ReqTechLastRow):
                if workSheet.cell(index, refColIndex).value == "":
                    pass
                else:
                    dict = {}
                    dict["value"] = workSheet.cell(index, refColIndex).value
                    dict["row"] = index
                    dict["col"] = refColIndex
                    tempDict.append(dict)


            DOC5 = xlrd.open_workbook(DOC5Name, on_demand=True)
            workSheetRef = DOC5.sheet_by_name("Effets techniques")
            nrCols = workSheetRef.ncols
            nrRows = workSheetRef.nrows
            amontColIndex = -1
            amontRowIndex = -1

            for index1 in range(0, nrRows):
                for index2 in range(0, nrCols):
                    if str(workSheetRef.cell(index1, index2).value).casefold().strip() == "Référence amont".casefold():
                        amontColIndex = index2
                        amontRowIndex = index1
                        break
                if amontColIndex != -1 and amontRowIndex != -1:
                    break

            for index in range(amontRowIndex + 1, nrRows):
                if workSheetRef.cell(index, amontColIndex).value == "":
                    pass
                else:
                    list_amont.append(workSheetRef.cell(index, amontColIndex).value)

            for element in tempDict:
                if element in list_amont:
                    pass
                else:
                   localisations.append(("Req. of tech. effects",element["row"], element["col"]))
                   check = True

            if not localisations:
                localisations = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2130(workBook, TSDApp):
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
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Référence".casefold() or str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Reference".casefold():
                refColIndex = index
                break

        if refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif refColIndex != -1:
            list_table = list()
            list_effets = list()

            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, refColIndex).value == "":
                    pass
                else:
                    dict = {}
                    dict["value"] = workSheet.cell(index, refColIndex).value
                    dict["row"] = index
                    dict["col"] = refColIndex
                    list_table.append(dict)

            if TSDApp.WorkbookStats.hasTechEff == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheetRef = workBook.sheet_by_index(TSDApp.WorkbookStats.TechEffIndex)
                nrCols = workSheetRef.ncols
                nrRows = workSheetRef.nrows
                effColIndex = -1

                for index1 in range(0, nrRows):
                    for index2 in range(0, nrCols):
                        if str(workSheetRef.cell(index1, index2).value).casefold().strip() == "Référence amont".casefold():
                            effColIndex = index2
                            effRowIndex = index1
                            break
                    if effColIndex != -1 and effRowIndex != -1:
                        break

                if effColIndex != -1 and effRowIndex != -1:
                    for index in range(TSDApp.techEffFirstInfoRow, nrRows):
                        if workSheet.cell(index, effColIndex).value == "":
                            pass
                        else:
                            list_effets.append(workSheet.cell(index, effColIndex).value)

                    for element in list_table:
                        if element["value"] in list_effets:
                            pass
                        else:
                            localisations.append(("tableau"),element["row"],element["col"])
                            check = True

                    if not localisations:
                        localisations = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook, TSDApp)

                elif effColIndex == 0:
                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                    check =True
    return check

def Test_02043_18_04939_COH_2140(workBook, TSDApp):
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
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "code défauts induits".casefold():
                refColIndex = index
                break

        if refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif refColIndex != -1:
            list_table = list()
            list_effets = list()

            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                list_table_dict = {}
                list_table_dict["value"] = workSheet.cell(index, refColIndex).value
                list_table_dict["row"] = index
                list_table_dict["col"] = refColIndex
                list_table.append(dict(list_table_dict))

            if TSDApp.WorkbookStats.hasCode == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.codeIndex)
                effColIndex = -1

                for index in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                    if str(workSheet.cell(TSDApp.codeHeaderRow, index).value).casefold().strip() == "Code défaut".casefold():
                        effColIndex = index
                        break

                if effColIndex != -1:
                    for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                        if workSheet.cell(index, effColIndex).value == "":
                            pass
                        else:
                            list_effets.append(workSheet.cell(index, effColIndex).value)

                    for element in list_table:
                        if element["value"] in list_effets or element["value"] == "N/A":
                            pass
                        else:
                            localisations.append(("tableau", element["row"], element["col"]))
                            check = True

                    if not localisations:
                        localisations = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook, TSDApp)

                elif effColIndex == 0:
                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                    check = True
    return check

def Test_02043_18_04939_COH_2150(workBook, TSDApp):
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
            if str(workSheet.cell(TSDApp.codeHeaderRow,index).value).casefold().strip() == "supporté par constituant (s)".casefold():
                refColIndex = index
                break

        if refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif refColIndex != -1:
            list_table = list()
            list_effets = list()

            for index in range(TSDApp.codeFirstInfoRow, TSDApp.WorkbookStats.codeLastRow):
                if workSheet.cell(index, refColIndex).value == "":
                    pass
                else:
                    list_table_dict = {}
                    list_table_dict["value"] = workSheet.cell(index, refColIndex).value
                    list_table_dict["row"] = index
                    list_table_dict["col"] = refColIndex
                    list_table.append(dict(list_table_dict))

            if TSDApp.WorkbookStats.hasConstituants == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.constituantsIndex)
                effColIndex = -1

                for index in range(0, TSDApp.WorkbookStats.constituantsLastCol):
                    if str(workSheet.cell(TSDApp.constituantsHeaderRow, index).value).casefold().strip() == "Noms".casefold():
                        effColIndex = index
                        break

                if effColIndex != -1:
                    for index in range(TSDApp.constituantsFirstInfoRow, TSDApp.WorkbookStats.constituantsLastRow):
                        if workSheet.cell(index, effColIndex).value == "":
                            pass
                        else:
                            list_effets.append(workSheet.cell(index, effColIndex).value)

                    for element in list_table:
                        if element["value"] in list_effets:
                            pass
                        else:
                            localisations.append(("codes défauts",element["row"],element["col"]))
                            check = True

                    if not localisations:
                        localisations = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations,
                           workBook, TSDApp)

                elif effColIndex == -1:
                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                    check = True
    return check

def Test_02043_18_04939_COH_2160(workBook, TSDApp):
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
            if str(workSheet.cell(TSDApp.measureHeaderRow, index).value).casefold().strip() == "supporté par constituant (s)".casefold():
                refColIndex = index
                break

        if refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif refColIndex != -1:
            list_table = list()
            list_effets = list()

            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.measureLastRow):
                if workSheet.cell(index, refColIndex).value == "":
                    pass
                else:
                    list_table_dict = {}
                    list_table_dict["value"] = workSheet.cell(index, refColIndex).value
                    list_table_dict["row"] = index
                    list_table_dict["col"] = refColIndex
                    list_table.append(dict(list_table_dict))

            if TSDApp.WorkbookStats.hasConstituants == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.constituantsIndex)
                effColIndex = -1

                for index in range(0, TSDApp.WorkbookStats.constituantsLastCol):
                    if str(workSheet.cell(TSDApp.constituantsHeaderRow, index).value).casefold().strip() == "Noms".casefold():
                        effColIndex = index
                        break

                if effColIndex != -1:
                    for index in range(TSDApp.constituantsFirstInfoRow, TSDApp.WorkbookStats.constituantsLastRow):
                        if workSheet.cell(index, effColIndex).value == "":
                            pass
                        else:
                            list_effets.append(workSheet.cell(index, effColIndex).value)

                    for element in list_table:
                        if element["value"] in list_effets:
                            pass
                        else:
                            localisations.append(("measure et commandes",element["row"], element["col"]))
                            check = True

                    if not localisations:
                        localisations = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations,workBook, TSDApp)

                elif effColIndex == 0:
                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                    check = True
    return check

def Test_02043_18_04939_COH_2170(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasDataCodes == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DataCodesIndex)
        refColIndex = -1
        localisations = list()

        for index in range(0, TSDApp.WorkbookStats.DataCodesLastCol):
            if str(workSheet.cell(TSDApp.dataCodesHeaderRow, index).value).casefold().strip() == "Stored by the ECU".casefold():
                refColIndex = index
                break

        if refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif refColIndex != -1:
            list_table = list()
            list_effets = list()

            for index in range(TSDApp.dataCodesFirstInfoRow, TSDApp.WorkbookStats.DataCodesLastRow):
                if workSheet.cell(index, refColIndex).value == "":
                    pass
                else:
                    dict = {}
                    dict["value"] = workSheet.cell(index, refColIndex).value
                    dict["row"] = index
                    dict["col"] = refColIndex
                    list_table.append(dict)

            if TSDApp.WorkbookStats.hasParts == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                effColIndex = -1
                for index in range(0, TSDApp.WorkbookStats.PartsLastCol):
                    if str(workSheet.cell(TSDApp.partsHeaderRow, index).value).casefold().strip() == "Name".casefold():
                        effColIndex = index
                        break

                if effColIndex != -1:
                    for index in range(TSDApp.partsFirstInfoRow, TSDApp.WorkbookStats.PartsLastRow):
                        if workSheet.cell(index, effColIndex).value == "":
                            pass
                        else:
                            list_effets.append(workSheet.cell(index, effColIndex).value)

                    for element in list_table:
                        if element["value"] in list_effets:
                            pass
                        else:
                            localisations.append(("Data trouble codes",element["row"],element["col"]))
                            check = True

                    if not localisations:
                        localisations = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook, TSDApp)

                elif effColIndex == -1:
                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                    check = True
    return check

def Test_02043_18_04939_COH_2180(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasReadDataIO == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = -1
        localisations = list()

        for index in range(0, TSDApp.WorkbookStats.ReadDataIOLastCol):
            if str(workSheet.cell(TSDApp.readDataIOHeaderRow, index).value).casefold().strip() == "Stored by the ECU".casefold():
                refColIndex = index
                break

        if refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif refColIndex != -1:
            list_table = list()
            list_effets = list()

            for index in range(TSDApp.readDataIOFirstInfoRow, TSDApp.WorkbookStats.ReadDataIOLastRow):
                if workSheet.cell(index, refColIndex).value == "":
                    pass
                else:
                    dict = {}
                    dict["value"] = workSheet.cell(index, refColIndex).value
                    dict["row"] = index
                    dict["col"] = refColIndex
                    list_table.append(dict)

            if TSDApp.WorkbookStats.hasParts == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                effColIndex = -1

                for index in range(0, TSDApp.WorkbookStats.PartsLastCol):
                    if str(workSheet.cell(TSDApp.partsHeaderRow, index).value).casefold().strip() == "Name".casefold():
                        effColIndex = index
                        break

                if effColIndex != -1:
                    for index in range(TSDApp.partsFirstInfoRow, TSDApp.WorkbookStats.PartsLastRow):
                        if workSheet.cell(index, effColIndex).value == "":
                            pass
                        else:
                            list_effets.append(workSheet.cell(index, effColIndex).value)

                    for element in list_table:
                        if element["value"] in list_effets:
                            pass
                        else:
                            localisations.append(("Read data and IO control"),element["row"],element["col"])
                            check = True

                    if not localisations:
                        localisations = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook, TSDApp)

                elif effColIndex == -1:
                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                    check = True
    return check

def Test_02043_18_04939_COH_2190(workBook, TSDApp):
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
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "situation de vie".casefold():
                refColIndex = index
                break

        if refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif refColIndex != -1:
            list_table = list()
            list_effets = list()

            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, refColIndex).value == "":
                    pass
                else:
                    list_table_dict = {}
                    list_table_dict["value"] = workSheet.cell(index, refColIndex).value
                    list_table_dict["row"] = index
                    list_table_dict["col"] = refColIndex
                    list_table.append(dict(list_table_dict))

            if TSDApp.WorkbookStats.hasSitDeVie == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.SitDeVieIndex)
                effColIndex = -1

                for index in range(0, TSDApp.WorkbookStats.SitDeVieLastCol):
                    if str(workSheet.cell(TSDApp.sitDeVieHeaderRow, index).value).casefold().strip() == "Situations de vie".casefold():
                        effColIndex = index
                        break

                if effColIndex != -1:
                    for index in range(TSDApp.sitDeVieFirstInfoRow, TSDApp.WorkbookStats.SitDeVieLastRow):
                        if workSheet.cell(index, effColIndex).value == "":
                            pass
                        else:
                            list_effets.append(workSheet.cell(index, effColIndex).value)

                    for element in list_table:
                        if element["value"] in list_effets or element["value"] == "N/A":
                            pass
                        else:
                            localisations.append(("tableau",element["row"],element["col"]))
                            check = True

                    if not localisations:
                        localisations = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations,
                           workBook, TSDApp)

                elif effColIndex == -1:
                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                    check = True
    return check

def Test_02043_18_04939_COH_2200(workBook, TSDApp):
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
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Situation".casefold():
                refColIndex = index
                break

        if refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True

        elif refColIndex != -1:
            list_table = list()
            list_effets = list()

            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                if workSheet.cell(index, refColIndex).value == "":
                    pass
                else:
                    list_table_dict = {}
                    list_table_dict["value"] = workSheet.cell(index, refColIndex).value
                    list_table_dict["row"] = index
                    list_table_dict["col"] = refColIndex
                    list_table.append(dict(list_table_dict))

            if TSDApp.WorkbookStats.hasSituation == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.SituationIndex)
                effColIndex = -1

                for index in range(0, TSDApp.WorkbookStats.SituationLastCol):
                    if str(workSheet.cell(TSDApp.situationHeaderRow, index).value).casefold().strip() == "Description".casefold():
                        effColIndex = index
                        break

                if effColIndex != -1:
                    for index in range(TSDApp.situationFirstInfoRow, TSDApp.WorkbookStats.SituationLastRow):
                        if workSheet.cell(index, effColIndex).value == "":
                            pass
                        else:
                            list_effets.append(workSheet.cell(index, effColIndex).value)

                    for element in list_table:
                        if element["value"] in list_effets or element["value"] == "N/A":
                            pass
                        else:
                            localisations.append(("Table"),element["row"],element["col"])
                            check = True

                    if not localisations:
                        localisations = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook, TSDApp)

                elif effColIndex == 0:
                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                    check = True
    return check

def Test_02043_18_04939_COH_2210(workBook, TSDApp):
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
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Diagnostic debarque".casefold():
                refColIndex = index
                break

        if refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif refColIndex != -1:
            list_table = list()
            list_effets = list()

            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                list_table_dict = {}
                if workSheet.cell(index, refColIndex).value is "":
                    pass
                else:
                    list_table_dict["value"] = workSheet.cell(index, refColIndex).value
                    list_table_dict["row"] = index
                    list_table_dict["col"] = refColIndex
                    list_table.append(dict(list_table_dict))

            if TSDApp.WorkbookStats.hasDiagDeb == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.DiagDebIndex)
                effColIndex = -1

                for index in range(0, TSDApp.WorkbookStats.DiagDebLastCol):
                    if str(workSheet.cell(TSDApp.diagDebHeaderRow, index).value).casefold().strip() == "libellé (signification)".casefold():
                        effColIndex = index
                        break

                if effColIndex != -1:
                    for index in range(TSDApp.diagDebFirstInfoRow, TSDApp.WorkbookStats.DiagDebLastRow):
                        if workSheet.cell(index, effColIndex).value == "":
                            pass
                        else:
                            list_effets.append(workSheet.cell(index, effColIndex).value)

                    for element in list_table:
                        if element["value"] in list_effets or element["value"] == "N/A":
                            pass
                        else:
                            localisations.append(("tableau",element["row"],element["col"]))
                            check = True

                    if not localisations:
                        localisations = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations,workBook, TSDApp)

                elif effColIndex == -1:
                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                    check = True
    return check

def Test_02043_18_04939_COH_2220(workBook, TSDApp):
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
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Non-embedded diagnosis".casefold():
                refColIndex = index
                break

        if refColIndex == -1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif refColIndex != -1:
            list_table = list()
            list_effets = list()

            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                list_table_dict = {}
                list_table_dict["value"] = workSheet.cell(index, refColIndex).value
                list_table_dict["row"] = index
                list_table_dict["col"] = refColIndex
                list_table.append(dict(list_table_dict))


            if TSDApp.WorkbookStats.hasNotEmbDiag == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.NotEmbDiagIndex)
                effColIndex = -1

                for index in range(0, TSDApp.WorkbookStats.NotEmbDiagLastCol):
                    if str(workSheet.cell(TSDApp.notEmbDiagHeaderRow, index).value).casefold().strip() == "Label".casefold():
                        effColIndex = index
                        break

                if effColIndex != -1:
                    for index in range(TSDApp.notEmbDiagFirstInfoRow, TSDApp.WorkbookStats.NotEmbDiagLastRow):
                        if workSheet.cell(index, effColIndex).value == "":
                            pass
                        else:
                            list_effets.append(workSheet.cell(index, effColIndex).value)

                    for element in list_table:
                        if element["value"] in list_effets or element["value"] == "N/A":
                            pass
                        else:
                            localisations.append(("Table"),element["row"],element["col"])
                            check = True

                    if not localisations:
                        localisations = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook, TSDApp)

                elif effColIndex == -1:
                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                    check = True
    return check

def Test_02043_18_04939_COH_2230(workBook, TSDApp, subfamily_name, DOC15List):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if subfamily_name is None and DOC15List is None:
        return True
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        refColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "mesures et commandes (Mesure Parametre et Test Actionneur) / Tests de cohérence".casefold():
                codeColIndex = index
                break

        if refColIndex != -1:
            localisations = []
            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                try:
                    cel = workSheet.cell(index, refColIndex).value.split("-")
                    if cel[0] == subfamily_name and cel[1].lstrip('_') in DOC15List:
                        pass
                    else:
                        localisations.append(("tableau",index, refColIndex))
                except:
                    localisations.append(("tableau", index, refColIndex))

            if not localisations:
                localisations = None
            if not localisations:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisations, workBook,
                       TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,
                       TSDApp)
    return check
    # testName = inspect.currentframe().f_code.co_name
    # print(testName)
    # check = False
    # if subfamily_name is None and DOC15List is None:
    #     return True
    # if TSDApp.WorkbookStats.hasTable == False:
    #     result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    #     check = True
    # else:
    #     workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
    #
    #     codeColIndex = 0
    #     codeRowIndex = 0
    #     var = 0
    #     for index1 in range(1, 15):
    #         for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
    #             if str(workSheet.cell(index1, index2).value).casefold().strip() == "mesures et commandes (Mesure Parametre et Test Actionneur) / Tests de cohérence".casefold():
    #                 codeColIndex = index2
    #                 codeRowIndex = index1
    #                 break
    #         if codeColIndex != 0:
    #             break
    #     if codeColIndex == 0:
    #         var = 1
    #
    #     if var == 0:
    #         refCellRange = workSheet.cell(codeRowIndex, codeColIndex).MergeArea
    #         nrLines = refCellRange.Rows.Count
    #         localisation = []
    #
    #         for index in range(codeRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
    #             try:
    #                 cel = workSheet.cell(index, codeColIndex).value.split("-")
    #                 if cel[0] == subfamily_name and cel[1].lstrip('_') in DOC15List:
    #                     pass
    #                 else:
    #                     localisation.append(workSheet.cell(index, codeColIndex))
    #             except:
    #                 localisation.append(workSheet.cell(index, codeColIndex))
    #
    #         if not localisation:
    #             localisation = None
    #         if not localisation:
    #             result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisation, workBook,TSDApp)
    #             check = True
    #         else:
    #             result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook,TSDApp)
    # return check

def Test_02043_18_04939_COH_2240(workBook, TSDApp, DOC13List):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        codeColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Variant/\noption".casefold() or str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Variante/\noption".casefold():
                codeColIndex = index
                break

        if codeColIndex != -1:
            localisations = []
            contor = 0

            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                list2 = ['AND', 'OR', "NOT", "N/A"]
                cel = []
                try:
                    cel = workSheet.cell(index, codeColIndex).value.split(" ")
                    list = []
                    for elem in cel:
                        objElem = {}
                        objElem['NAME'] = elem
                        objElem['CHECK'] = False
                        list.append(objElem)

                    check_list1 = False
                    for i in range(len(list)):
                        leng = len(list[i]['NAME'])
                        if leng == 0:
                            list[i]['CHECK'] = True

                        poz = 0
                        if list[i]['NAME'] == "(":
                            for j in range(i+1,len(list)):
                                if list[j]['NAME'] == '':
                                    poz = poz + 1
                                    list[i+poz]['CHECK'] = True
                                    check_list1 = True
                                else:
                                    for k in range(len(DOC13List)):
                                        new_val = DOC13List[k] + ')'
                                        if list[j]['NAME'] == DOC13List[k] or list[j]['NAME'] == new_val:
                                            list[i]['CHECK'] = True
                                            check_list1 = True
                                            break
                                    break

                        if list[i]['NAME'] == ")":
                            for j in range(i - 1, -1, -1):
                                if list[j]['NAME'] == '':
                                    poz = poz + 1
                                    list[i - poz]['CHECK'] = True
                                    check_list1 = True
                                else:
                                    for k in range(len(DOC13List)):
                                        new_val = '(' + DOC13List[k]
                                        if list[j]['NAME'] == DOC13List[k] or list[j]['NAME'] == new_val:
                                            list[i]['CHECK'] = True
                                            check_list1 = True
                                            break
                                    break

                        if leng > 1:
                            for j in range(len(DOC13List)):
                                if list[i]['NAME'][0] == '(' or list[i]['NAME'][-1] == ")":
                                    new_elem1 = list[i]['NAME'].replace("(", "").replace(")", "")
                                    if new_elem1 == DOC13List[j]:
                                        list[i]['CHECK'] = True
                                        check_list1 = True
                                        break
                                else:
                                    if list[i]['NAME'] == DOC13List[j]:
                                        list[i]['CHECK'] = True
                                        check_list1 = True
                                        break

                    check_list2 = False
                    for elem1 in list:
                        for elem2 in list2:
                            if elem1['NAME'] == elem2:
                                elem1['CHECK'] = True
                                check_list2 = True
                                break

                    cnt = 0
                    for elem in list:
                        if elem['CHECK'] == True:
                            cnt = cnt + 1
                    if cnt == len(list) and check_list1 == True and check_list2 == True:
                        contor = contor + 1
                    else:
                        localisations.append(("tableau",index,codeColIndex))
                except:
                    pass

            if not localisations:
                localisations = None

            if contor == TSDApp.WorkbookStats.tableLastRow - TSDApp.tableFirstInfoRow - 1:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisations, workBook,TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,TSDApp)
    return check

def Test_02043_18_04939_COH_2241(workBook, TSDApp, DOC13List):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        codeColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Diversity".casefold():
                codeColIndex = index
                break

        if codeColIndex != -1:
            localisations = []
            contor = 0

            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                list2 = ['AND', 'OR', "NOT", "N/A",""]
                cel = []
                try:
                    cel = workSheet.cell(index, codeColIndex).value.split(" ")
                    list = []
                    for elem in cel:
                        objElem = {}
                        objElem['NAME'] = elem
                        objElem['CHECK'] = False
                        list.append(objElem)

                    check_list1 = False
                    for i in range(len(list)):
                        leng = len(list[i]['NAME'])
                        if leng == 0:
                            list[i]['CHECK'] = True

                        poz = 0
                        if list[i]['NAME'] == "(":
                            for j in range(i+1,len(list)):
                                if list[j]['NAME'] == '':
                                    poz = poz + 1
                                    list[i+poz]['CHECK'] = True
                                    check_list1 = True
                                else:
                                    for k in range(len(DOC13List)):
                                        new_val = DOC13List[k] + ')'
                                        if list[j]['NAME'] == DOC13List[k] or list[j]['NAME'] == new_val:
                                            list[i]['CHECK'] = True
                                            check_list1 = True
                                            break
                                    break

                        if list[i]['NAME'] == ")":
                            for j in range(i - 1, -1, -1):
                                if list[j]['NAME'] == '':
                                    poz = poz + 1
                                    list[i - poz]['CHECK'] = True
                                    check_list1 = True
                                else:
                                    for k in range(len(DOC13List)):
                                        new_val = '(' + DOC13List[k]
                                        if list[j]['NAME'] == DOC13List[k] or list[j]['NAME'] == new_val:
                                            list[i]['CHECK'] = True
                                            check_list1 = True
                                            break
                                    break

                        if leng > 1:
                            for j in range(len(DOC13List)):
                                if list[i]['NAME'][0] == '(' or list[i]['NAME'][-1] == ")":
                                    new_elem1 = list[i]['NAME'].replace("(", "").replace(")", "")
                                    if new_elem1 == DOC13List[j]:
                                        list[i]['CHECK'] = True
                                        check_list1 = True
                                        break
                                else:
                                    if list[i]['NAME'] == DOC13List[j]:
                                        list[i]['CHECK'] = True
                                        check_list1 = True
                                        break

                    check_list2 = False
                    for elem1 in list:
                        for elem2 in list2:
                            if elem1['NAME'] == elem2:
                                elem1['CHECK'] = True
                                check_list2 = True
                                break

                    cnt = 0
                    for elem in list:
                        if elem['CHECK'] == True:
                            cnt = cnt + 1
                    if cnt == len(list) and check_list1 == True and check_list2 == True:
                        contor = contor + 1
                    else:
                        localisations.append(("Table",index,codeColIndex))
                except:
                    pass

            if not localisations:
                localisations = None


            if contor == TSDApp.WorkbookStats.tableLastRow - TSDApp.tableFirstInfoRow - 1:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisations, workBook,TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,TSDApp)
    return check

def Test_02043_18_04939_COH_2250(workBook, TSDApp, DOC13List):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        codeColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Variant/\noption".casefold() or str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Variante/\noption".casefold():
                codeColIndex = index
                break

        if codeColIndex != -1:
            localisations = []
            contor = 0

            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                list2 = ['AND', 'OR', "NOT", "N/A"]
                cel = []
                try:
                    cel = workSheet.cell(index, codeColIndex).value.split(" ")
                    list = []
                    for elem in cel:
                        objElem = {}
                        objElem['NAME'] = elem
                        objElem['CHECK'] = False
                        list.append(objElem)

                    check_list1 = False
                    for i in range(len(list)):
                        leng = len(list[i]['NAME'])
                        if leng == 0:
                            list[i]['CHECK'] = True

                        poz = 0
                        if list[i]['NAME'] == "(":
                            for j in range(i+1,len(list)):
                                if list[j]['NAME'] == '':
                                    poz = poz + 1
                                    list[i+poz]['CHECK'] = True
                                    check_list1 = True
                                else:
                                    for k in range(len(DOC13List)):
                                        new_val = DOC13List[k] + ')'
                                        if list[j]['NAME'] == DOC13List[k] or list[j]['NAME'] == new_val:
                                            list[i]['CHECK'] = True
                                            check_list1 = True
                                            break
                                    break

                        if list[i]['NAME'] == ")":
                            for j in range(i - 1, -1, -1):
                                if list[j]['NAME'] == '':
                                    poz = poz + 1
                                    list[i - poz]['CHECK'] = True
                                    check_list1 = True
                                else:
                                    for k in range(len(DOC13List)):
                                        new_val = '(' + DOC13List[k]
                                        if list[j]['NAME'] == DOC13List[k] or list[j]['NAME'] == new_val:
                                            list[i]['CHECK'] = True
                                            check_list1 = True
                                            break
                                    break

                        if leng > 1:
                            for j in range(len(DOC13List)):
                                if list[i]['NAME'][0] == '(' or list[i]['NAME'][-1] == ")":
                                    new_elem1 = list[i]['NAME'].replace("(", "").replace(")", "")
                                    if new_elem1 == DOC13List[j]:
                                        list[i]['CHECK'] = True
                                        check_list1 = True
                                        break
                                else:
                                    if list[i]['NAME'] == DOC13List[j]:
                                        list[i]['CHECK'] = True
                                        check_list1 = True
                                        break

                    check_list2 = False
                    for elem1 in list:
                        for elem2 in list2:
                            if elem1['NAME'] == elem2:
                                elem1['CHECK'] = True
                                check_list2 = True
                                break

                    cnt = 0
                    for elem in list:
                        if elem['CHECK'] == True:
                            cnt = cnt + 1
                    if cnt == len(list) and check_list1 == True and check_list2 == True:
                        contor = contor + 1
                    else:
                        localisations.append(("tableau",index,codeColIndex))
                except:
                    pass

            if not localisations:
                localisation = None

            if contor == TSDApp.WorkbookStats.tableLastRow - TSDApp.tableFirstInfoRow - 1:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisations, workBook,TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,TSDApp)
    return check

def Test_02043_18_04939_COH_2251(workBook, TSDApp, DOC13List):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.sheet_by_index(TSDApp.WorkbookStats.tableIndex)
        codeColIndex = -1

        for index in range(0, TSDApp.WorkbookStats.tableLastCol):
            if str(workSheet.cell(TSDApp.tableHeaderRow, index).value).casefold().strip() == "Diversity".casefold():
                codeColIndex = index
                break

        if codeColIndex != -1:
            localisations = []
            contor = 0

            for index in range(TSDApp.tableFirstInfoRow, TSDApp.WorkbookStats.tableLastRow):
                list2 = ['AND', 'OR', "NOT", "N/A"]
                try:
                    cel = cel.replace(",", "").replace(";", "")
                    cel = workSheet.cell(index, codeColIndex).value.split(" ")
                    list = []
                    for elem in cel:
                        objElem = {}
                        objElem['NAME'] = elem
                        objElem['CHECK'] = False
                        list.append(objElem)

                    check_list1 = False
                    for i in range(len(list)):
                        leng = len(list[i]['NAME'])
                        if leng == 0:
                            list[i]['CHECK'] = True

                        poz = 0
                        if list[i]['NAME'] == "(":
                            for j in range(i+1,len(list)):
                                if list[j]['NAME'] == '':
                                    poz = poz + 1
                                    list[i+poz]['CHECK'] = True
                                    check_list1 = True
                                else:
                                    for k in range(len(DOC13List)):
                                        new_val = DOC13List[k] + ')'
                                        if list[j]['NAME'] == DOC13List[k] or list[j]['NAME'] == new_val:
                                            list[i]['CHECK'] = True
                                            check_list1 = True
                                            break
                                    break

                        if list[i]['NAME'] == ")":
                            for j in range(i - 1, -1, -1):
                                if list[j]['NAME'] == '':
                                    poz = poz + 1
                                    list[i - poz]['CHECK'] = True
                                    check_list1 = True
                                else:
                                    for k in range(len(DOC13List)):
                                        new_val = '(' + DOC13List[k]
                                        if list[j]['NAME'] == DOC13List[k] or list[j]['NAME'] == new_val:
                                            list[i]['CHECK'] = True
                                            check_list1 = True
                                            break
                                    break

                        if leng > 1:
                            for j in range(len(DOC13List)):
                                if list[i]['NAME'][0] == '(' or list[i]['NAME'][-1] == ")":
                                    new_elem1 = list[i]['NAME'].replace("(", "").replace(")", "")
                                    if new_elem1 == DOC13List[j]:
                                        list[i]['CHECK'] = True
                                        check_list1 = True
                                        break
                                else:
                                    if list[i]['NAME'] == DOC13List[j]:
                                        list[i]['CHECK'] = True
                                        check_list1 = True
                                        break

                    check_list2 = False
                    for elem1 in list:
                        for elem2 in list2:
                            if elem1['NAME'] == elem2:
                                elem1['CHECK'] = True
                                check_list2 = True
                                break

                    cnt = 0
                    for elem in list:
                        if elem['CHECK'] == True:
                            cnt = cnt + 1
                    if cnt == len(list) and check_list1 == True and check_list2 == True:
                        contor = contor + 1
                    else:
                        localisations.append(("Table",index,codeColIndex))

                except:
                    pass

            if not localisations:
                localisations = None

            if contor == TSDApp.WorkbookStats.tableLastRow - TSDApp.tableFirstInfoRow - 1:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisations, workBook,TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisations, workBook,TSDApp)
    return check