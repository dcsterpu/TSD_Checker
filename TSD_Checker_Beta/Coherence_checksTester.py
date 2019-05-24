import TSD_Checker_V3_4
import inspect
from ExcelEdit import TestReturn as result
from ErrorMessages import errorMessagesDict as error


#Coherence checks requirements

def Test_02043_18_04939_COH_2000(workBook, TSDApp):
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
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "mesures et commandes (Mesure Parametre et Test Actionneur) / Tests de cohérence".casefold():
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
            list_table = list()
            list_measure = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value == "N/A" or workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    dict = {}
                    dict["value"] = workSheet.Cells(index, refColIndex).Value
                    dict["localisation"] = workSheet.Cells(index, refColIndex)
                    list_table.append(dict)

            if TSDApp.WorkbookStats.hasMeasure == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
                workSheetRange = workSheet.UsedRange
                nrCols = workSheetRange.Columns.Count
                measureColIndex = 0
                var = 0
                for index1 in range(1, 15):
                    for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                        if str(workSheet.Cells(index1,index2).Value).casefold().strip() == "libellé (signification)".casefold():
                            measureColIndex = index2
                            measureRowIndex = index1
                            break
                    if measureColIndex != 0:
                        break

                measureCellRange = workSheet.Cells(measureRowIndex, measureColIndex).MergeArea
                nrLines = measureCellRange.Rows.Count
                nrCols = measureCellRange.Columns.Count
                localisation = list()

                for index in range(measureRowIndex + nrLines, TSDApp.WorkbookStats.measureLastRow + 1):
                    if workSheet.Cells(index, measureColIndex).Value == None:
                        pass
                    else:
                        list_measure.append(workSheet.Cells(index, measureColIndex).Value)

            for element in list_table:
                if element in list_measure:
                    pass
                else:
                    localisation.append(element["localisation"])

            if not localisation:
                localisation = None

            if localisation is None:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisation, workBook,TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

    return check

def is_ascii(s):
    return all(ord(c) < 128 for c in s)

def Test_02043_18_04939_COH_2001(workBook, TSDApp):
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
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "mesures et commandes (Mesure Parametre et Test Actionneur) / Tests de cohérence".casefold():
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
            contor = 0

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value is None:
                    pass
                else:
                    try:
                        cel = workSheet.Cells(index, refColIndex).Value.split("-")
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
                                contor = contor + 1
                        else:
                            localisation.append(workSheet.Cells(index, refColIndex))
                    except:
                        localisation.append(workSheet.Cells(index, refColIndex))
            if not localisation:
                localisation = None
            if contor == TSDApp.WorkbookStats.tableLastRow - refRowIndex - 1:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisation, workBook,TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook,TSDApp)
    return check

def Test_02043_18_04939_COH_2002(workBook, TSDApp, DOC8List):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)

        codeRowIndex = 0
        codeColIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Code défaut".casefold():
                    codeColIndex = index2
                    codeRowIndex = index1
                    break
            if codeColIndex != 0:
                break
        if codeColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(codeRowIndex, codeColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            localisation = []

            for index in range(codeRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow + 1):
                try:
                    cel = workSheet.Cells(index, codeColIndex).Value.split("-")
                    if cel[0] not in DOC8List:
                        localisation.append(workSheet.Cells(index, codeColIndex))
                except:
                    localisation.append(workSheet.Cells(index, codeColIndex))
            if not localisation:
                localisation = None
            if not localisation:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisation, workBook,TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook,TSDApp)
    return check

def Test_02043_18_04939_COH_2005(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        var = 0
        codeColIndex = 0
        codeRowIndex = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Code défaut".casefold():
                    codeColIndex = index2
                    codeRowIndex = index1
                    break
            if codeColIndex != 0:
                break
        if codeColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(codeRowIndex, codeColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            localisation = []

            for index in range(codeRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow + 1):
                try:
                    cel = workSheet.Cells(index, codeColIndex).Value.split("-")
                    if len(cel) == 2:
                        if cel[0].isascii() and cel[1][0].isalpha() and len(cel[1]) == 5:
                            try:
                                int(cel[1][1:], 16)
                            except:
                                localisation.append(workSheet.Cells(index,codeColIndex))
                    else:
                        if len(cel) == 3:
                            a = 3
                except:
                    localisation.append(workSheet.Cells(index, codeColIndex))
            if not localisation:
                localisation = None
            if not localisation:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisation, workBook,
                       TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook,
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
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        codeColIndex = 0
        codeRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Code défaut".casefold():
                    codeColIndex = index2
                    codeRowIndex = index1
                    break
            if codeColIndex != 0:
                break
        if codeColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(codeRowIndex, codeColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            localisation = []
            contor = 0

            for index in range(codeRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow + 1):
                try:
                    cel = workSheet.Cells(index, codeColIndex).Value.split("-")
                    if cel[0] not in DOC8List:
                        localisation.append(workSheet.Cells(index, codeColIndex))
                except:
                    localisation.append(workSheet.Cells(index, codeColIndex))
            if not localisation:
                localisation = None
            if not localisation:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisation, workBook,TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook,TSDApp)
    return check

def Test_02043_18_04939_COH_2007(ExcelApp, workBook, TSDApp, DOC14Name):
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
            for cell in cellRow.Cells:
                if str(cell.Value).casefold().strip() == "Data Trouble Code (DTC)".casefold():
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
            localisation = []
            flag = False
            list_ref  = []


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


            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        return check

def Test_02043_18_04939_COH_2010(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        list_code = list()
        tempList = list()
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

        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            nrCols = refCellRange.Columns.Count
            localisation = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
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
                for index1 in range(1, 15):
                    for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                        if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Code défaut".casefold():
                            codeColIndex = index2
                            codeRowIndex = index1
                            break
                    if codeColIndex != 0:
                        break

                codeCellRange = workSheet.Cells(codeRowIndex, codeColIndex).MergeArea
                nrLines = codeCellRange.Rows.Count
                nrCols = codeCellRange.Columns.Count
                localisation = list()


                for index in range(codeRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow + 1):
                    if workSheet.Cells(index, codeColIndex).Value == None:
                        pass
                    else:
                        list_code.append(workSheet.Cells(index, codeColIndex).Value.strip())



            for element in tempList:
                if ',' in element["value"]:
                    elem = element["value"].split(",")
                    for i in elem:
                        if i.strip() in list_code:
                            pass
                        else:
                            localisation.append(element["localisation"])
                            check = True
                else:
                    if element["value"].strip() in list_code or element["value"] in list_code:
                        pass
                    else:
                        localisation.append(element["localisation"])

            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2020(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        workSheetRange = workSheet.UsedRange
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

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
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
                for index1 in range(1, 15):
                    for index2 in range(1, TSDApp.WorkbookStats.constituantsLastCol + 1):
                        if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Noms".casefold():
                            constituantsColIndex = index2
                            constituantsRowIndex = index1
                            break
                    if constituantsColIndex != 0:
                        break

                constituantsCellRange = workSheet.Cells(constituantsRowIndex, constituantsColIndex).MergeArea
                nrLines = constituantsCellRange.Rows.Count
                localisation = list()

                for index in range(constituantsRowIndex + nrLines, TSDApp.WorkbookStats.constituantsLastRow + 1):
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

            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2030(workBook, TSDApp):
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
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Effet(s) client(s)".casefold():
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
            list_table = dict()
            list_eff = list()
            tempList = list()


            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
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
                for index1 in range(1, 15):
                    for index2 in range(1, TSDApp.WorkbookStats.EffClientsLastCol + 1):
                        if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Noms".casefold():
                            effColIndex = index2
                            effRowIndex = index1
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


                for index in range(effRowIndex + nrLines, nrRows + 1):
                    if workSheet.Cells(index, effColIndex).Value == None:
                        pass
                    else:
                        list_eff.append(workSheet.Cells(index, effColIndex).Value.strip())

            for element in tempList:
                if element["value"].strip() in list_eff:
                    pass
                else:
                    localisation.append(element["localisation"])
                    check = True

            if  not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2040(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        nrRows = workSheet.Rows.Count
        refColIndex = 0
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

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
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
                for index1 in range(1, 15):
                    for index2 in range(1, TSDApp.WorkbookStats.DiagDebLastCol + 1):
                        if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "libellé (signification)".casefold():
                            diagColIndex = index2
                            diagRowIndex = index1
                            break
                    if diagColIndex != 0:
                        break

                diagCellRange = workSheet.Cells(diagRowIndex, diagColIndex).MergeArea
                nrLines = diagCellRange.Rows.Count
                nrCols = diagCellRange.Columns.Count
                localisation = list()

                for index in range(diagRowIndex + nrLines, TSDApp.WorkbookStats.DiagDebLastRow + 1):
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

            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2050(workBook, TSDApp):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        nrRows = workSheet.Rows.Count
        refColIndex = 0
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

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
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
                for index1 in range(1, 15):
                    for index2 in range(1, TSDApp.WorkbookStats.ERLastCol + 1):
                        if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "nom".casefold():
                            ERColIndex = index2
                            ERRowIndex = index1
                            break
                    if ERColIndex != 0:
                        break

                ERCellRange = workSheet.Cells(ERRowIndex, ERColIndex).MergeArea
                nrLines = ERCellRange.Rows.Count
                localisation = list()
                for index in range(ERRowIndex + nrLines, TSDApp.WorkbookStats.ERLastRow + 1):
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

            if not localisation:
                localisation = None

            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
    return check

def Test_02043_18_04939_COH_2060(ExcelApp, workBook, TSDApp, DOC7Name):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.EffClientsIndex)
        nrRows = workSheet.Rows
        effColIndex = 0
        var = 0
        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.EffClientsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Noms".casefold():
                    effColIndex = index2
                    effRowIndex = index1
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

            for index in range(effRowIndex + nrLines, TSDApp.WorkbookStats.EffClientsLastRow + 1):
                if workSheet.Cells(index, effColIndex).Value == None:
                    pass
                else:
                    dict = {}
                    dict["value"] = workSheet.Cells(index, effColIndex).Value
                    dict["localisation"] = workSheet.Cells(index, effColIndex)
                    list_eff.append(dict)

            DOC7 = ExcelApp.Workbooks.Open(DOC7Name)
            if workBook.Sheets("Effets clients"):
                workSheetRef = DOC7.Sheets("FR")
            elif workBook.Sheets("Customer Effects") or workBook.Sheets("Customer Effect"):
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
            if N2RowIndex != 0:
                for index in range(N2RowIndex + nrLines, nrRows + 1):
                    if workSheetRef.Cells(index, N2ColIndex).Value == None:
                        pass
                    else:
                        list_ref.append(workSheetRef.Cells(index, N2ColIndex).Value)
            if N3RowIndex != 0:
                for index in range(N3RowIndex + nrLines, nrRows + 1):
                    if workSheetRef.Cells(index, N3ColIndex).Value == None:
                        pass
                    else:
                        list_ref.append(workSheetRef.Cells(index, N3ColIndex).Value)

            for element in list_eff:
                if element["value"] in list_ref:
                    pass
                else:
                    localisation.append(element["localisation"])
                    check = True

            if not localisation:
                localisation = None

        if localisation is not None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)
        else:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
    return check

def Test_02043_18_04939_COH_2070(ExcelApp, workBook, TSDApp, DOC7Name):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasEffClients == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.EffClientsIndex)
        nrRows = workSheet.Rows
        effColIndex = 0
        var = 0
        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.EffClientsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Name".casefold():
                    effColIndex = index2
                    effRowIndex = index1
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

            for index in range(effRowIndex + nrLines, TSDApp.WorkbookStats.EffClientsLastRow + 1):
                if workSheet.Cells(index, effColIndex).Value == None:
                    pass
                else:
                    dict = {}
                    dict["value"] = workSheet.Cells(index, effColIndex).Value
                    dict["localisation"] = workSheet.Cells(index, effColIndex)
                    list_eff.append(dict)

            DOC7 = ExcelApp.Workbooks.Open(DOC7Name)
            if workBook.Sheets("Effets clients"):
                workSheetRef = DOC7.Sheets("FR")
            elif workBook.Sheets("Customer Effects") or workBook.Sheets("Customer Effect"):
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
            if N2RowIndex != 0:
                for index in range(N2RowIndex + nrLines, nrRows + 1):
                    if workSheetRef.Cells(index, N2ColIndex).Value == None:
                        pass
                    else:
                        list_ref.append(workSheetRef.Cells(index, N2ColIndex).Value)
            if N3RowIndex != 0:
                for index in range(N3RowIndex + nrLines, nrRows + 1):
                    if workSheetRef.Cells(index, N3ColIndex).Value == None:
                        pass
                    else:
                        list_ref.append(workSheetRef.Cells(index, N3ColIndex).Value)

            for element in list_eff:
                if element["value"] in list_ref:
                    pass
                else:
                    localisation.append(element["localisation"])
                    check = True

            if not localisation:
                localisation = None

        if localisation is not None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook,
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
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.EffClientsIndex)
        nrRows = workSheet.Rows
        effColIndex = 0
        var = 0
        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.EffClientsLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Noms".casefold():
                    effColIndex = index2
                    effRowIndex = index1
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

            for index in range(effRowIndex + nrLines, TSDApp.WorkbookStats.EffClientsLastRow + 1):
                if workSheet.Cells(index, effColIndex).Value == None:
                    pass
                else:
                    dict = {}
                    dict["value"] = workSheet.Cells(index, effColIndex).Value
                    dict["localisation"] = workSheet.Cells(index, effColIndex)
                    list_eff.append(dict)

            DOC7 = ExcelApp.Workbooks.Open(DOC7Name)
            if workBook.Sheets("Effets clients"):
                workSheetRef = DOC7.Sheets("FR")
            elif workBook.Sheets("Customer Effects") or workBook.Sheets("Customer Effect"):
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
            if N2RowIndex != 0:
                for index in range(N2RowIndex + nrLines, nrRows + 1):
                    if workSheetRef.Cells(index, N2ColIndex).Value == None:
                        pass
                    else:
                        list_ref.append(workSheetRef.Cells(index, N2ColIndex).Value)
            if N3RowIndex != 0:
                for index in range(N3RowIndex + nrLines, nrRows + 1):
                    if workSheetRef.Cells(index, N3ColIndex).Value == None:
                        pass
                    else:
                        list_ref.append(workSheetRef.Cells(index, N3ColIndex).Value)

            for element in list_eff:
                if element["value"] in list_ref:
                    pass
                else:
                    localisation.append(element["localisation"])
                    check = True

            if not localisation:
                localisation = None

        if localisation is not None:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook,
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

def Test_02043_18_04939_COH_2100(workBook, TSDApp, DOC8List):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasCode == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)

        codeColIndex = 0
        codeRowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "supporté par constituant (s)".casefold():
                    codeColIndex = index2
                    codeRowIndex = index1
                    break
            if codeColIndex != 0:
                break
        if codeColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(codeRowIndex, codeColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            localisation = []
            for index in range(codeRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow + 1):
                cel = workSheet.Cells(index, codeColIndex).Value
                if cel not in DOC8List:
                    localisation.append(workSheet.Cells(index, codeColIndex))

            if not localisation:
                localisation = None
            if not localisation:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisation, workBook,TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook,TSDApp)
    return check

def Test_02043_18_04939_COH_2110(workBook, TSDApp, DOC8List):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasMeasure == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)

        ColIndex = 0
        RowIndex = 0
        var = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.measureLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "supporté par constituant (s)".casefold():
                    ColIndex = index2
                    RowIndex = index1
                    break
            if ColIndex != 0:
                break
        if ColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(RowIndex, ColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            localisation = []
            contor = 0

            for index in range(RowIndex + nrLines, TSDApp.WorkbookStats.measureLastRow + 1):
                cel = workSheet.Cells(index, ColIndex).Value
                if cel is None:
                    pass
                else:
                    if cel not in DOC8List:
                        localisation.append(workSheet.Cells(index, ColIndex))

            if not localisation:
                localisation = None
            if not localisation:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisation, workBook,TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook,TSDApp)
    return check

def Test_02043_18_04939_COH_2120(ExcelApp, workBook, TSDApp, DOC5Name):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasReqTech == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReqTechIndex)

        list_amont = list()
        tempDict = list()
        var = 0
        localisation = list()


        if TSDApp.WorkbookStats.ReqTechRefColIndex == 0:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        else:
            refCellRange = workSheet.Cells(TSDApp.WorkbookStats.ReqTechRefRowIndex, TSDApp.WorkbookStats.ReqTechRefColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()
            flag = False

            for index in range(TSDApp.WorkbookStats.ReqTechRefRowIndex + nrLines, TSDApp.WorkbookStats.ReqTechLastRow + 1):
                if workSheet.Cells(index, TSDApp.WorkbookStats.ReqTechRefColIndex).Value == None:
                    pass
                else:
                    tempDict.append(workSheet.Cells(index, TSDApp.WorkbookStats.ReqTechRefColIndex).Value)


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
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        var = 0
        localisation = list()

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Référence".casefold() or str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Reference".casefold():
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
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        var = 0
        localisation = list()

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "code défauts induits".casefold():
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
            localisation = list()
            list_table = list()
            list_effets = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
                list_table_dict = {}
                list_table_dict["value"] = workSheet.Cells(index, refColIndex).Value
                list_table_dict["localisation"] = workSheet.Cells(index, refColIndex)
                list_table.append(dict(list_table_dict))



            if TSDApp.WorkbookStats.hasCode == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
                effColIndex = 0

                for index1 in range(1, 15):
                    for index2 in range(1, TSDApp.WorkbookStats.codeLastCol + 1):
                        if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Code défaut".casefold():
                            effColIndex = index2
                            effRowIndex = index1
                            break
                    if effColIndex != 0:
                        break

                if effColIndex != 0:

                    effCellRange = workSheet.Cells(effRowIndex, effColIndex).MergeArea
                    nrLines = effCellRange.Rows.Count
                    nrCols = effCellRange.Columns.Count


                    for index in range(effRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow + 1):
                        if workSheet.Cells(index, effColIndex).Value == None:
                            pass
                        else:
                            list_effets.append(workSheet.Cells(index, effColIndex).Value)

                    for element in list_table:
                        if element["value"] in list_effets or element["value"] == "N/A":
                            pass
                        else:
                            localisation.append(element["localisation"])
                            check = True

                    if not localisation:
                        localisation = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

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
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        refColIndex = 0
        var = 0
        localisation = list()

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

        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()
            list_table = list()
            list_effets = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.codeLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    list_table_dict = {}
                    list_table_dict["value"] = workSheet.Cells(index, refColIndex).Value
                    list_table_dict["localisation"] = workSheet.Cells(index, refColIndex)
                    list_table.append(dict(list_table_dict))

            if TSDApp.WorkbookStats.hasConstituants == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.Sheets(TSDApp.WorkbookStats.constituantsIndex)
                effColIndex = 0

                for index1 in range(1, 15):
                    for index2 in range(1, TSDApp.WorkbookStats.constituantsLastCol + 1):
                        if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Noms".casefold():
                            effColIndex = index2
                            effRowIndex = index1
                            break
                    if effColIndex != 0:
                        break
                if effColIndex != 0:

                    effCellRange = workSheet.Cells(effRowIndex, effColIndex).MergeArea
                    nrLines = effCellRange.Rows.Count
                    nrCols = effCellRange.Columns.Count

                    for index in range(effRowIndex + nrLines, TSDApp.WorkbookStats.constituantsLastRow + 1):
                        if workSheet.Cells(index, effColIndex).Value == None:
                            pass
                        else:
                            list_effets.append(workSheet.Cells(index, effColIndex).Value)

                    for element in list_table:
                        if element["value"] in list_effets:
                            pass
                        else:
                            localisation.append(element["localisation"])
                            check = True

                    if not localisation:
                        localisation = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation,
                           workBook, TSDApp)

                elif effColIndex == 0:
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
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        refColIndex = 0
        var = 0
        localisation = list()

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

        if var == 1:
            result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
            check = True
        elif var == 0:
            refCellRange = workSheet.Cells(refRowIndex, refColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = list()
            list_table = list()
            list_effets = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.measureLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    list_table_dict = {}
                    list_table_dict["value"] = workSheet.Cells(index, refColIndex).Value
                    list_table_dict["localisation"] = workSheet.Cells(index, refColIndex)
                    list_table.append(dict(list_table_dict))

            if TSDApp.WorkbookStats.hasConstituants == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.Sheets(TSDApp.WorkbookStats.constituantsIndex)
                effColIndex = 0

                for index1 in range(1, 15):
                    for index2 in range(1, TSDApp.WorkbookStats.constituantsLastCol + 1):
                        if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Noms".casefold():
                            effColIndex = index2
                            effRowIndex = index1
                            break
                    if effColIndex != 0:
                        break

                if effColIndex != 0:

                    effCellRange = workSheet.Cells(effRowIndex, effColIndex).MergeArea
                    nrLines = effCellRange.Rows.Count
                    nrCols = effCellRange.Columns.Count

                    for index in range(effRowIndex + nrLines, TSDApp.WorkbookStats.constituantsLastRow + 1):
                        if workSheet.Cells(index, effColIndex).Value == None:
                            pass
                        else:
                            list_effets.append(workSheet.Cells(index, effColIndex).Value)

                    for element in list_table:
                        if element["value"] in list_effets:
                            pass
                        else:
                            localisation.append(element["localisation"])
                            check = True

                    if not localisation:
                        localisation = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation,
                           workBook, TSDApp)

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
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DataCodesIndex)
        workSheetRange = workSheet.UsedRange
        nrCols = workSheetRange.Columns.Count
        nrRows = workSheetRange.Rows.Count
        refColIndex = 0
        var = 0
        localisation = list()

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.DataCodesLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Stored by the ECU".casefold():
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
            localisation = list()
            list_table = list()
            list_effets = list()

            for index in range(refRowIndex + nrLines, nrRows + 1):
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

                for index1 in range(1, 15):
                    for index2 in range(1, TSDApp.WorkbookStats.constituantsLastCol + 1):
                        if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Name".casefold():
                            effColIndex = index2
                            effRowIndex = index1
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


                    for index in range(effRowIndex + nrLines, nrRows + 1):
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
    print(testName)
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

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.ReadDataIOLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Stored by the ECU".casefold():
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
            localisation = list()
            list_table = list()
            list_effets = list()

            for index in range(refRowIndex + nrLines, nrRows + 1):
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

                for index1 in range(1, 15):
                    for index2 in range(1, TSDApp.WorkbookStats.constituantsLastCol + 1):
                        if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Name".casefold():
                            effColIndex = index2
                            effRowIndex = index1
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


                    for index in range(effRowIndex + nrLines, nrRows + 1):
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
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        var = 0
        localisation = list()

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "situation de vie".casefold():
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
            localisation = list()
            list_table = list()
            list_effets = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
                if workSheet.Cells(index, refColIndex).Value == None:
                    pass
                else:
                    list_table_dict = {}
                    list_table_dict["value"] = workSheet.Cells(index, refColIndex).Value
                    list_table_dict["localisation"] = workSheet.Cells(index, refColIndex)
                    list_table.append(dict(list_table_dict))

            if TSDApp.WorkbookStats.hasSitDeVie == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.Sheets(TSDApp.WorkbookStats.SitDeVieIndex)
                effColIndex = 0

                for index1 in range(1, 15):
                    for index2 in range(1, TSDApp.WorkbookStats.SitDeVieLastCol + 1):
                        if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Situations de vie".casefold():
                            effColIndex = index2
                            effRowIndex = index1
                            break
                    if effColIndex != 0:
                        break

                if effColIndex != 0:

                    effCellRange = workSheet.Cells(effRowIndex, effColIndex).MergeArea
                    nrLines = effCellRange.Rows.Count
                    nrCols = effCellRange.Columns.Count

                    for index in range(effRowIndex + nrLines, TSDApp.WorkbookStats.SitDeVieLastRow + 1):
                        if workSheet.Cells(index, effColIndex).Value == None:
                            pass
                        else:
                            list_effets.append(workSheet.Cells(index, effColIndex).Value)

                    for element in list_table:
                        if element["value"] in list_effets or element["value"] == "N/A":
                            pass
                        else:
                            localisation.append(element["localisation"])
                            check = True

                    if not localisation:
                        localisation = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation,
                           workBook, TSDApp)

                elif effColIndex == 0:
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
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        var = 0
        localisation = list()

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Situation".casefold():
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
            localisation = list()
            list_table = list()
            list_effets = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
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

                for index1 in range(1, 15):
                    for index2 in range(1, TSDApp.WorkbookStats.SitDeVieLastCol + 1):
                        if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Description".casefold():
                            effColIndex = index2
                            effRowIndex = index1
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


                    for index in range(effRowIndex + nrLines, nrRows + 1):
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
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        var = 0
        localisation = list()

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Diagnostic debarque".casefold():
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
            localisation = list()
            list_table = list()
            list_effets = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
                list_table_dict = {}
                if workSheet.Cells(index, refColIndex).Value is None:
                    pass
                else:
                    list_table_dict["value"] = workSheet.Cells(index, refColIndex).Value
                    list_table_dict["localisation"] = workSheet.Cells(index, refColIndex)
                    list_table.append(dict(list_table_dict))

            if TSDApp.WorkbookStats.hasDiagDeb == False:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
                check = True
            else:
                workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagDebIndex)
                effColIndex = 0

                for index1 in range(1, 15):
                    for index2 in range(1, TSDApp.WorkbookStats.DiagDebLastCol + 1):
                        if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "libellé (signification)".casefold():
                            effColIndex = index2
                            effRowIndex = index1
                            break
                    if effColIndex != 0:
                        break

                if effColIndex != 0:

                    effCellRange = workSheet.Cells(effRowIndex, effColIndex).MergeArea
                    nrLines = effCellRange.Rows.Count
                    nrCols = effCellRange.Columns.Count

                    for index in range(effRowIndex + nrLines, TSDApp.WorkbookStats.DiagDebLastRow + 1):
                        if workSheet.Cells(index, effColIndex).Value == None:
                            pass
                        else:
                            list_effets.append(workSheet.Cells(index, effColIndex).Value)

                    for element in list_table:
                        if element["value"] in list_effets or element["value"] == "N/A":
                            pass
                        else:
                            localisation.append(element["localisation"])
                            check = True

                    if not localisation:
                        localisation = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation,
                           workBook, TSDApp)

                elif effColIndex == 0:
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
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        refColIndex = 0
        var = 0
        localisation = list()

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Non-embedded diagnosis".casefold():
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
            localisation = list()
            list_table = list()
            list_effets = list()

            for index in range(refRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
                list_table_dict = {}
                list_table_dict["value"] = workSheet.Cells(index, refColIndex).Value
                list_table_dict["localisation"] = workSheet.Cells(index, refColIndex)
                list_table.append(dict(list_table_dict))


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

                    for index in range(effRowIndex + nrLines, nrRows + 1):
                        if workSheet.Cells(index, effColIndex).Value == None:
                            pass
                        else:
                            list_effets.append(workSheet.Cells(index, effColIndex).Value)

                    for element in list_table:
                        if element["value"] in list_effets or element["value"] == "N/A":
                            pass
                        else:
                            localisation.append(element["localisation"])
                            check = True

                    if not localisation:
                        localisation = None

                    result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook, TSDApp)

                elif effColIndex == 0:
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
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)

        codeColIndex = 0
        codeRowIndex = 0
        var = 0
        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "mesures et commandes (Mesure Parametre et Test Actionneur) / Tests de cohérence".casefold():
                    codeColIndex = index2
                    codeRowIndex = index1
                    break
            if codeColIndex != 0:
                break
        if codeColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(codeRowIndex, codeColIndex).MergeArea
            nrLines = refCellRange.Rows.Count
            localisation = []

            for index in range(codeRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
                try:
                    cel = workSheet.Cells(index, codeColIndex).Value.split("-")
                    if cel[0] == subfamily_name and cel[1].lstrip('_') in DOC15List:
                        pass
                    else:
                        localisation.append(workSheet.Cells(index, codeColIndex))
                except:
                    localisation.append(workSheet.Cells(index, codeColIndex))

            if not localisation:
                localisation = None
            if not localisation:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisation, workBook,TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook,TSDApp)
    return check

def Test_02043_18_04939_COH_2240(workBook, TSDApp, DOC13List):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        var = 0
        codeColIndex = 0
        codeRowIndex = 0
        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Variant/\noption".casefold() or str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Variante/\noption".casefold():
                    codeColIndex = index2
                    codeRowIndex = index1
                    break
            if codeColIndex != 0:
                break
        if codeColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(codeRowIndex, codeColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            localisation = []
            contor = 0

            for index in range(codeRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
                list2 = ['AND', 'OR', "NOT", "N/A"]
                cel = []
                try:
                    cel = workSheet.Cells(index, codeColIndex).Value.split(" ")
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
                        localisation.append(workSheet.Cells(index,codeColIndex))

                except:
                    pass

            if not localisation:
                localisation = None

            if contor == TSDApp.WorkbookStats.tableLastRow - codeRowIndex - 1:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisation, workBook,TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook,TSDApp)
    return check

def Test_02043_18_04939_COH_2241(workBook, TSDApp, DOC13List):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)

        var = 0
        codeColIndex = 0
        codeRowIndex = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Diversity".casefold():
                    codeColIndex = index2
                    codeRowIndex = index1
                    break
            if codeColIndex != 0:
                break
        if codeColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(codeRowIndex, codeColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            localisation = []
            contor = 0

            for index in range(codeRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
                list2 = ['AND', 'OR', "NOT", "N/A",""]
                cel = []
                try:
                    cel = workSheet.Cells(index, codeColIndex).Value.split(" ")
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
                        localisation.append(workSheet.Cells(index,codeColIndex))
                except:
                    pass

            if not localisation:
                localisation = None


            if contor == TSDApp.WorkbookStats.tableLastRow - codeRowIndex - 1:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisation, workBook,TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook,TSDApp)
    return check

def Test_02043_18_04939_COH_2250(workBook, TSDApp, DOC13List):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)

        var = 0
        codeColIndex = 0
        codeRowIndex = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Variant/\noption".casefold() or str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Variante/\noption".casefold():
                    codeColIndex = index2
                    codeRowIndex = index1
                    break
            if codeColIndex != 0:
                break
        if codeColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(codeRowIndex, codeColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            localisation = []
            contor = 0

            for index in range(codeRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
                list2 = ['AND', 'OR', "NOT", "N/A"]
                cel = []
                try:
                    cel = workSheet.Cells(index, codeColIndex).Value.split(" ")
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
                        localisation.append(workSheet.Cells(index,codeColIndex))
                except:
                    pass

            if not localisation:
                localisation = None

            if contor == TSDApp.WorkbookStats.tableLastRow - codeRowIndex - 1:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisation, workBook,TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook,TSDApp)
    return check

def Test_02043_18_04939_COH_2251(workBook, TSDApp, DOC13List):
    testName = inspect.currentframe().f_code.co_name
    print(testName)
    check = False
    if TSDApp.WorkbookStats.hasTable == False:
        result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], "", workBook, TSDApp)
        check = True
    else:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)

        var = 0
        codeColIndex = 0
        codeRowIndex = 0

        for index1 in range(1, 15):
            for index2 in range(1, TSDApp.WorkbookStats.tableLastCol + 1):
                if str(workSheet.Cells(index1, index2).Value).casefold().strip() == "Diversity".casefold():
                    codeColIndex = index2
                    codeRowIndex = index1
                    break
            if codeColIndex != 0:
                break
        if codeColIndex == 0:
            var = 1

        if var == 0:
            refCellRange = workSheet.Cells(codeRowIndex, codeColIndex).MergeArea
            nrLines = refCellRange.Rows.Count

            localisation = []
            contor = 0

            for index in range(codeRowIndex + nrLines, TSDApp.WorkbookStats.tableLastRow + 1):
                list2 = ['AND', 'OR', "NOT", "N/A"]
                try:
                    cel = cel.replace(",", "").replace(";", "")
                    cel = workSheet.Cells(index, codeColIndex).Value.split(" ")
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
                        localisation.append(workSheet.Cells(index,codeColIndex))

                except:
                    pass

            if not localisation:
                localisation = None

            if contor == TSDApp.WorkbookStats.tableLastRow - codeRowIndex - 1:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error["None"], localisation, workBook,TSDApp)
                check = True
            else:
                result(TSDApp.DOC9Dict[testName][TSDApp.checkLevel], testName, error[testName], localisation, workBook,TSDApp)
    return check