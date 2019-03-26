import TSD_Checker_V1_0

def DOC9Parser(ExcelApp, DOC9Path):

    DOC9 = ExcelApp.Workbooks.Open(DOC9Path)
    workSheet = DOC9.Sheets("Configuration")
    workSheetRangeValuesTuple = workSheet.UsedRange.Value
    fillDictFlag = False
    DOC9Dict = dict()

    for rowTuple in workSheetRangeValuesTuple:
        if fillDictFlag == True:
            tempDict = dict()
            tempDict["previsional"] = rowTuple[prevCol]
            tempDict["consolidated"] = rowTuple[consCol]
            tempDict["validated"] = rowTuple[valiCol]
            try:
                testName = "Test_" + rowTuple[testCol].strip()
                DOC9Dict[testName] = tempDict
            except:
                pass

        elif "Requirements" in rowTuple:
            testCol = rowTuple.index("Requirements")
            prevCol = rowTuple.index("Previsionnal")
            consCol = rowTuple.index("Consolidated")
            valiCol = rowTuple.index("Validated")
            fillDictFlag = True

    DOC9.Close()

    return DOC9Dict
