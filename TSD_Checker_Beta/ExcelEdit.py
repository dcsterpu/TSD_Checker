import TSD_Checker_V3_1
import time

def TestReturn(criticity, testName, message, localisation, workBook, TSDApp):
    testReportSheet = workBook.Sheets("Test report")
    lastRow = testReportSheet.UsedRange.Rows.Count + 1

    testReportSheet.Cells(lastRow, 1).Value = criticity
    ColorCell(criticity, testReportSheet.Cells(lastRow, 1))
    tempString = str()

    testReportSheet.Cells(lastRow, 2).Value = testName

    if localisation is None:
        testReportSheet.Cells(lastRow, 3).Value = "OK"
        tempString = "OK"
    else:
        testReportSheet.Cells(lastRow, 3).Value = message
        tempString = "NOK"

    if localisation is None or localisation == "":
        testReportSheet.Cells(lastRow, 4).Value = localisation
    else:
        for index, element in enumerate(localisation):
            testReportSheet.Cells(lastRow + index, 4).Formula = "=HYPERLINK(\"#\'" + element.Worksheet.Name + "\'!"+ element.Address + "\",\"" + element.Address +"\")"

    textBoxText = TSDApp.tab1.textbox.toPlainText()
    textBoxText = textBoxText + "\n" + testName + " " + tempString
    TSDApp.tab1.textbox.setText(textBoxText)

    TSDApp.IncrementProgressBar()

def ColorCell(criticity, cell):
    if criticity == "Blocking":
        cell.Interior.ColorIndex = 3
    elif criticity == "Warning":
        cell.Interior.ColorIndex = 12
    elif criticity == "Information":
        cell.Interior.ColorIndex = 6

def AddTestReportSheets(workBook):
    try:
        workSheet = workBook.Sheets("Report information")
        workSheet.Application.DisplayAlerts = False
        workSheet.Delete()
        workSheet.Application.DisplayAlerts = True
    except:
        pass
    try:
        workSheet = workBook.Sheets("Test report")
        workSheet.Application.DisplayAlerts = False
        workSheet.Delete()
        workSheet.Application.DisplayAlerts = True
    except:
        pass

    reportInfoWorkSheet = workBook.Sheets.Add(None,workBook.Sheets(workBook.Sheets.Count),1,None)
    reportInfoWorkSheet.Name = "Report information"
    testReportWorkSheet = workBook.Sheets.Add(None,workBook.Sheets(workBook.Sheets.Count),1,None)
    testReportWorkSheet.Name = "Test report"

def AddTestReportSheetHeader(workBook):
    testReportWorkSheet = workBook.Sheets("Test report")
    textList = ["Criticity", "Requirements", "Message", "Localisation"]
    testReportWorkSheet.Range("A1:D1").Value = textList
    for column in testReportWorkSheet.Range("A1:D145").Columns:
        column.AutoFit()

    #testReportWorkSheet.Range("A1:D1").Font.Bold = True

def WriteReportInformationSheet(workBook, TSDApp):
    reportInformationWorkSheet = workBook.Sheets("Report information")
    colList = list()
    colList.append(list(("Tool version:", TSD_Checker_V3_1.appName)))
    colList.append(list(("Criticity configuration file:", TSDApp.DOC9Name)))
    colList.append(list(("","")))
    colList.append(list(("Extract CESARE file:", TSDApp.DOC8Name)))
    colList.append(list(("Customer effects file:", TSDApp.DOC7Name)))
    colList.append(list(("Check level:", TSDApp.checkLevel)))
    colList.append(list(("","")))
    colList.append(list(("Date of the test:", time.strftime("%x"))))
    colList.append(list(("Time of the test:", time.strftime("%X"))))
    colList.append(list(("","")))
    colList.append(list(("TSD file checked:", TSDApp.DOC3Path)))
    colList.append(list(("TSD function file checked:", TSDApp.DOC4Path)))
    colList.append(list(("TSD system file checked:", TSDApp.DOC5Path)))
    colList.append(list(("","")))
    colList.append(list(("AMDEC:",TSDApp.AMDECName)))
    colList.append(list(("Export MedialecMatrice:",TSDApp.MedialecName)))
    colList.append(list(("","")))
    colList.append(list(("Status:", str(TSDApp.status))))
    colList.append(list(("Coverage Indicator:", str(TSDApp.coverage)[0:4] + "%")))
    colList.append(list(("Convergence Indicator:", str(TSDApp.convergence)[0:4] + "%")))
    colList.append(list(("", "")))
    colList.append(list(("Blocking Points", "")))
    colList.append(list(("Warning Points", "")))
    colList.append(list(("Information Points", "")))
    reportInformationWorkSheet.Range("A1:B24").Value = colList
    for column in reportInformationWorkSheet.Range("A1:B24").Columns:
        column.AutoFit()
