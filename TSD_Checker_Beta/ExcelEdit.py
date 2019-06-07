import TSD_Checker_V4_0
import time
from PyQt5 import QtGui
import xlrd
import xlwt
from xlutils.copy import copy

from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem

def TestReturn(criticity, testName, message, localisation, workBook, TSDApp):

    return_dict = {}
    return_dict["criticity"] = criticity
    return_dict["testName"] = testName
    return_dict["message"] = message
    return_dict["localisation"] = localisation
    TSDApp.return_list.append(return_dict)

    if criticity.casefold() == "blocking":
        TSDApp.criticity_blocking += 1
    else:
        if criticity.casefold() == "warning":
            TSDApp.criticity_warning += 1
        else:
            TSDApp.criticity_information += 1

    tempString = str()
    if localisation is None or localisation == "":
        tempString = "OK"
    else:
        tempString = "NOK"

    textBoxText = TSDApp.tab1.textbox.toPlainText()
    textBoxText = textBoxText + "\n" + testName + " " + tempString
    TSDApp.tab1.textbox.setText(textBoxText)
    TSDApp.tab1.textbox.moveCursor(QtGui.QTextCursor.End)

    TSDApp.IncrementProgressBar()

def ExcelWrite(return_list, workBook):
    # testReportSheet = workBook.Sheets("Test report")
    # lastRow = testReportSheet.UsedRange.Rows.Count + 1
    #
    # for elem in return_list:
    #     testReportSheet.Cells(lastRow, 1).Value = elem["criticity"]
    #
    #     ColorCell(elem["criticity"], testReportSheet.Cells(lastRow, 1))
    #     tempString = str()
    #     nrRows = -1
    #
    #     testReportSheet.Cells(lastRow, 2).Value = elem["testName"]
    #     if elem["localisation"] is not None:
    #         localisation_len = len(elem["localisation"])
    #         nrRows = 0
    #         for index in range(1,localisation_len):
    #             testReportSheet.Cells(lastRow + index, 2).Value = elem["testName"]
    #             testReportSheet.Cells(lastRow + index, 2).Font.ColorIndex = 2
    #             testReportSheet.Cells(lastRow + index, 1).Value = elem["criticity"]
    #             testReportSheet.Cells(lastRow + index, 1).Font.ColorIndex = 2
    #             nrRows = index
    #
    #
    #     if elem["localisation"] is None or elem["localisation"] == "":
    #         testReportSheet.Cells(lastRow, 3).Value = "OK"
    #     else:
    #         testReportSheet.Cells(lastRow, 3).Value = elem["message"]
    #
    #     if elem["localisation"] is None or elem["localisation"] == "":
    #         testReportSheet.Cells(lastRow, 4).Value = elem["localisation"]
    #         if nrRows > 0:
    #             lastRow += nrRows
    #         else:
    #             lastRow += 1
    #     else:
    #         if isinstance(elem["localisation"][0],str):
    #             for index, element in enumerate(elem["localisation"]):
    #                 testReportSheet.Cells(lastRow + index, 4).Value = element
    #                 lastRow += 1
    #         else:
    #             for index, element in enumerate(elem["localisation"]):
    #                 testReportSheet.Cells(lastRow + index, 4).Formula = "=HYPERLINK(\"#\'" + element.Worksheet.Name + "\'!"+ element.Address + "\",\"" + element.Address +"\")"
    #             lastRow = lastRow + index + 1

    # rb = xlrd.open_workbook(workBook, on_demand=True)
    # rb_sheet = rb.sheet_by_name("Test report")
    # lastRow = rb_sheet.nrows + 1
    # wb = copy(rb)
    # rb.release_resources()
    # del rb
    # ws = wb.get_sheet("Test report")
    #
    # for elem in return_list:
    #     blocking_style = xlwt.easyxf('pattern: pattern solid, fore_colour red;')
    #     ws.write(lastRow, 1, elem["criticity"], blocking_style)
    #     lastRow += 1
    #
    # wb.save(workBook)
    # ########################################################################################################
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Test report',  cell_overwrite_ok=True)
    ws11 = wb.add_sheet('codes défauts', cell_overwrite_ok=True)
    lastRow = 0

    ws.write(lastRow, 0, 'Criticity')
    ws.write(lastRow, 1, 'Requirements')
    ws.write(lastRow, 2, 'Message')
    ws.write(lastRow, 3, 'Localisation')

    lastRow += 1
    blocking_style = xlwt.easyxf('pattern: pattern solid, fore_colour red;')
    warning_style = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')
    text_style = xlwt.easyxf('font: colour white, bold False;')

    for elem in return_list:

        if elem["criticity"].casefold().strip() == "blocking":
            ws.write(lastRow, 0, elem["criticity"], blocking_style)
        elif elem["criticity"].casefold().strip() == "warning":
            ws.write(lastRow, 0, elem["criticity"], warning_style)
        else:
            ws.write(lastRow, 0, elem["criticity"])

        ws.write(lastRow, 1, elem["testName"])

        if elem["localisation"] is None or elem["localisation"] == "":
            ws.write(lastRow, 2, "OK")
        else:
            ws.write(lastRow, 2, elem["message"])

        if elem["localisation"] is None or elem["localisation"] == "":
            ws.write(lastRow, 3, elem["localisation"])
            lastRow += 1

        if elem["localisation"] is not None and elem["localisation"] != "":
            if isinstance(elem["localisation"][0], str):
                for index, element in enumerate(elem["localisation"]):
                    ws.write(lastRow + index, 3, element)

                for index in range(1, len(elem["localisation"]) + 1):
                    ws.write(lastRow + index, 0, elem["criticity"], text_style)
                    ws.write(lastRow + index, 1, elem["testName"], text_style)

                lastRow += index
            else:
                for index, element in enumerate(elem["localisation"]):

                    link = "HYPERLINK(\"#\'"+ str(element.Worksheet.Name) + "\'!" + str(element.Address) + "\",\"" + str(element.Address) + "\")"
                    ws.write(lastRow + index, 3, xlwt.Formula(link))

                for index in range(1, len(elem["localisation"]) + 1):
                    ws.write(lastRow + index, 0, elem["criticity"], text_style)
                    ws.write(lastRow + index, 1, elem["testName"], text_style)

                lastRow = lastRow + index


    wb.save('C:\\Users\\msnecula\\Downloads\\output.xls')

    # ########################################################################################################


    # def TestReturn(criticity, testName, message, localisation, workBook, TSDApp):
#     testReportSheet = workBook.Sheets("Test report")
#     lastRow = testReportSheet.UsedRange.Rows.Count + 1
#
#     testReportSheet.Cells(lastRow, 1).Value = criticity
#
#
#     ColorCell(criticity, testReportSheet.Cells(lastRow, 1))
#     tempString = str()
#
#     testReportSheet.Cells(lastRow, 2).Value = testName
#     if localisation is not None:
#         localisation_len = len(localisation)
#         for index in range(1,localisation_len):
#             testReportSheet.Cells(lastRow + index, 2).Value = testName
#             testReportSheet.Cells(lastRow + index, 2).Font.ColorIndex = 2
#             testReportSheet.Cells(lastRow + index, 1).Value = criticity
#             testReportSheet.Cells(lastRow + index, 1).Font.ColorIndex = 2
#
#     if localisation is None or localisation == "":
#         testReportSheet.Cells(lastRow, 3).Value = "OK"
#         tempString = "OK"
#     else:
#         testReportSheet.Cells(lastRow, 3).Value = message
#         tempString = "NOK"
#
#     if localisation is None or localisation == "":
#         testReportSheet.Cells(lastRow, 4).Value = localisation
#     else:
#         for index, element in enumerate(localisation):
#             testReportSheet.Cells(lastRow + index, 4).Formula = "=HYPERLINK(\"#\'" + element.Worksheet.Name + "\'!"+ element.Address + "\",\"" + element.Address +"\")"
#
#     textBoxText = TSDApp.tab1.textbox.toPlainText()
#     textBoxText = textBoxText + "\n" + testName + " " + tempString
#     TSDApp.tab1.textbox.setText(textBoxText)
#     TSDApp.tab1.textbox.moveCursor(QtGui.QTextCursor.End)
#
#     TSDApp.IncrementProgressBar()



def TestReturnName(criticity, testName, message, name, workBook, TSDApp):
    return_dict = {}
    return_dict["criticity"] = criticity
    return_dict["testName"] = testName
    return_dict["message"] = message
    return_dict["localisation"] = name
    TSDApp.return_list.append(return_dict)

    if criticity.casefold() == "blocking":
        TSDApp.criticity_blocking += 1
    else:
        if criticity.casefold() == "warning":
            TSDApp.criticity_warning += 1
        else:
            TSDApp.criticity_information += 1

    tempString = str()
    if name is None or name == "":
        tempString = "OK"
    else:
        tempString = "NOK"

    textBoxText = TSDApp.tab1.textbox.toPlainText()
    textBoxText = textBoxText + "\n" + testName + " " + tempString
    TSDApp.tab1.textbox.setText(textBoxText)
    TSDApp.tab1.textbox.moveCursor(QtGui.QTextCursor.End)

    TSDApp.IncrementProgressBar()

    # testReportSheet = workBook.Sheets("Test report")
    # lastRow = testReportSheet.UsedRange.Rows.Count + 1
    #
    # testReportSheet.Cells(lastRow, 1).Value = criticity
    # if criticity.casefold() == "blocking":
    #     TSDApp.criticity_blocking += 1
    # else:
    #     if criticity.casefold() == "warning":
    #         TSDApp.criticity_warning += 1
    #     else:
    #         TSDApp.criticity_information += 1
    #
    # ColorCell(criticity, testReportSheet.Cells(lastRow, 1))
    # tempString = str()
    #
    # testReportSheet.Cells(lastRow, 2).Value = testName
    # if name is not None:
    #     name_len = len(name)
    #     for index in range(1,name_len):
    #         testReportSheet.Cells(lastRow + index, 2).Value = testName
    #         testReportSheet.Cells(lastRow + index, 2).Font.ColorIndex = 2
    #         testReportSheet.Cells(lastRow + index, 1).Value = criticity
    #         testReportSheet.Cells(lastRow + index, 1).Font.ColorIndex = 2
    #
    # if name is None or name == "":
    #     testReportSheet.Cells(lastRow, 3).Value = "OK"
    #     tempString = "OK"
    # else:
    #     testReportSheet.Cells(lastRow, 3).Value = message
    #     tempString = "NOK"
    #
    # if name is None or name == "":
    #     testReportSheet.Cells(lastRow, 4).Value = name
    # else:
    #     for index, element in enumerate(name):
    #         testReportSheet.Cells(lastRow + index, 4).Value = name
    #
    # textBoxText = TSDApp.tab1.textbox.toPlainText()
    # textBoxText = textBoxText + "\n" + testName + " " + tempString
    # TSDApp.tab1.textbox.setText(textBoxText)
    # TSDApp.tab1.textbox.moveCursor(QtGui.QTextCursor.End)
    #
    # TSDApp.IncrementProgressBar()

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
    testReportWorkSheet.Columns("A").ColumnWidth = 12
    testReportWorkSheet.Columns("B").ColumnWidth = 35
    testReportWorkSheet.Columns("C").ColumnWidth = 150
    testReportWorkSheet.Columns("D").ColumnWidth = 12


    #testReportWorkSheet.Range("A1:D1").Font.Bold = True

def WriteReportInformationSheet(workBook, TSDApp):
    reportInformationWorkSheet = workBook.Sheets("Report information")
    colList = list()
    colList.append(list(("Tool version:", TSD_Checker_V4_0.appName)))
    colList.append(list(("Criticity configuration file:", TSDApp.DOC9Path)))
    colList.append(list(("","")))
    colList.append(list(("Extract CESARE file:", TSDApp.DOC8Path)))
    colList.append(list(("Customer effects file:", TSDApp.DOC7Name)))
    colList.append(list(("Diversity management file:", TSDApp.DOC13Path)))
    colList.append(list(("CESARE file reference:", TSDApp.DOC8Link.split("/")[-3])))
    colList.append(list(("Criticity configuration file reference:", TSDApp.DOC9Link.split("/")[-3])))
    colList.append(list(("Customer effect file reference:", TSDApp.DOC7Link.split("/")[-3])))
    colList.append(list(("Diversity management file reference:", TSDApp.DOC13Link.split("/")[-3])))
    colList.append(list(("Check level:", TSDApp.checkLevel)))
    colList.append(list(("","")))
    colList.append(list(("Date of the test:", time.strftime("%x"))))
    colList.append(list(("Time of the test:", time.strftime("%X"))))
    colList.append(list(("Test duration:", time.strftime('%H:%M:%S', time.gmtime(TSDApp.end_time - TSDApp.start_time)))))
    colList.append(list(("Opening duration:", time.strftime('%H:%M:%S', time.gmtime(TSDApp.opening_time - TSDApp.start_time)) )))
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
    colList.append(list(("Blocking Points", str(TSDApp.criticity_blocking))))
    colList.append(list(("Warning Points", str(TSDApp.criticity_warning))))
    colList.append(list(("Information Points", str(TSDApp.criticity_information))))
    reportInformationWorkSheet.Range("A1:B26").Value = colList
    for column in reportInformationWorkSheet.Range("A1:B26").Columns:
        column.AutoFit()
