import TSD_Checker_V5_6
import time
from PyQt5 import QtGui
import xlrd
import xlwt
from xlutils.copy import copy
import openpyxl
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors

from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem
import os
import zipfile
import shutil
from shutil import copyfile
from shutil import rmtree



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


def deleteSheet(TSDApp, workbook, sheet_name1, sheet_name2):
    index_test_report = -1
    index_info_report = -1
    for sheetname in TSDApp.WorkbookStats.sheetNames:
        if sheetname == 'test report':
            index_test_report = TSDApp.WorkbookStats.sheetNames.index('test report')
        if sheetname == 'report information':
            index_info_report = TSDApp.WorkbookStats.sheetNames.index('report information')

    new_wb = copy(workbook)
    if index_info_report != -1:
        new_wb._Workbook__worksheets = [worksheet for worksheet in new_wb._Workbook__worksheets if worksheet.name.casefold() != sheet_name1]
    if index_test_report != -1:
        new_wb._Workbook__worksheets = [worksheet for worksheet in new_wb._Workbook__worksheets if worksheet.name.casefold() != sheet_name2]

    if index_info_report != -1:
        if new_wb._Workbook__worksheet_idx_from_name['report information'] > -1:
            del new_wb._Workbook__worksheet_idx_from_name['report information']
    if index_test_report != -1:
        if new_wb._Workbook__worksheet_idx_from_name['test report'] > -1:
            del new_wb._Workbook__worksheet_idx_from_name['test report']
    return new_wb


def ExcelWrite_del_information(return_list, workBook, TSDApp):

    DOC3 = TSDApp.DOC3Workbook

    new_wb = deleteSheet(TSDApp ,DOC3,"report information","test report")

    workSheet_info_report = new_wb.add_sheet('Report information', cell_overwrite_ok=True)

    workSheet_info_report.write(0, 0, "Tool version:")
    workSheet_info_report.write(0, 1, TSD_Checker_V5_6.appName)

    workSheet_info_report.write(2, 0, "Criticity configuration file:")
    workSheet_info_report.write(2, 1, TSDApp.DOC9Path)

    workSheet_info_report.write(3, 0, "Extract CESARE file:")
    workSheet_info_report.write(3, 1, TSDApp.DOC8Path)

    workSheet_info_report.write(4, 0, "Customer effects file:")
    workSheet_info_report.write(4, 1, TSDApp.DOC7Name)

    workSheet_info_report.write(5, 0, "Diversity management file:")
    workSheet_info_report.write(5, 1, TSDApp.DOC13Path)

    workSheet_info_report.write(6, 0, "CESARE file reference:")
    workSheet_info_report.write(6, 1, TSDApp.DOC8Link.split("/")[-3])

    workSheet_info_report.write(7, 0, "Criticity configuration file reference:")
    workSheet_info_report.write(7, 1, TSDApp.DOC9Link.split("/")[-3])

    workSheet_info_report.write(8, 0, "Customer effect file reference:")
    workSheet_info_report.write(8, 1, TSDApp.DOC7Link.split("/")[-3])

    workSheet_info_report.write(9, 0, "Diversity management file reference:")
    workSheet_info_report.write(9, 1, TSDApp.DOC13Link.split("/")[-3])

    workSheet_info_report.write(10, 0, "Check level:")
    workSheet_info_report.write(10, 1, TSDApp.checkLevel)

    workSheet_info_report.write(12, 0, "Date of the test:")
    workSheet_info_report.write(12, 1, time.strftime("%x"))

    workSheet_info_report.write(13, 0, "Time of the test:")
    workSheet_info_report.write(13, 1, time.strftime("%X"))

    workSheet_info_report.write(14, 0, "Test duration:")
    workSheet_info_report.write(14, 1, time.strftime('%H:%M:%S', time.gmtime(TSDApp.end_time - TSDApp.start_time)))

    workSheet_info_report.write(15, 0, "Opening duration:")
    workSheet_info_report.write(15, 1, time.strftime('%H:%M:%S', time.gmtime(TSDApp.opening_time - TSDApp.start_time)))

    workSheet_info_report.write(17, 0, "TSD file checked:")
    workSheet_info_report.write(17, 1, TSDApp.DOC3Path)

    workSheet_info_report.write(18, 0, "TSD function file checked:")
    workSheet_info_report.write(18, 1, TSDApp.DOC4Path)

    workSheet_info_report.write(19, 0, "TSD system file checked:")
    workSheet_info_report.write(19, 1, TSDApp.DOC5Path)

    workSheet_info_report.write(21, 0, "AMDEC:")
    workSheet_info_report.write(21, 1, TSDApp.AMDECName)

    workSheet_info_report.write(22, 0, "Export MedialecMatrice:")
    workSheet_info_report.write(22, 1, TSDApp.MedialecName)

    workSheet_info_report.write(24, 0, "Status:")
    workSheet_info_report.write(24, 1, str(TSDApp.status))

    workSheet_info_report.write(25, 0, "Coverage Indicator:")
    workSheet_info_report.write(25, 1, str(TSDApp.coverage)[0:4] + "%")

    workSheet_info_report.write(26, 0, "Convergence Indicator:")
    workSheet_info_report.write(26, 1, str(TSDApp.convergence)[0:4] + "%")

    workSheet_info_report.write(28, 0, "Blocking Points")
    workSheet_info_report.write(28, 1, str(TSDApp.criticity_blocking))

    workSheet_info_report.write(29, 0, "Warning Points")
    workSheet_info_report.write(29, 1, str(TSDApp.criticity_warning))

    workSheet_info_report.write(30, 0, "Information Points")
    workSheet_info_report.write(30, 1, str(TSDApp.criticity_information))


    list_alpha = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S',
                  'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ',
                  'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
                  'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP',
                  'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ', 'CA', 'CB', 'CC', 'CD', 'CE', 'CF',
                  'CG', 'CH', 'CI', 'CJ', 'CK', 'CL', 'CM', 'CN', 'CO', 'CP', 'CQ', 'CR', 'CS', 'CT', 'CU', 'CV',
                  'CW', 'CX', 'CY', 'CZ', 'DA', 'DB', 'DC', 'DD', 'DE', 'DF', 'DG', 'DH', 'DI', 'DJ', 'DK', 'DL',
                  'DM', 'DN', 'DO', 'DP', 'DQ', 'DR', 'DS', 'DT', 'DU', 'DV', 'DW', 'DX', 'DY', 'DZ']


    workSheet_test_report = new_wb.add_sheet('Test report', cell_overwrite_ok=True)

    lastRow = 0
    workSheet_test_report.write(lastRow, 0, 'Criticity')
    workSheet_test_report.write(lastRow, 1, 'Requirements')
    workSheet_test_report.write(lastRow, 2, 'Message')
    workSheet_test_report.write(lastRow, 3, 'Localisation')

    lastRow += 1
    blocking_style = xlwt.easyxf('pattern: pattern solid, fore_colour red;')
    warning_style = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')
    text_style = xlwt.easyxf('font: colour white, bold False;')

    for elem in return_list:

        if elem["criticity"].casefold().strip() == "blocking":
            workSheet_test_report.write(lastRow, 0, elem["criticity"], blocking_style)
        elif elem["criticity"].casefold().strip() == "warning":
            workSheet_test_report.write(lastRow, 0, elem["criticity"], warning_style)
        else:
            workSheet_test_report.write(lastRow, 0, elem["criticity"])

        workSheet_test_report.write(lastRow, 1, elem["testName"])

        if elem["localisation"] is None or elem["localisation"] == "":
            workSheet_test_report.write(lastRow, 2, "OK")
        else:
            workSheet_test_report.write(lastRow, 2, elem["message"])

        if elem["localisation"] is None or elem["localisation"] == "":
            workSheet_test_report.write(lastRow, 3, elem["localisation"])
            lastRow += 1

        try:
            if elem["localisation"] is not None and elem["localisation"] != "":
                if isinstance(elem["localisation"][0], str):
                    for index, element in enumerate(elem["localisation"]):
                        workSheet_test_report.write(lastRow + index, 3, element)

                    for index in range(1, len(elem["localisation"]) + 1):
                        workSheet_test_report.write(lastRow + index, 0, elem["criticity"], text_style)
                        workSheet_test_report.write(lastRow + index, 1, elem["testName"], text_style)

                    lastRow += index
                else:
                    for index, element in enumerate(elem["localisation"]):
                        link = "HYPERLINK(\"#\'" + str(element[0]) + "\'!$" + str(list_alpha[element[2]]) + "$" + str(element[1] + 1) + "\",\"$" + str(list_alpha[element[2]]) + "$" + str(element[1] + 1) + "\")"
                        workSheet_test_report.write(lastRow + index, 3, xlwt.Formula(link))

                    for index in range(1, len(elem["localisation"]) + 1):
                        workSheet_test_report.write(lastRow + index, 0, elem["criticity"], text_style)
                        workSheet_test_report.write(lastRow + index, 1, elem["testName"], text_style)

                    lastRow = lastRow + index
        except:
            TSDApp.tab1.textbox.setText("ERROR: when trying to write data in Test report")
            new_wb.save(TSDApp.DOC3Path)
            return

    new_wb.save(TSDApp.DOC3Path)








# def deleteSheet(workbook, sheet_name, xls_file):
#     new_wb = copy(workbook)
#     new_wb._Workbook__worksheets = [worksheet for worksheet in new_wb._Workbook__worksheets if worksheet.name != sheet_name]
#     new_wb.save(xls_file)
#     # return new_wb
#
#
# def ExcelWrite_del_information(return_list, workBook, TSDApp):
#
#     DOC3 = xlrd.open_workbook(workBook, formatting_info=True)
#     # DOC3 = TSDApp.DOC3Workbook
#     index_test_report = -1
#     index_info_report = -1
#     for sheetname in TSDApp.WorkbookStats.sheetNames:
#         if sheetname == 'test report':
#             index_test_report = TSDApp.WorkbookStats.sheetNames.index('test report')
#         if sheetname == 'report information':
#             index_info_report = TSDApp.WorkbookStats.sheetNames.index('report information')
#
#     xls_file = "C:\\Users\\msnecula\\Downloads\\Test TSDs\\02016_12_06260_TSD_Systeme_Perception_LKA_DAA_V9.1 (6).xls"
#     if index_info_report != -1:
#         deleteSheet(DOC3,"report information",xls_file)
#     print("a")

def ExcelWrite(return_list, workBook, TSDApp):
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


    # ########################################################################################################


    index_test_report = -1
    index_info_report = -1
    for sheetname in TSDApp.WorkbookStats.sheetNames:
        if sheetname == 'test report':
            index_test_report = TSDApp.WorkbookStats.sheetNames.index('test report')
        if sheetname == 'report information':
            index_info_report = TSDApp.WorkbookStats.sheetNames.index('report information')

    # DOC3 = xlrd.open_workbook(workBook, on_demand=True)
    DOC3 = TSDApp.DOC3Workbook
    list_alpha = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ','BA','BB','BC','BD','BE','BF','BG','BH','BI','BJ','BK','BL','BM','BN','BO','BP','BQ','BR','BS','BT','BU','BV','BW','BX','BY','BZ']

    if index_test_report != -1:
        workSheet_test_report = DOC3.sheet_by_index(index_test_report)
        nrCols_test_report = workSheet_test_report.ncols
        nrRows_test_report = workSheet_test_report.nrows

    if index_info_report != -1:
        workSheet_info_report = DOC3.sheet_by_index(index_info_report)
        nrCols_info_report = workSheet_info_report.ncols
        nrRows_info_report = workSheet_info_report.nrows

    workBook2 = copy(DOC3)

    if index_info_report != -1:
        workSheet_info_report = workBook2.get_sheet(index_info_report)
        # for index1 in range(0, nrRows_info_report):
        #     for index2 in range(0, nrCols_info_report):
        #         workSheet_info_report.write(index1, index2, "")

        workSheet_info_report._cell_overwrite_ok = True

        workSheet_info_report.write(0, 0, "Tool version:")
        workSheet_info_report.write(0, 1, TSD_Checker_V5_6.appName)

        workSheet_info_report.write(2, 0, "Criticity configuration file:")
        workSheet_info_report.write(2, 1, TSDApp.DOC9Path)

        workSheet_info_report.write(3, 0, "Extract CESARE file:")
        workSheet_info_report.write(3, 1, TSDApp.DOC8Path)

        workSheet_info_report.write(4, 0, "Customer effects file:")
        workSheet_info_report.write(4, 1, TSDApp.DOC7Name)

        workSheet_info_report.write(5, 0, "Diversity management file:")
        workSheet_info_report.write(5, 1, TSDApp.DOC13Path)

        workSheet_info_report.write(6, 0, "CESARE file reference:")
        workSheet_info_report.write(6, 1, TSDApp.DOC8Link.split("/")[-3])

        workSheet_info_report.write(7, 0, "Criticity configuration file reference:")
        workSheet_info_report.write(7, 1, TSDApp.DOC9Link.split("/")[-3])

        workSheet_info_report.write(8, 0, "Customer effect file reference:")
        workSheet_info_report.write(8, 1, TSDApp.DOC7Link.split("/")[-3])

        workSheet_info_report.write(9, 0, "Diversity management file reference:")
        workSheet_info_report.write(9, 1, TSDApp.DOC13Link.split("/")[-3])

        workSheet_info_report.write(10, 0, "Check level:")
        workSheet_info_report.write(10, 1, TSDApp.checkLevel)

        workSheet_info_report.write(12, 0, "Date of the test:")
        workSheet_info_report.write(12, 1, time.strftime("%x"))

        workSheet_info_report.write(13, 0, "Time of the test:")
        workSheet_info_report.write(13, 1, time.strftime("%X"))

        workSheet_info_report.write(14, 0, "Test duration:")
        workSheet_info_report.write(14, 1, time.strftime('%H:%M:%S', time.gmtime(TSDApp.end_time - TSDApp.start_time)))

        workSheet_info_report.write(15, 0, "Opening duration:")
        workSheet_info_report.write(15, 1,time.strftime('%H:%M:%S', time.gmtime(TSDApp.opening_time - TSDApp.start_time)))

        workSheet_info_report.write(17, 0, "TSD file checked:")
        workSheet_info_report.write(17, 1, TSDApp.DOC3Path)

        workSheet_info_report.write(18, 0, "TSD function file checked:")
        workSheet_info_report.write(18, 1, TSDApp.DOC4Path)

        workSheet_info_report.write(19, 0, "TSD system file checked:")
        workSheet_info_report.write(19, 1, TSDApp.DOC5Path)

        workSheet_info_report.write(21, 0, "AMDEC:")
        workSheet_info_report.write(21, 1, TSDApp.AMDECName)

        workSheet_info_report.write(22, 0, "Export MedialecMatrice:")
        workSheet_info_report.write(22, 1, TSDApp.MedialecName)

        workSheet_info_report.write(24, 0, "Status:")
        workSheet_info_report.write(24, 1, str(TSDApp.status))

        workSheet_info_report.write(25, 0, "Coverage Indicator:")
        workSheet_info_report.write(25, 1, str(TSDApp.coverage)[0:4] + "%")

        workSheet_info_report.write(26, 0, "Convergence Indicator:")
        workSheet_info_report.write(26, 1, str(TSDApp.convergence)[0:4] + "%")

        workSheet_info_report.write(28, 0, "Blocking Points")
        workSheet_info_report.write(28, 1, str(TSDApp.criticity_blocking))

        workSheet_info_report.write(29, 0, "Warning Points")
        workSheet_info_report.write(29, 1, str(TSDApp.criticity_warning))

        workSheet_info_report.write(30, 0, "Information Points")
        workSheet_info_report.write(30, 1, str(TSDApp.criticity_information))

        workSheet_info_report._cell_overwrite_ok = False
    else:
        workSheet_info_report = workBook2.add_sheet('Report information', cell_overwrite_ok=True)

        workSheet_info_report.write(0, 0, "Tool version:")
        workSheet_info_report.write(0, 1, TSD_Checker_V5_6.appName)

        workSheet_info_report.write(2, 0, "Criticity configuration file:")
        workSheet_info_report.write(2, 1, TSDApp.DOC9Path)

        workSheet_info_report.write(3, 0, "Extract CESARE file:")
        workSheet_info_report.write(3, 1, TSDApp.DOC8Path)

        workSheet_info_report.write(4, 0, "Customer effects file:")
        workSheet_info_report.write(4, 1, TSDApp.DOC7Name)

        workSheet_info_report.write(5, 0, "Diversity management file:")
        workSheet_info_report.write(5, 1, TSDApp.DOC13Path)

        workSheet_info_report.write(6, 0, "CESARE file reference:")
        workSheet_info_report.write(6, 1, TSDApp.DOC8Link.split("/")[-3])

        workSheet_info_report.write(7, 0, "Criticity configuration file reference:")
        workSheet_info_report.write(7, 1, TSDApp.DOC9Link.split("/")[-3])

        workSheet_info_report.write(8, 0, "Customer effect file reference:")
        workSheet_info_report.write(8, 1, TSDApp.DOC7Link.split("/")[-3])

        workSheet_info_report.write(9, 0, "Diversity management file reference:")
        workSheet_info_report.write(9, 1, TSDApp.DOC13Link.split("/")[-3])

        workSheet_info_report.write(10, 0, "Check level:")
        workSheet_info_report.write(10, 1, TSDApp.checkLevel)

        workSheet_info_report.write(12, 0, "Date of the test:")
        workSheet_info_report.write(12, 1, time.strftime("%x"))

        workSheet_info_report.write(13, 0, "Time of the test:")
        workSheet_info_report.write(13, 1, time.strftime("%X"))

        workSheet_info_report.write(14, 0, "Test duration:")
        workSheet_info_report.write(14, 1, time.strftime('%H:%M:%S', time.gmtime(TSDApp.end_time - TSDApp.start_time)))

        workSheet_info_report.write(15, 0, "Opening duration:")
        workSheet_info_report.write(15, 1, time.strftime('%H:%M:%S', time.gmtime(TSDApp.opening_time - TSDApp.start_time)))

        workSheet_info_report.write(17, 0, "TSD file checked:")
        workSheet_info_report.write(17, 1, TSDApp.DOC3Path)

        workSheet_info_report.write(18, 0, "TSD function file checked:")
        workSheet_info_report.write(18, 1, TSDApp.DOC4Path)

        workSheet_info_report.write(19, 0, "TSD system file checked:")
        workSheet_info_report.write(19, 1, TSDApp.DOC5Path)

        workSheet_info_report.write(21, 0, "AMDEC:")
        workSheet_info_report.write(21, 1, TSDApp.AMDECName)

        workSheet_info_report.write(22, 0, "Export MedialecMatrice:")
        workSheet_info_report.write(22, 1, TSDApp.MedialecName)

        workSheet_info_report.write(24, 0, "Status:")
        workSheet_info_report.write(24, 1, str(TSDApp.status))

        workSheet_info_report.write(25, 0, "Coverage Indicator:")
        workSheet_info_report.write(25, 1, str(TSDApp.coverage)[0:4] + "%")

        workSheet_info_report.write(26, 0, "Convergence Indicator:")
        workSheet_info_report.write(26, 1, str(TSDApp.convergence)[0:4] + "%")

        workSheet_info_report.write(28, 0, "Blocking Points")
        workSheet_info_report.write(28, 1, str(TSDApp.criticity_blocking))

        workSheet_info_report.write(29, 0, "Warning Points")
        workSheet_info_report.write(29, 1, str(TSDApp.criticity_warning))

        workSheet_info_report.write(30, 0, "Information Points")
        workSheet_info_report.write(30, 1, str(TSDApp.criticity_information))



    if index_test_report != -1:
        workSheet_test_report = workBook2.get_sheet(index_test_report)


        for index1 in range(0, nrRows_test_report):
            for index2 in range(0, nrCols_test_report):
                workSheet_test_report.write(index1, index2, "")

        workSheet_test_report._cell_overwrite_ok = True


        lastRow = 0
        workSheet_test_report.write(lastRow, 0, 'Criticity')
        workSheet_test_report.write(lastRow, 1, 'Requirements')
        workSheet_test_report.write(lastRow, 2, 'Message')
        workSheet_test_report.write(lastRow, 3, 'Localisation')

        lastRow += 1
        blocking_style = xlwt.easyxf('pattern: pattern solid, fore_colour red;')
        warning_style = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')
        text_style = xlwt.easyxf('font: colour white, bold False;')

        for elem in return_list:

            if elem["criticity"].casefold().strip() == "blocking":
                workSheet_test_report.write(lastRow, 0, elem["criticity"], blocking_style)
            elif elem["criticity"].casefold().strip() == "warning":
                workSheet_test_report.write(lastRow, 0, elem["criticity"], warning_style)
            else:
                workSheet_test_report.write(lastRow, 0, elem["criticity"])

            workSheet_test_report.write(lastRow, 1, elem["testName"])

            if elem["localisation"] is None or elem["localisation"] == "":
                workSheet_test_report.write(lastRow, 2, "OK")
            else:
                workSheet_test_report.write(lastRow, 2, elem["message"])

            if elem["localisation"] is None or elem["localisation"] == "":
                workSheet_test_report.write(lastRow, 3, elem["localisation"])
                lastRow += 1

            if elem["localisation"] is not None and elem["localisation"] != "":
                if isinstance(elem["localisation"][0], str):
                    for index, element in enumerate(elem["localisation"]):
                        workSheet_test_report.write(lastRow + index, 3, element)

                    for index in range(1, len(elem["localisation"]) + 1):
                        workSheet_test_report.write(lastRow + index, 0, elem["criticity"], text_style)
                        workSheet_test_report.write(lastRow + index, 1, elem["testName"], text_style)

                    lastRow += index
                else:
                    for index, element in enumerate(elem["localisation"]):
                        link = "HYPERLINK(\"#\'" + str(element[0]) + "\'!$" + str(list_alpha[element[2]]) + "$" + str(
                            element[1] + 1) + "\",\"$" + str(list_alpha[element[2]]) + "$" + str(element[1] + 1) + "\")"
                        workSheet_test_report.write(lastRow + index, 3, xlwt.Formula(link))

                    for index in range(1, len(elem["localisation"]) + 1):
                        workSheet_test_report.write(lastRow + index, 0, elem["criticity"], text_style)
                        workSheet_test_report.write(lastRow + index, 1, elem["testName"], text_style)

                    lastRow = lastRow + index

        workSheet_test_report._cell_overwrite_ok = False

    else:
        workSheet_test_report = workBook2.add_sheet('Test report', cell_overwrite_ok=True)

        lastRow = 0
        workSheet_test_report.write(lastRow, 0, 'Criticity')
        workSheet_test_report.write(lastRow, 1, 'Requirements')
        workSheet_test_report.write(lastRow, 2, 'Message')
        workSheet_test_report.write(lastRow, 3, 'Localisation')

        lastRow += 1
        blocking_style = xlwt.easyxf('pattern: pattern solid, fore_colour red;')
        warning_style = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')
        text_style = xlwt.easyxf('font: colour white, bold False;')

        for elem in return_list:

            if elem["criticity"].casefold().strip() == "blocking":
                workSheet_test_report.write(lastRow, 0, elem["criticity"], blocking_style)
            elif elem["criticity"].casefold().strip() == "warning":
                workSheet_test_report.write(lastRow, 0, elem["criticity"], warning_style)
            else:
                workSheet_test_report.write(lastRow, 0, elem["criticity"])

            workSheet_test_report.write(lastRow, 1, elem["testName"])

            if elem["localisation"] is None or elem["localisation"] == "":
                workSheet_test_report.write(lastRow, 2, "OK")
            else:
                workSheet_test_report.write(lastRow, 2, elem["message"])

            if elem["localisation"] is None or elem["localisation"] == "":
                workSheet_test_report.write(lastRow, 3, elem["localisation"])
                lastRow += 1

            if elem["localisation"] is not None and elem["localisation"] != "":
                if isinstance(elem["localisation"][0], str):
                    for index, element in enumerate(elem["localisation"]):
                        workSheet_test_report.write(lastRow + index, 3, element)

                    for index in range(1, len(elem["localisation"]) + 1):
                        workSheet_test_report.write(lastRow + index, 0, elem["criticity"], text_style)
                        workSheet_test_report.write(lastRow + index, 1, elem["testName"], text_style)

                    lastRow += index
                else:
                    for index, element in enumerate(elem["localisation"]):
                        link = "HYPERLINK(\"#\'" + str(element[0]) + "\'!$" + str(list_alpha[element[2]]) + "$" + str(
                            element[1] + 1) + "\",\"$" + str(list_alpha[element[2]]) + "$" + str(element[1] + 1) + "\")"
                        workSheet_test_report.write(lastRow + index, 3, xlwt.Formula(link))

                    for index in range(1, len(elem["localisation"]) + 1):
                        workSheet_test_report.write(lastRow + index, 0, elem["criticity"], text_style)
                        workSheet_test_report.write(lastRow + index, 1, elem["testName"], text_style)

                    lastRow = lastRow + index

            # if elem["localisation"] is not None and elem["localisation"] != "":
            #     if isinstance(elem["localisation"][0], str):
            #         for index, element in enumerate(elem["localisation"]):
            #             workSheet_test_report.write(lastRow + index, 3, element)
            #
            #         if len(elem['localisation']) > 1:
            #             for index in range(1, len(elem["localisation"])):
            #                 workSheet_test_report.write(lastRow + index, 0, elem["criticity"], text_style)
            #                 workSheet_test_report.write(lastRow + index, 1, elem["testName"], text_style)
            #
            #         lastRow += index + 1
            #     else:
            #         for index, element in enumerate(elem["localisation"]):
            #             link = "HYPERLINK(\"#\'"+ str(element[0]) + "\'!$" + str(list_alpha[element[2]]) + "$" + str(element[1]+1) +"\",\"$" + str(list_alpha[element[2]]) + "$" + str(element[1]+1) + "\")"
            #             workSheet_test_report.write(lastRow + index, 3, xlwt.Formula(link))
            #
            #         if len(elem['localisation']) > 1:
            #             for index in range(1, len(elem["localisation"])):
            #                 workSheet_test_report.write(lastRow + index, 0, elem["criticity"], text_style)
            #                 workSheet_test_report.write(lastRow + index, 1, elem["testName"], text_style)
            #
            #         lastRow += index + 1

    workBook2.save(TSDApp.DOC3Path)


    # ws = wb.add_sheet('Test report',  cell_overwrite_ok=True)
    # ws11 = wb.add_sheet('codes défauts', cell_overwrite_ok=True)
    # lastRow = 0
    #
    # ws.write(lastRow, 0, 'Criticity')
    # ws.write(lastRow, 1, 'Requirements')
    # ws.write(lastRow, 2, 'Message')
    # ws.write(lastRow, 3, 'Localisation')
    #
    # lastRow += 1
    # blocking_style = xlwt.easyxf('pattern: pattern solid, fore_colour red;')
    # warning_style = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')
    # text_style = xlwt.easyxf('font: colour white, bold False;')
    #
    # for elem in return_list:
    #
    #     if elem["criticity"].casefold().strip() == "blocking":
    #         ws.write(lastRow, 0, elem["criticity"], blocking_style)
    #     elif elem["criticity"].casefold().strip() == "warning":
    #         ws.write(lastRow, 0, elem["criticity"], warning_style)
    #     else:
    #         ws.write(lastRow, 0, elem["criticity"])
    #
    #     ws.write(lastRow, 1, elem["testName"])
    #
    #     if elem["localisation"] is None or elem["localisation"] == "":
    #         ws.write(lastRow, 2, "OK")
    #     else:
    #         ws.write(lastRow, 2, elem["message"])
    #
    #     if elem["localisation"] is None or elem["localisation"] == "":
    #         ws.write(lastRow, 3, elem["localisation"])
    #         lastRow += 1
    #
    #     if elem["localisation"] is not None and elem["localisation"] != "":
    #         if isinstance(elem["localisation"][0], str):
    #             for index, element in enumerate(elem["localisation"]):
    #                 ws.write(lastRow + index, 3, element)
    #
    #             for index in range(1, len(elem["localisation"]) + 1):
    #                 ws.write(lastRow + index, 0, elem["criticity"], text_style)
    #                 ws.write(lastRow + index, 1, elem["testName"], text_style)
    #
    #             lastRow += index
    #         else:
    #             for index, element in enumerate(elem["localisation"]):
    #
    #                 link = "HYPERLINK(\"#\'"+ str(element.Worksheet.Name) + "\'!" + str(element.Address) + "\",\"" + str(element.Address) + "\")"
    #                 ws.write(lastRow + index, 3, xlwt.Formula(link))
    #
    #             for index in range(1, len(elem["localisation"]) + 1):
    #                 ws.write(lastRow + index, 0, elem["criticity"], text_style)
    #                 ws.write(lastRow + index, 1, elem["testName"], text_style)
    #
    #             lastRow = lastRow + index
    #
    #
    # wb.save('C:\\Users\\msnecula\\Downloads\\output.xls')

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



def ExcelWrite2(return_list, workBook, TSDApp):

    if TSDApp.DOC3Path.split('.')[-1] == 'xlsm':
        try:
            wb = openpyxl.load_workbook(TSDApp.DOC3Path, keep_vba=True)
        except:
            return
    else:
        wb = openpyxl.load_workbook(TSDApp.DOC3Path, keep_vba=False)

    index_test_report = -1
    index_info_report = -1
    for sheetname in TSDApp.WorkbookStats.sheetNames:
        if sheetname == 'test report':
            index_test_report = TSDApp.WorkbookStats.sheetNames.index('test report')
        if sheetname == 'report information':
            index_info_report = TSDApp.WorkbookStats.sheetNames.index('report information')

    if index_info_report == -1:
        workSheet_info_report = wb.create_sheet("Report information")

        workSheet_info_report['A1'] = "Tool version:"
        workSheet_info_report['B1'] = TSD_Checker_V5_6.appName

        workSheet_info_report['A3'] = "Criticity configuration file:"
        workSheet_info_report['B3'] = TSDApp.DOC9Path

        workSheet_info_report['A4'] = "Extract CESARE file:"
        workSheet_info_report['B4'] = TSDApp.DOC8Path

        workSheet_info_report['A5'] = "Customer effects file:"
        workSheet_info_report['B5'] = TSDApp.DOC7Name

        workSheet_info_report['A6'] = "Diversity management file:"
        workSheet_info_report['B6'] = TSDApp.DOC13Path

        workSheet_info_report['A7'] = "CESARE file reference:"
        workSheet_info_report['B7'] = TSDApp.DOC8Link.split("/")[-3]

        workSheet_info_report['A8'] = "Criticity configuration file reference:"
        workSheet_info_report['B8'] = TSDApp.DOC9Link.split("/")[-3]

        workSheet_info_report['A9'] = "Customer effect file reference:"
        workSheet_info_report['B9'] = TSDApp.DOC7Link.split("/")[-3]

        workSheet_info_report['A10'] = "Diversity management file reference:"
        workSheet_info_report['B10'] = TSDApp.DOC13Link.split("/")[-3]

        workSheet_info_report['A11'] = "Check level:"
        workSheet_info_report['B11'] = TSDApp.checkLevel

        workSheet_info_report['A13'] = "Date of the test:"
        workSheet_info_report['B13'] = time.strftime("%x")

        workSheet_info_report['A14'] = "Time of the test:"
        workSheet_info_report['B14'] = time.strftime("%X")

        workSheet_info_report['A15'] = "Test duration:"
        workSheet_info_report['B15'] = time.strftime('%H:%M:%S', time.gmtime(TSDApp.end_time - TSDApp.start_time))

        workSheet_info_report['A16'] = "Opening duration:"
        workSheet_info_report['B16'] = time.strftime('%H:%M:%S', time.gmtime(TSDApp.opening_time - TSDApp.start_time))

        workSheet_info_report['A18'] = "TSD file checked:"
        workSheet_info_report['B18'] = TSDApp.DOC3Path

        workSheet_info_report['A19'] = "TSD function file checked:"
        workSheet_info_report['b19'] = TSDApp.DOC4Path

        workSheet_info_report['A20'] = "TSD system file checked:"
        workSheet_info_report['B20'] = TSDApp.DOC5Path

        workSheet_info_report['A22'] = "AMDEC:"
        workSheet_info_report['B22'] = TSDApp.AMDECName

        workSheet_info_report['A23'] = "Export MedialecMatrice:"
        workSheet_info_report['B23'] = TSDApp.MedialecName

        workSheet_info_report['A25'] = "Status:"
        workSheet_info_report['B25'] = str(TSDApp.status)

        workSheet_info_report['A26'] = "Coverage Indicator:"
        workSheet_info_report['B26'] = str(TSDApp.coverage)[0:4] + "%"

        workSheet_info_report['A27'] = "Convergence Indicator:"
        workSheet_info_report['B27'] = str(TSDApp.convergence)[0:4] + "%"

        workSheet_info_report['A29'] = "Blocking Points"
        workSheet_info_report['B29'] = str(TSDApp.criticity_blocking)

        workSheet_info_report['A30'] = "Warning Points"
        workSheet_info_report['B30'] = str(TSDApp.criticity_warning)

        workSheet_info_report['A31'] = "Information Points"
        workSheet_info_report['B31'] = str(TSDApp.criticity_information)

    else:
        workSheet_info_report = wb.get_sheet_by_name("Report information")
        wb.remove_sheet(workSheet_info_report)
        workSheet_info_report = wb.create_sheet("Report information")

        workSheet_info_report['A1'] = "Tool version:"
        workSheet_info_report['B1'] = TSD_Checker_V5_6.appName

        workSheet_info_report['A3'] = "Criticity configuration file:"
        workSheet_info_report['B3'] = TSDApp.DOC9Path

        workSheet_info_report['A4'] = "Extract CESARE file:"
        workSheet_info_report['B4'] = TSDApp.DOC8Path

        workSheet_info_report['A5'] = "Customer effects file:"
        workSheet_info_report['B5'] = TSDApp.DOC7Name

        workSheet_info_report['A6'] = "Diversity management file:"
        workSheet_info_report['B6'] = TSDApp.DOC13Path

        workSheet_info_report['A7'] = "CESARE file reference:"
        workSheet_info_report['B7'] = TSDApp.DOC8Link.split("/")[-3]

        workSheet_info_report['A8'] = "Criticity configuration file reference:"
        workSheet_info_report['B8'] = TSDApp.DOC9Link.split("/")[-3]

        workSheet_info_report['A9'] = "Customer effect file reference:"
        workSheet_info_report['B9'] = TSDApp.DOC7Link.split("/")[-3]

        workSheet_info_report['A10'] = "Diversity management file reference:"
        workSheet_info_report['B10'] = TSDApp.DOC13Link.split("/")[-3]

        workSheet_info_report['A11'] = "Check level:"
        workSheet_info_report['B11'] = TSDApp.checkLevel

        workSheet_info_report['A13'] = "Date of the test:"
        workSheet_info_report['B13'] = time.strftime("%x")

        workSheet_info_report['A14'] = "Time of the test:"
        workSheet_info_report['B14'] = time.strftime("%X")

        workSheet_info_report['A15'] = "Test duration:"
        workSheet_info_report['B15'] = time.strftime('%H:%M:%S', time.gmtime(TSDApp.end_time - TSDApp.start_time))

        workSheet_info_report['A16'] = "Opening duration:"
        workSheet_info_report['B16'] = time.strftime('%H:%M:%S', time.gmtime(TSDApp.opening_time - TSDApp.start_time))

        workSheet_info_report['A18'] = "TSD file checked:"
        workSheet_info_report['B18'] = TSDApp.DOC3Path

        workSheet_info_report['A19'] = "TSD function file checked:"
        workSheet_info_report['b19'] = TSDApp.DOC4Path

        workSheet_info_report['A20'] = "TSD system file checked:"
        workSheet_info_report['B20'] = TSDApp.DOC5Path

        workSheet_info_report['A22'] = "AMDEC:"
        workSheet_info_report['B22'] = TSDApp.AMDECName

        workSheet_info_report['A23'] = "Export MedialecMatrice:"
        workSheet_info_report['B23'] = TSDApp.MedialecName

        workSheet_info_report['A25'] = "Status:"
        workSheet_info_report['B25'] = str(TSDApp.status)

        workSheet_info_report['A26'] = "Coverage Indicator:"
        workSheet_info_report['B26'] = str(TSDApp.coverage)[0:4] + "%"

        workSheet_info_report['A27'] = "Convergence Indicator:"
        workSheet_info_report['B27'] = str(TSDApp.convergence)[0:4] + "%"

        workSheet_info_report['A29'] = "Blocking Points"
        workSheet_info_report['B29'] = str(TSDApp.criticity_blocking)

        workSheet_info_report['A30'] = "Warning Points"
        workSheet_info_report['B30'] = str(TSDApp.criticity_warning)

        workSheet_info_report['A31'] = "Information Points"
        workSheet_info_report['B31'] = str(TSDApp.criticity_information)

    workSheet_info_report.column_dimensions['A'].width = 40
    workSheet_info_report.column_dimensions['B'].width = 140

    list_alpha = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ','BA','BB','BC','BD','BE','BF','BG','BH','BI','BJ','BK','BL','BM','BN','BO','BP','BQ','BR','BS','BT','BU','BV','BW','BX','BY','BZ']
    if index_test_report == -1:
        workSheet_test_report = wb.create_sheet("Test report")

        lastRow = 1
        workSheet_test_report['A1'] = "Criticity"
        workSheet_test_report['B1'] = "Requirements"
        workSheet_test_report['C1'] = "Message"
        workSheet_test_report['D1'] = "Localisation"

        lastRow += 1

        my_red = openpyxl.styles.colors.Color(rgb='00FF0000')
        blocking_style = openpyxl.styles.fills.PatternFill(patternType='solid', fgColor=my_red)
        my_yellow = openpyxl.styles.colors.Color(rgb='00FFFF00')
        warning_style = openpyxl.styles.fills.PatternFill(patternType='solid', fgColor=my_yellow)
        text_style = Font(color='FFFFFFFF')


        for elem in return_list:

            if elem["criticity"].casefold().strip() == "blocking":
                workSheet_test_report.cell(lastRow, 1, elem["criticity"]).fill = blocking_style
            elif elem["criticity"].casefold().strip() == "warning":
                workSheet_test_report.cell(lastRow, 1, elem["criticity"]).fill = warning_style
            else:
                workSheet_test_report.cell(lastRow, 1, elem["criticity"])

            workSheet_test_report.cell(lastRow, 2, elem["testName"])

            if elem["localisation"] is None or elem["localisation"] == "":
                workSheet_test_report.cell(lastRow, 3, "OK")
            else:
                workSheet_test_report.cell(lastRow, 3, elem["message"])

            if elem["localisation"] is None or elem["localisation"] == "":
                workSheet_test_report.cell(lastRow, 4, elem["localisation"])
                lastRow += 1

            if elem["localisation"] is not None and elem["localisation"] != "":
                if isinstance(elem["localisation"][0], str):
                    for index, element in enumerate(elem["localisation"]):
                        workSheet_test_report.cell(lastRow + index, 4, element)

                    if len(elem['localisation']) > 1:
                        for index in range(1, len(elem["localisation"])):
                            workSheet_test_report.cell(lastRow + index, 1, elem["criticity"]).font = text_style
                            workSheet_test_report.cell(lastRow + index, 2, elem["testName"]).font = text_style

                    lastRow += index + 1
                else:
                    for index, element in enumerate(elem["localisation"]):
                        workSheet_test_report.cell(lastRow + index, 4).value = '$' + str(list_alpha[element[2]]) + '$' + str(element[1] + 1)
                        # workSheet_test_report.cell(lastRow + index, 4).hyperlink = '#%s %s %s!%s' % ("'" + str(element[0]).split(' ')[-3],str(element[0]).split(' ')[-2],str(element[0]).split(' ')[-1] + "'", str(list_alpha[element[2]])+str(element[1] + 1) )
                        workSheet_test_report.cell(lastRow + index, 4).hyperlink = '#%s!%s' % ("'" + str(element[0]) + "'", str(list_alpha[element[2]])+str(element[1] + 1) )


                    if len(elem['localisation']) > 1:
                        for index in range(1, len(elem["localisation"])):
                            workSheet_test_report.cell(lastRow + index, 1, elem["criticity"]).font = text_style
                            workSheet_test_report.cell(lastRow + index, 2, elem["testName"]).font = text_style

                    lastRow += index + 1
    else:
        workSheet_test_report = wb.get_sheet_by_name("Test report")
        wb.remove_sheet(workSheet_test_report)
        workSheet_test_report = wb.create_sheet("Test report")

        lastRow = 1
        workSheet_test_report['A1'] = "Criticity"
        workSheet_test_report['B1'] = "Requirements"
        workSheet_test_report['C1'] = "Message"
        workSheet_test_report['D1'] = "Localisation"

        lastRow += 1

        my_red = openpyxl.styles.colors.Color(rgb='00FF0000')
        blocking_style = openpyxl.styles.fills.PatternFill(patternType='solid', fgColor=my_red)
        my_yellow = openpyxl.styles.colors.Color(rgb='00FFFF00')
        warning_style = openpyxl.styles.fills.PatternFill(patternType='solid', fgColor=my_yellow)
        text_style = Font(color='FFFFFFFF')

        for elem in return_list:

            if elem["criticity"].casefold().strip() == "blocking":
                workSheet_test_report.cell(lastRow, 1, elem["criticity"]).fill = blocking_style
            elif elem["criticity"].casefold().strip() == "warning":
                workSheet_test_report.cell(lastRow, 1, elem["criticity"]).fill = warning_style
            else:
                workSheet_test_report.cell(lastRow, 1, elem["criticity"])

            workSheet_test_report.cell(lastRow, 2, elem["testName"])

            if elem["localisation"] is None or elem["localisation"] == "":
                workSheet_test_report.cell(lastRow, 3, "OK")
            else:
                workSheet_test_report.cell(lastRow, 3, elem["message"])

            if elem["localisation"] is None or elem["localisation"] == "":
                workSheet_test_report.cell(lastRow, 4, elem["localisation"])
                lastRow += 1

            if elem["localisation"] is not None and elem["localisation"] != "":
                if isinstance(elem["localisation"][0], str):
                    for index, element in enumerate(elem["localisation"]):
                        workSheet_test_report.cell(lastRow + index, 4, element)

                    if len(elem['localisation']) > 1:
                        for index in range(1, len(elem["localisation"])):
                            workSheet_test_report.cell(lastRow + index, 1, elem["criticity"]).font = text_style
                            workSheet_test_report.cell(lastRow + index, 2, elem["testName"]).font = text_style

                    lastRow += index + 1
                else:
                    for index, element in enumerate(elem["localisation"]):
                        workSheet_test_report.cell(lastRow + index, 4).value = '$' + str(list_alpha[element[2]]) + '$' + str(element[1] + 1)
                        # workSheet_test_report.cell(lastRow + index, 4).hyperlink = '#%s %s %s!%s' % ("'" + str(element[0]).split(' ')[-3],str(element[0]).split(' ')[-2],str(element[0]).split(' ')[-1] + "'", str(list_alpha[element[2]])+str(element[1] + 1) )
                        workSheet_test_report.cell(lastRow + index, 4).hyperlink = '#%s!%s' % (
                        "'" + str(element[0]) + "'", str(list_alpha[element[2]]) + str(element[1] + 1))

                    if len(elem['localisation']) > 1:
                        for index in range(1, len(elem["localisation"])):
                            workSheet_test_report.cell(lastRow + index, 1, elem["criticity"]).font = text_style
                            workSheet_test_report.cell(lastRow + index, 2, elem["testName"]).font = text_style

                    lastRow += index + 1

    workSheet_test_report.column_dimensions['A'].width = 20
    workSheet_test_report.column_dimensions['B'].width = 40
    workSheet_test_report.column_dimensions['C'].width = 80
    workSheet_test_report.column_dimensions['D'].width = 20

    wb.save(workBook)

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
    colList.append(list(("Tool version:", TSD_Checker_V5_6.appName)))
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


# import TSD_Checker_V5_6
# import time
# from PyQt5 import QtGui
# import xlrd
# import xlwt
# from xlutils.copy import copy
# import openpyxl
# from openpyxl.styles import Color, Font
# from openpyxl.styles import colors
#
# from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem
#
# def TestReturn(criticity, testName, message, localisation, workBook, TSDApp):
#
#     return_dict = {}
#     return_dict["criticity"] = criticity
#     return_dict["testName"] = testName
#     return_dict["message"] = message
#     return_dict["localisation"] = localisation
#     TSDApp.return_list.append(return_dict)
#
#     if criticity.casefold() == "blocking":
#         TSDApp.criticity_blocking += 1
#     else:
#         if criticity.casefold() == "warning":
#             TSDApp.criticity_warning += 1
#         else:
#             TSDApp.criticity_information += 1
#
#     tempString = str()
#     if localisation is None or localisation == "":
#         tempString = "OK"
#     else:
#         tempString = "NOK"
#
#     textBoxText = TSDApp.tab1.textbox.toPlainText()
#     textBoxText = textBoxText + "\n" + testName + " " + tempString
#     TSDApp.tab1.textbox.setText(textBoxText)
#     TSDApp.tab1.textbox.moveCursor(QtGui.QTextCursor.End)
#
#     TSDApp.IncrementProgressBar()
#
# def ExcelWrite(return_list, workBook, TSDApp):
#     # testReportSheet = workBook.Sheets("Test report")
#     # lastRow = testReportSheet.UsedRange.Rows.Count + 1
#     #
#     # for elem in return_list:
#     #     testReportSheet.Cells(lastRow, 1).Value = elem["criticity"]
#     #
#     #     ColorCell(elem["criticity"], testReportSheet.Cells(lastRow, 1))
#     #     tempString = str()
#     #     nrRows = -1
#     #
#     #     testReportSheet.Cells(lastRow, 2).Value = elem["testName"]
#     #     if elem["localisation"] is not None:
#     #         localisation_len = len(elem["localisation"])
#     #         nrRows = 0
#     #         for index in range(1,localisation_len):
#     #             testReportSheet.Cells(lastRow + index, 2).Value = elem["testName"]
#     #             testReportSheet.Cells(lastRow + index, 2).Font.ColorIndex = 2
#     #             testReportSheet.Cells(lastRow + index, 1).Value = elem["criticity"]
#     #             testReportSheet.Cells(lastRow + index, 1).Font.ColorIndex = 2
#     #             nrRows = index
#     #
#     #
#     #     if elem["localisation"] is None or elem["localisation"] == "":
#     #         testReportSheet.Cells(lastRow, 3).Value = "OK"
#     #     else:
#     #         testReportSheet.Cells(lastRow, 3).Value = elem["message"]
#     #
#     #     if elem["localisation"] is None or elem["localisation"] == "":
#     #         testReportSheet.Cells(lastRow, 4).Value = elem["localisation"]
#     #         if nrRows > 0:
#     #             lastRow += nrRows
#     #         else:
#     #             lastRow += 1
#     #     else:
#     #         if isinstance(elem["localisation"][0],str):
#     #             for index, element in enumerate(elem["localisation"]):
#     #                 testReportSheet.Cells(lastRow + index, 4).Value = element
#     #                 lastRow += 1
#     #         else:
#     #             for index, element in enumerate(elem["localisation"]):
#     #                 testReportSheet.Cells(lastRow + index, 4).Formula = "=HYPERLINK(\"#\'" + element.Worksheet.Name + "\'!"+ element.Address + "\",\"" + element.Address +"\")"
#     #             lastRow = lastRow + index + 1
#
#
#     # ########################################################################################################
#
#
#     index_test_report = -1
#     index_info_report = -1
#     for sheetname in TSDApp.WorkbookStats.sheetNames:
#         if sheetname == 'test report':
#             index_test_report = TSDApp.WorkbookStats.sheetNames.index('test report')
#         if sheetname == 'report information':
#             index_info_report = TSDApp.WorkbookStats.sheetNames.index('report information')
#
#     # DOC3 = xlrd.open_workbook(workBook, on_demand=True)
#     DOC3 = TSDApp.DOC3Workbook
#     list_alpha = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ']
#
#     if index_test_report != -1:
#         workSheet_test_report = DOC3.sheet_by_index(index_test_report)
#         nrCols_test_report = workSheet_test_report.ncols
#         nrRows_test_report = workSheet_test_report.nrows
#
#     if index_info_report != -1:
#         workSheet_info_report = DOC3.sheet_by_index(index_info_report)
#         nrCols_info_report = workSheet_info_report.ncols
#         nrRows_info_report = workSheet_info_report.nrows
#
#     workBook2 = copy(DOC3)
#
#     if index_info_report != -1:
#         workSheet_info_report = workBook2.get_sheet(index_info_report)
#         # for index1 in range(0, nrRows_info_report):
#         #     for index2 in range(0, nrCols_info_report):
#         #         workSheet_info_report.write(index1, index2, "")
#
#         workSheet_info_report._cell_overwrite_ok = True
#
#         workSheet_info_report.write(0, 0, "Tool version:")
#         workSheet_info_report.write(0, 1, TSD_Checker_V5_6.appName)
#
#         workSheet_info_report.write(2, 0, "Criticity configuration file:")
#         workSheet_info_report.write(2, 1, TSDApp.DOC9Path)
#
#         workSheet_info_report.write(3, 0, "Extract CESARE file:")
#         workSheet_info_report.write(3, 1, TSDApp.DOC8Path)
#
#         workSheet_info_report.write(4, 0, "Customer effects file:")
#         workSheet_info_report.write(4, 1, TSDApp.DOC7Name)
#
#         workSheet_info_report.write(5, 0, "Diversity management file:")
#         workSheet_info_report.write(5, 1, TSDApp.DOC13Path)
#
#         workSheet_info_report.write(6, 0, "CESARE file reference:")
#         workSheet_info_report.write(6, 1, TSDApp.DOC8Link.split("/")[-3])
#
#         workSheet_info_report.write(7, 0, "Criticity configuration file reference:")
#         workSheet_info_report.write(7, 1, TSDApp.DOC9Link.split("/")[-3])
#
#         workSheet_info_report.write(8, 0, "Customer effect file reference:")
#         workSheet_info_report.write(8, 1, TSDApp.DOC7Link.split("/")[-3])
#
#         workSheet_info_report.write(9, 0, "Diversity management file reference:")
#         workSheet_info_report.write(9, 1, TSDApp.DOC13Link.split("/")[-3])
#
#         workSheet_info_report.write(10, 0, "Check level:")
#         workSheet_info_report.write(10, 1, TSDApp.checkLevel)
#
#         workSheet_info_report.write(12, 0, "Date of the test:")
#         workSheet_info_report.write(12, 1, time.strftime("%x"))
#
#         workSheet_info_report.write(13, 0, "Time of the test:")
#         workSheet_info_report.write(13, 1, time.strftime("%X"))
#
#         workSheet_info_report.write(14, 0, "Test duration:")
#         workSheet_info_report.write(14, 1, time.strftime('%H:%M:%S', time.gmtime(TSDApp.end_time - TSDApp.start_time)))
#
#         workSheet_info_report.write(15, 0, "Opening duration:")
#         workSheet_info_report.write(15, 1,time.strftime('%H:%M:%S', time.gmtime(TSDApp.opening_time - TSDApp.start_time)))
#
#         workSheet_info_report.write(17, 0, "TSD file checked:")
#         workSheet_info_report.write(17, 1, TSDApp.DOC3Path)
#
#         workSheet_info_report.write(18, 0, "TSD function file checked:")
#         workSheet_info_report.write(18, 1, TSDApp.DOC4Path)
#
#         workSheet_info_report.write(19, 0, "TSD system file checked:")
#         workSheet_info_report.write(19, 1, TSDApp.DOC5Path)
#
#         workSheet_info_report.write(21, 0, "AMDEC:")
#         workSheet_info_report.write(21, 1, TSDApp.AMDECName)
#
#         workSheet_info_report.write(22, 0, "Export MedialecMatrice:")
#         workSheet_info_report.write(22, 1, TSDApp.MedialecName)
#
#         workSheet_info_report.write(24, 0, "Status:")
#         workSheet_info_report.write(24, 1, str(TSDApp.status))
#
#         workSheet_info_report.write(25, 0, "Coverage Indicator:")
#         workSheet_info_report.write(25, 1, str(TSDApp.coverage)[0:4] + "%")
#
#         workSheet_info_report.write(26, 0, "Convergence Indicator:")
#         workSheet_info_report.write(26, 1, str(TSDApp.convergence)[0:4] + "%")
#
#         workSheet_info_report.write(28, 0, "Blocking Points")
#         workSheet_info_report.write(28, 1, str(TSDApp.criticity_blocking))
#
#         workSheet_info_report.write(29, 0, "Warning Points")
#         workSheet_info_report.write(29, 1, str(TSDApp.criticity_warning))
#
#         workSheet_info_report.write(30, 0, "Information Points")
#         workSheet_info_report.write(30, 1, str(TSDApp.criticity_information))
#
#         workSheet_info_report._cell_overwrite_ok = False
#     else:
#         workSheet_info_report = workBook2.add_sheet('Report information', cell_overwrite_ok=True)
#
#         workSheet_info_report.write(0,0,"Tool version:")
#         workSheet_info_report.write(0, 1, TSD_Checker_V5_6.appName)
#
#         workSheet_info_report.write(2, 0, "Criticity configuration file:")
#         workSheet_info_report.write(2, 1, TSDApp.DOC9Path)
#
#         workSheet_info_report.write(3, 0, "Extract CESARE file:")
#         workSheet_info_report.write(3, 1, TSDApp.DOC8Path)
#
#         workSheet_info_report.write(4, 0, "Customer effects file:")
#         workSheet_info_report.write(4, 1, TSDApp.DOC7Name)
#
#         workSheet_info_report.write(5, 0, "Diversity management file:")
#         workSheet_info_report.write(5, 1, TSDApp.DOC13Path)
#
#         workSheet_info_report.write(6, 0, "CESARE file reference:")
#         workSheet_info_report.write(6, 1, TSDApp.DOC8Link.split("/")[-3])
#
#         workSheet_info_report.write(7, 0, "Criticity configuration file reference:")
#         workSheet_info_report.write(7, 1, TSDApp.DOC9Link.split("/")[-3])
#
#         workSheet_info_report.write(8, 0, "Customer effect file reference:")
#         workSheet_info_report.write(8, 1, TSDApp.DOC7Link.split("/")[-3])
#
#         workSheet_info_report.write(9, 0, "Diversity management file reference:")
#         workSheet_info_report.write(9, 1, TSDApp.DOC13Link.split("/")[-3])
#
#         workSheet_info_report.write(10, 0, "Check level:")
#         workSheet_info_report.write(10, 1, TSDApp.checkLevel)
#
#         workSheet_info_report.write(12, 0, "Date of the test:")
#         workSheet_info_report.write(12, 1, time.strftime("%x"))
#
#         workSheet_info_report.write(13, 0, "Time of the test:")
#         workSheet_info_report.write(13, 1, time.strftime("%X"))
#
#         workSheet_info_report.write(14, 0, "Test duration:")
#         workSheet_info_report.write(14, 1, time.strftime('%H:%M:%S', time.gmtime(TSDApp.end_time - TSDApp.start_time)))
#
#         workSheet_info_report.write(15, 0, "Opening duration:")
#         workSheet_info_report.write(15, 1, time.strftime('%H:%M:%S', time.gmtime(TSDApp.opening_time - TSDApp.start_time)))
#
#         workSheet_info_report.write(17, 0, "TSD file checked:")
#         workSheet_info_report.write(17, 1, TSDApp.DOC3Path)
#
#         workSheet_info_report.write(18, 0, "TSD function file checked:")
#         workSheet_info_report.write(18, 1, TSDApp.DOC4Path)
#
#         workSheet_info_report.write(19, 0, "TSD system file checked:")
#         workSheet_info_report.write(19, 1, TSDApp.DOC5Path)
#
#         workSheet_info_report.write(21, 0, "AMDEC:")
#         workSheet_info_report.write(21, 1, TSDApp.AMDECName)
#
#         workSheet_info_report.write(22, 0, "Export MedialecMatrice:")
#         workSheet_info_report.write(22, 1, TSDApp.MedialecName)
#
#         workSheet_info_report.write(24, 0, "Status:")
#         workSheet_info_report.write(24, 1, str(TSDApp.status))
#
#         workSheet_info_report.write(25, 0, "Coverage Indicator:")
#         workSheet_info_report.write(25, 1, str(TSDApp.coverage)[0:4] + "%")
#
#         workSheet_info_report.write(26, 0, "Convergence Indicator:")
#         workSheet_info_report.write(26, 1, str(TSDApp.convergence)[0:4] + "%")
#
#         workSheet_info_report.write(28, 0, "Blocking Points")
#         workSheet_info_report.write(28, 1, str(TSDApp.criticity_blocking))
#
#         workSheet_info_report.write(29, 0, "Warning Points")
#         workSheet_info_report.write(29, 1, str(TSDApp.criticity_warning))
#
#         workSheet_info_report.write(30, 0, "Information Points")
#         workSheet_info_report.write(30, 1, str(TSDApp.criticity_information))
#
#
#
#     if index_test_report != -1:
#         workSheet_test_report = workBook2.get_sheet(index_test_report)
#
#
#         for index1 in range(0, nrRows_test_report):
#             for index2 in range(0, nrCols_test_report):
#                 workSheet_test_report.write(index1, index2, "")
#
#         workSheet_test_report._cell_overwrite_ok = True
#
#
#         lastRow = 0
#         workSheet_test_report.write(lastRow, 0, 'Criticity')
#         workSheet_test_report.write(lastRow, 1, 'Requirements')
#         workSheet_test_report.write(lastRow, 2, 'Message')
#         workSheet_test_report.write(lastRow, 3, 'Localisation')
#
#         lastRow += 1
#         blocking_style = xlwt.easyxf('pattern: pattern solid, fore_colour red;')
#         warning_style = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')
#         text_style = xlwt.easyxf('font: colour white, bold False;')
#
#         for elem in return_list:
#
#             if elem["criticity"].casefold().strip() == "blocking":
#                 workSheet_test_report.write(lastRow, 0, elem["criticity"], blocking_style)
#             elif elem["criticity"].casefold().strip() == "warning":
#                 workSheet_test_report.write(lastRow, 0, elem["criticity"], warning_style)
#             else:
#                 workSheet_test_report.write(lastRow, 0, elem["criticity"])
#
#             workSheet_test_report.write(lastRow, 1, elem["testName"])
#
#             if elem["localisation"] is None or elem["localisation"] == "":
#                 workSheet_test_report.write(lastRow, 2, "OK")
#             else:
#                 workSheet_test_report.write(lastRow, 2, elem["message"])
#
#             if elem["localisation"] is None or elem["localisation"] == "":
#                 workSheet_test_report.write(lastRow, 3, elem["localisation"])
#                 lastRow += 1
#
#             if elem["localisation"] is not None and elem["localisation"] != "":
#                 if isinstance(elem["localisation"][0], str):
#                     for index, element in enumerate(elem["localisation"]):
#                         workSheet_test_report.write(lastRow + index, 3, element)
#
#                     for index in range(1, len(elem["localisation"]) + 1):
#                         workSheet_test_report.write(lastRow + index, 0, elem["criticity"], text_style)
#                         workSheet_test_report.write(lastRow + index, 1, elem["testName"], text_style)
#
#                     lastRow += index
#                 else:
#                     for index, element in enumerate(elem["localisation"]):
#                         link = "HYPERLINK(\"#\'" + str(element[0]) + "\'!$" + str(list_alpha[element[2]]) + "$" + str(
#                             element[1] + 1) + "\",\"$" + str(list_alpha[element[2]]) + "$" + str(element[1] + 1) + "\")"
#                         workSheet_test_report.write(lastRow + index, 3, xlwt.Formula(link))
#
#                     for index in range(1, len(elem["localisation"]) + 1):
#                         workSheet_test_report.write(lastRow + index, 0, elem["criticity"], text_style)
#                         workSheet_test_report.write(lastRow + index, 1, elem["testName"], text_style)
#
#                     lastRow = lastRow + index
#
#         workSheet_test_report._cell_overwrite_ok = False
#
#     else:
#         workSheet_test_report = workBook2.add_sheet('Test report', cell_overwrite_ok=True)
#
#         lastRow = 0
#         workSheet_test_report.write(lastRow, 0, 'Criticity')
#         workSheet_test_report.write(lastRow, 1, 'Requirements')
#         workSheet_test_report.write(lastRow, 2, 'Message')
#         workSheet_test_report.write(lastRow, 3, 'Localisation')
#
#         lastRow += 1
#         blocking_style = xlwt.easyxf('pattern: pattern solid, fore_colour red;')
#         warning_style = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')
#         text_style = xlwt.easyxf('font: colour white, bold False;')
#
#         for elem in return_list:
#
#             if elem["criticity"].casefold().strip() == "blocking":
#                 workSheet_test_report.write(lastRow, 0, elem["criticity"], blocking_style)
#             elif elem["criticity"].casefold().strip() == "warning":
#                 workSheet_test_report.write(lastRow, 0, elem["criticity"], warning_style)
#             else:
#                 workSheet_test_report.write(lastRow, 0, elem["criticity"])
#
#             workSheet_test_report.write(lastRow, 1, elem["testName"])
#
#             if elem["localisation"] is None or elem["localisation"] == "":
#                 workSheet_test_report.write(lastRow, 2, "OK")
#             else:
#                 workSheet_test_report.write(lastRow, 2, elem["message"])
#
#             if elem["localisation"] is None or elem["localisation"] == "":
#                 workSheet_test_report.write(lastRow, 3, elem["localisation"])
#                 lastRow += 1
#
#             if elem["localisation"] is not None and elem["localisation"] != "":
#                 if isinstance(elem["localisation"][0], str):
#                     for index, element in enumerate(elem["localisation"]):
#                         workSheet_test_report.write(lastRow + index, 3, element)
#
#                     for index in range(1, len(elem["localisation"]) + 1):
#                         workSheet_test_report.write(lastRow + index, 0, elem["criticity"], text_style)
#                         workSheet_test_report.write(lastRow + index, 1, elem["testName"], text_style)
#
#                     lastRow += index
#                 else:
#                     for index, element in enumerate(elem["localisation"]):
#
#                         link = "HYPERLINK(\"#\'"+ str(element[0]) + "\'!$" + str(list_alpha[element[2]]) + "$" + str(element[1]+1) +"\",\"$" + str(list_alpha[element[2]]) + "$" + str(element[1]+1) + "\")"
#                         workSheet_test_report.write(lastRow + index, 3, xlwt.Formula(link))
#
#                     for index in range(1, len(elem["localisation"]) + 1):
#                         workSheet_test_report.write(lastRow + index, 0, elem["criticity"], text_style)
#                         workSheet_test_report.write(lastRow + index, 1, elem["testName"], text_style)
#
#                     lastRow = lastRow + index
#
#     workBook2.save(TSDApp.DOC3Path)
#
#
#     # ws = wb.add_sheet('Test report',  cell_overwrite_ok=True)
#     # ws11 = wb.add_sheet('codes défauts', cell_overwrite_ok=True)
#     # lastRow = 0
#     #
#     # ws.write(lastRow, 0, 'Criticity')
#     # ws.write(lastRow, 1, 'Requirements')
#     # ws.write(lastRow, 2, 'Message')
#     # ws.write(lastRow, 3, 'Localisation')
#     #
#     # lastRow += 1
#     # blocking_style = xlwt.easyxf('pattern: pattern solid, fore_colour red;')
#     # warning_style = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')
#     # text_style = xlwt.easyxf('font: colour white, bold False;')
#     #
#     # for elem in return_list:
#     #
#     #     if elem["criticity"].casefold().strip() == "blocking":
#     #         ws.write(lastRow, 0, elem["criticity"], blocking_style)
#     #     elif elem["criticity"].casefold().strip() == "warning":
#     #         ws.write(lastRow, 0, elem["criticity"], warning_style)
#     #     else:
#     #         ws.write(lastRow, 0, elem["criticity"])
#     #
#     #     ws.write(lastRow, 1, elem["testName"])
#     #
#     #     if elem["localisation"] is None or elem["localisation"] == "":
#     #         ws.write(lastRow, 2, "OK")
#     #     else:
#     #         ws.write(lastRow, 2, elem["message"])
#     #
#     #     if elem["localisation"] is None or elem["localisation"] == "":
#     #         ws.write(lastRow, 3, elem["localisation"])
#     #         lastRow += 1
#     #
#     #     if elem["localisation"] is not None and elem["localisation"] != "":
#     #         if isinstance(elem["localisation"][0], str):
#     #             for index, element in enumerate(elem["localisation"]):
#     #                 ws.write(lastRow + index, 3, element)
#     #
#     #             for index in range(1, len(elem["localisation"]) + 1):
#     #                 ws.write(lastRow + index, 0, elem["criticity"], text_style)
#     #                 ws.write(lastRow + index, 1, elem["testName"], text_style)
#     #
#     #             lastRow += index
#     #         else:
#     #             for index, element in enumerate(elem["localisation"]):
#     #
#     #                 link = "HYPERLINK(\"#\'"+ str(element.Worksheet.Name) + "\'!" + str(element.Address) + "\",\"" + str(element.Address) + "\")"
#     #                 ws.write(lastRow + index, 3, xlwt.Formula(link))
#     #
#     #             for index in range(1, len(elem["localisation"]) + 1):
#     #                 ws.write(lastRow + index, 0, elem["criticity"], text_style)
#     #                 ws.write(lastRow + index, 1, elem["testName"], text_style)
#     #
#     #             lastRow = lastRow + index
#     #
#     #
#     # wb.save('C:\\Users\\msnecula\\Downloads\\output.xls')
#
#     # ########################################################################################################
#
#
#     # def TestReturn(criticity, testName, message, localisation, workBook, TSDApp):
# #     testReportSheet = workBook.Sheets("Test report")
# #     lastRow = testReportSheet.UsedRange.Rows.Count + 1
# #
# #     testReportSheet.Cells(lastRow, 1).Value = criticity
# #
# #
# #     ColorCell(criticity, testReportSheet.Cells(lastRow, 1))
# #     tempString = str()
# #
# #     testReportSheet.Cells(lastRow, 2).Value = testName
# #     if localisation is not None:
# #         localisation_len = len(localisation)
# #         for index in range(1,localisation_len):
# #             testReportSheet.Cells(lastRow + index, 2).Value = testName
# #             testReportSheet.Cells(lastRow + index, 2).Font.ColorIndex = 2
# #             testReportSheet.Cells(lastRow + index, 1).Value = criticity
# #             testReportSheet.Cells(lastRow + index, 1).Font.ColorIndex = 2
# #
# #     if localisation is None or localisation == "":
# #         testReportSheet.Cells(lastRow, 3).Value = "OK"
# #         tempString = "OK"
# #     else:
# #         testReportSheet.Cells(lastRow, 3).Value = message
# #         tempString = "NOK"
# #
# #     if localisation is None or localisation == "":
# #         testReportSheet.Cells(lastRow, 4).Value = localisation
# #     else:
# #         for index, element in enumerate(localisation):
# #             testReportSheet.Cells(lastRow + index, 4).Formula = "=HYPERLINK(\"#\'" + element.Worksheet.Name + "\'!"+ element.Address + "\",\"" + element.Address +"\")"
# #
# #     textBoxText = TSDApp.tab1.textbox.toPlainText()
# #     textBoxText = textBoxText + "\n" + testName + " " + tempString
# #     TSDApp.tab1.textbox.setText(textBoxText)
# #     TSDApp.tab1.textbox.moveCursor(QtGui.QTextCursor.End)
# #
# #     TSDApp.IncrementProgressBar()
#
#
# def ExcelWrite2(return_list, workBook, TSDApp):
#     wb = openpyxl.load_workbook(workBook, keep_vba=True)
#
#     index_test_report = -1
#     index_info_report = -1
#     for sheetname in TSDApp.WorkbookStats.sheetNames:
#         if sheetname == 'test report':
#             index_test_report = TSDApp.WorkbookStats.sheetNames.index('test report')
#         if sheetname == 'report information':
#             index_info_report = TSDApp.WorkbookStats.sheetNames.index('report information')
#
#     fontStyle_b = Font(name="Calibri", size=12, color=colors.BLACK)
#     fontStyle_w = Font(name="Calibri", size=12, color=colors.WHITE)
#
#     if index_info_report == -1:
#         workSheet_info_report = wb.create_sheet("Report information")
#
#         workSheet_info_report['A1'] = "Tool version:"
#         # workSheet_info_report['A1'].font = fontStyle_b
#         workSheet_info_report['B1'] = TSD_Checker_V5_6.appName
#         # workSheet_info_report['B1'].font = fontStyle_b
#
#         workSheet_info_report['A3'] = "Criticity configuration file:"
#         # workSheet_info_report['A3'].font = fontStyle_b
#         workSheet_info_report['B3'] = TSDApp.DOC9Path
#         # workSheet_info_report['B3'].font = fontStyle_b
#
#         workSheet_info_report['A4'] = "Extract CESARE file:"
#         # workSheet_info_report['A4'].font = fontStyle_b
#         workSheet_info_report['B4'] = TSDApp.DOC8Path
#         # workSheet_info_report['B4'].font = fontStyle_b
#
#         workSheet_info_report['A5'] = "Customer effects file:"
#         # workSheet_info_report['A5'].font = fontStyle_b
#         workSheet_info_report['B5'] = TSDApp.DOC7Name
#         # workSheet_info_report['B5'].font = fontStyle_b
#
#         workSheet_info_report['A6'] = "Diversity management file:"
#         # workSheet_info_report['A6'].font = fontStyle_b
#         workSheet_info_report['B6'] = TSDApp.DOC13Path
#         # workSheet_info_report['B6'].font = fontStyle_b
#
#         workSheet_info_report['A7'] = "CESARE file reference:"
#         # workSheet_info_report['A7'].font = fontStyle_b
#         workSheet_info_report['B7'] = TSDApp.DOC8Link.split("/")[-3]
#         # workSheet_info_report['B7'].font = fontStyle_b
#
#         workSheet_info_report['A8'] = "Criticity configuration file reference:"
#         # workSheet_info_report['A8'].font = fontStyle_b
#         workSheet_info_report['B8'] = TSDApp.DOC9Link.split("/")[-3]
#         # workSheet_info_report['B8'].font = fontStyle_b
#
#         workSheet_info_report['A9'] = "Customer effect file reference:"
#         # workSheet_info_report['A9'].font = fontStyle_b
#         workSheet_info_report['B9'] = TSDApp.DOC7Link.split("/")[-3]
#         # workSheet_info_report['B9'].font = fontStyle_b
#
#         workSheet_info_report['A10'] = "Diversity management file reference:"
#         # workSheet_info_report['A10'].font = fontStyle_b
#         workSheet_info_report['B10'] = TSDApp.DOC13Link.split("/")[-3]
#         # workSheet_info_report['B10'].font = fontStyle_b
#
#         workSheet_info_report['A11'] = "Check level:"
#         # workSheet_info_report['A11'].font = fontStyle_b
#         workSheet_info_report['B11'] = TSDApp.checkLevel
#         # workSheet_info_report['B11'].font = fontStyle_b
#
#         workSheet_info_report['A13'] = "Date of the test:"
#         # workSheet_info_report['A13'].font = fontStyle_b
#         workSheet_info_report['B13'] = time.strftime("%x")
#         # workSheet_info_report['B13'].font = fontStyle_b
#
#         workSheet_info_report['A14'] = "Time of the test:"
#         # workSheet_info_report['A14'].font = fontStyle_b
#         workSheet_info_report['B14'] = time.strftime("%X")
#         # workSheet_info_report['B14'].font = fontStyle_b
#
#         workSheet_info_report['A15'] = "Test duration:"
#         # workSheet_info_report['A15'].font = fontStyle_b
#         workSheet_info_report['B15'] = time.strftime('%H:%M:%S', time.gmtime(TSDApp.end_time - TSDApp.start_time))
#         # workSheet_info_report['B15'].font = fontStyle_b
#
#         workSheet_info_report['A16'] = "Opening duration:"
#         # workSheet_info_report['A16'].font = fontStyle_b
#         workSheet_info_report['B16'] = time.strftime('%H:%M:%S', time.gmtime(TSDApp.opening_time - TSDApp.start_time))
#         # workSheet_info_report['B16'].font = fontStyle_b
#
#         workSheet_info_report['A18'] = "TSD file checked:"
#         # workSheet_info_report['A18'].font = fontStyle_b
#         workSheet_info_report['B18'] = TSDApp.DOC3Path
#         # workSheet_info_report['B18'].font = fontStyle_b
#
#         workSheet_info_report['A19'] = "TSD function file checked:"
#         # workSheet_info_report['A19'].font = fontStyle_b
#         workSheet_info_report['B19'] = TSDApp.DOC4Path
#         # workSheet_info_report['B19'].font = fontStyle_b
#
#         workSheet_info_report['A20'] = "TSD system file checked:"
#         # workSheet_info_report['A20'].font = fontStyle_b
#         workSheet_info_report['B20'] = TSDApp.DOC5Path
#         # workSheet_info_report['B20'].font = fontStyle_b
#
#         workSheet_info_report['A22'] = "AMDEC:"
#         # workSheet_info_report['A22'].font = fontStyle_b
#         workSheet_info_report['B22'] = TSDApp.AMDECName
#         # workSheet_info_report['B22'].font = fontStyle_b
#
#         workSheet_info_report['A23'] = "Export MedialecMatrice:"
#         # workSheet_info_report['A23'].font = fontStyle_b
#         workSheet_info_report['B23'] = TSDApp.MedialecName
#         # workSheet_info_report['B23'].font = fontStyle_b
#
#         workSheet_info_report['A25'] = "Status:"
#         # workSheet_info_report['A25'].font = fontStyle_b
#         workSheet_info_report['B25'] = str(TSDApp.status)
#         # workSheet_info_report['B25'].font = fontStyle_b
#
#         workSheet_info_report['A26'] = "Coverage Indicator:"
#         # workSheet_info_report['A26'].font = fontStyle_b
#         workSheet_info_report['B26'] = str(TSDApp.coverage)[0:4] + "%"
#         # workSheet_info_report['B26'].font = fontStyle_b
#
#         workSheet_info_report['A27'] = "Convergence Indicator:"
#         # workSheet_info_report['A27'].font = fontStyle_b
#         workSheet_info_report['B27'] = str(TSDApp.convergence)[0:4] + "%"
#         # workSheet_info_report['B27'].font = fontStyle_b
#
#         workSheet_info_report['A29'] = "Blocking Points"
#         # workSheet_info_report['A29'].font = fontStyle_b
#         workSheet_info_report['B29'] = str(TSDApp.criticity_blocking)
#         # workSheet_info_report['B29'].font = fontStyle_b
#
#         workSheet_info_report['A30'] = "Warning Points"
#         # workSheet_info_report['A30'].font = fontStyle_b
#         workSheet_info_report['B30'] = str(TSDApp.criticity_warning)
#         # workSheet_info_report['B30'].font = fontStyle_b
#
#         workSheet_info_report['A31'] = "Information Points"
#         # workSheet_info_report['A31'].font = fontStyle_b
#         workSheet_info_report['B31'] = str(TSDApp.criticity_information)
#         # workSheet_info_report['B31'].font = fontStyle_b
#
#     else:
#         workSheet_info_report = wb.get_sheet_by_name("Report information")
#         wb.remove_sheet(workSheet_info_report)
#         workSheet_info_report = wb.create_sheet("Report information")
#
#         workSheet_info_report['A1'] = "Tool version:"
#         # workSheet_info_report['A1'].font = fontStyle_b
#         workSheet_info_report['B1'] = TSD_Checker_V5_6.appName
#         # workSheet_info_report['B1'].font = fontStyle_b
#
#         workSheet_info_report['A3'] = "Criticity configuration file:"
#         # workSheet_info_report['A3'].font = fontStyle_b
#         workSheet_info_report['B3'] = TSDApp.DOC9Path
#         # workSheet_info_report['B3'].font = fontStyle_b
#
#         workSheet_info_report['A4'] = "Extract CESARE file:"
#         # workSheet_info_report['A4'].font = fontStyle_b
#         workSheet_info_report['B4'] = TSDApp.DOC8Path
#         # workSheet_info_report['B4'].font = fontStyle_b
#
#         workSheet_info_report['A5'] = "Customer effects file:"
#         # workSheet_info_report['A5'].font = fontStyle_b
#         workSheet_info_report['B5'] = TSDApp.DOC7Name
#         # workSheet_info_report['B5'].font = fontStyle_b
#
#         workSheet_info_report['A6'] = "Diversity management file:"
#         # workSheet_info_report['A6'].font = fontStyle_b
#         workSheet_info_report['B6'] = TSDApp.DOC13Path
#         # workSheet_info_report['B6'].font = fontStyle_b
#
#         workSheet_info_report['A7'] = "CESARE file reference:"
#         # workSheet_info_report['A7'].font = fontStyle_b
#         workSheet_info_report['B7'] = TSDApp.DOC8Link.split("/")[-3]
#         # workSheet_info_report['B7'].font = fontStyle_b
#
#         workSheet_info_report['A8'] = "Criticity configuration file reference:"
#         # workSheet_info_report['A8'].font = fontStyle_b
#         workSheet_info_report['B8'] = TSDApp.DOC9Link.split("/")[-3]
#         # workSheet_info_report['B8'].font = fontStyle_b
#
#         workSheet_info_report['A9'] = "Customer effect file reference:"
#         # workSheet_info_report['A9'].font = fontStyle_b
#         workSheet_info_report['B9'] = TSDApp.DOC7Link.split("/")[-3]
#         # workSheet_info_report['B9'].font = fontStyle_b
#
#         workSheet_info_report['A10'] = "Diversity management file reference:"
#         # workSheet_info_report['A10'].font = fontStyle_b
#         workSheet_info_report['B10'] = TSDApp.DOC13Link.split("/")[-3]
#         # workSheet_info_report['B10'].font = fontStyle_b
#
#         workSheet_info_report['A11'] = "Check level:"
#         # workSheet_info_report['A11'].font = fontStyle_b
#         workSheet_info_report['B11'] = TSDApp.checkLevel
#         # workSheet_info_report['B11'].font = fontStyle_b
#
#         workSheet_info_report['A13'] = "Date of the test:"
#         # workSheet_info_report['A13'].font = fontStyle_b
#         workSheet_info_report['B13'] = time.strftime("%x")
#         # workSheet_info_report['B13'].font = fontStyle_b
#
#         workSheet_info_report['A14'] = "Time of the test:"
#         # workSheet_info_report['A14'].font = fontStyle_b
#         workSheet_info_report['B14'] = time.strftime("%X")
#         # workSheet_info_report['B14'].font = fontStyle_b
#
#         workSheet_info_report['A15'] = "Test duration:"
#         # workSheet_info_report['A15'].font = fontStyle_b
#         workSheet_info_report['B15'] = time.strftime('%H:%M:%S', time.gmtime(TSDApp.end_time - TSDApp.start_time))
#         # workSheet_info_report['B15'].font = fontStyle_b
#
#         workSheet_info_report['A16'] = "Opening duration:"
#         # workSheet_info_report['A16'].font = fontStyle_b
#         workSheet_info_report['B16'] = time.strftime('%H:%M:%S', time.gmtime(TSDApp.opening_time - TSDApp.start_time))
#         # workSheet_info_report['B16'].font = fontStyle_b
#
#         workSheet_info_report['A18'] = "TSD file checked:"
#         # workSheet_info_report['A18'].font = fontStyle_b
#         workSheet_info_report['B18'] = TSDApp.DOC3Path
#         # workSheet_info_report['B18'].font = fontStyle_b
#
#         workSheet_info_report['A19'] = "TSD function file checked:"
#         # workSheet_info_report['A19'].font = fontStyle_b
#         workSheet_info_report['B19'] = TSDApp.DOC4Path
#         # workSheet_info_report['B19'].font = fontStyle_b
#
#         workSheet_info_report['A20'] = "TSD system file checked:"
#         # workSheet_info_report['A20'].font = fontStyle_b
#         workSheet_info_report['B20'] = TSDApp.DOC5Path
#         # workSheet_info_report['B20'].font = fontStyle_b
#
#         workSheet_info_report['A22'] = "AMDEC:"
#         # workSheet_info_report['A22'].font = fontStyle_b
#         workSheet_info_report['B22'] = TSDApp.AMDECName
#         # workSheet_info_report['B22'].font = fontStyle_b
#
#         workSheet_info_report['A23'] = "Export MedialecMatrice:"
#         # workSheet_info_report['A23'].font = fontStyle_b
#         workSheet_info_report['B23'] = TSDApp.MedialecName
#         # workSheet_info_report['B23'].font = fontStyle_b
#
#         workSheet_info_report['A25'] = "Status:"
#         # workSheet_info_report['A25'].font = fontStyle_b
#         workSheet_info_report['B25'] = str(TSDApp.status)
#         # workSheet_info_report['B25'].font = fontStyle_b
#
#         workSheet_info_report['A26'] = "Coverage Indicator:"
#         # workSheet_info_report['A26'].font = fontStyle_b
#         workSheet_info_report['B26'] = str(TSDApp.coverage)[0:4] + "%"
#         # workSheet_info_report['B26'].font = fontStyle_b
#
#         workSheet_info_report['A27'] = "Convergence Indicator:"
#         # workSheet_info_report['A27'].font = fontStyle_b
#         workSheet_info_report['B27'] = str(TSDApp.convergence)[0:4] + "%"
#         # workSheet_info_report['B27'].font = fontStyle_b
#
#         workSheet_info_report['A29'] = "Blocking Points"
#         # workSheet_info_report['A29'].font = fontStyle_b
#         workSheet_info_report['B29'] = str(TSDApp.criticity_blocking)
#         # workSheet_info_report['B29'].font = fontStyle_b
#
#         workSheet_info_report['A30'] = "Warning Points"
#         # workSheet_info_report['A30'].font = fontStyle_b
#         workSheet_info_report['B30'] = str(TSDApp.criticity_warning)
#         # workSheet_info_report['B30'].font = fontStyle_b
#
#         workSheet_info_report['A31'] = "Information Points"
#         # workSheet_info_report['A31'].font = fontStyle_b
#         workSheet_info_report['B31'] = str(TSDApp.criticity_information)
#         # workSheet_info_report['B31'].font = fontStyle_b
#
#     workSheet_info_report.column_dimensions['A'].width = 40
#     workSheet_info_report.column_dimensions['B'].width = 140
#
#     list_alpha = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ']
#     if index_test_report == -1:
#         workSheet_test_report = wb.create_sheet("Test report")
#
#         lastRow = 1
#         workSheet_test_report['A1'] = "Criticity"
#         # workSheet_test_report['A1'].font = fontStyle_b
#         workSheet_test_report['B1'] = "Requirements"
#         # workSheet_test_report['B1'].font = fontStyle_b
#         workSheet_test_report['C1'] = "Message"
#         # workSheet_test_report['C1'].font = fontStyle_b
#         workSheet_test_report['D1'] = "Localisation"
#         # workSheet_test_report['D1'].font = fontStyle_b
#
#
#         lastRow += 1
#
#         my_red = openpyxl.styles.colors.Color(rgb='00FF0000')
#         blocking_style = openpyxl.styles.fills.PatternFill(patternType='solid', fgColor=my_red)
#         my_yellow = openpyxl.styles.colors.Color(rgb='00FFFF00')
#         warning_style = openpyxl.styles.fills.PatternFill(patternType='solid', fgColor=my_yellow)
#         text_style = Font(color='FFFFFFFF')
#
#
#         for elem in return_list:
#
#             if elem["criticity"].casefold().strip() == "blocking":
#                 workSheet_test_report.cell(lastRow, 1, elem["criticity"]).font = blocking_style
#             elif elem["criticity"].casefold().strip() == "warning":
#                 workSheet_test_report.cell(lastRow, 1, elem["criticity"]).fill = warning_style
#             else:
#                 workSheet_test_report.cell(lastRow, 1, elem["criticity"])
#
#             workSheet_test_report.cell(lastRow, 2, elem["testName"])
#
#             if elem["localisation"] is None or elem["localisation"] == "":
#                 workSheet_test_report.cell(lastRow, 3, "OK")
#             else:
#                 workSheet_test_report.cell(lastRow, 3, elem["message"])
#
#             if elem["localisation"] is None or elem["localisation"] == "":
#                 workSheet_test_report.cell(lastRow, 4, elem["localisation"])
#                 lastRow += 1
#
#             if elem["localisation"] is not None and elem["localisation"] != "":
#                 if isinstance(elem["localisation"][0], str):
#                     for index, element in enumerate(elem["localisation"]):
#                         workSheet_test_report.cell(lastRow + index, 4, element)
#
#                     if len(elem['localisation']) > 1:
#                         for index in range(1, len(elem["localisation"])):
#                             workSheet_test_report.cell(lastRow + index, 1, elem["criticity"]).font = text_style
#                             workSheet_test_report.cell(lastRow + index, 2, elem["testName"]).font = text_style
#
#                     lastRow += index + 1
#                 else:
#                     for index, element in enumerate(elem["localisation"]):
#                         workSheet_test_report.cell(lastRow + index, 4).value = '$' + str(list_alpha[element[2]]) + '$' + str(element[1] + 1)
#                         workSheet_test_report.cell(lastRow + index, 4).hyperlink = '#%s!%s' % ("'" + str(element[0]) + "'", str(list_alpha[element[2]])+str(element[1] + 1) )
#
#
#                     if len(elem['localisation']) > 1:
#                         for index in range(1, len(elem["localisation"])):
#                             workSheet_test_report.cell(lastRow + index, 1, elem["criticity"]).font = text_style
#                             workSheet_test_report.cell(lastRow + index, 2, elem["testName"]).font = text_style
#
#                     lastRow += index + 1
#     else:
#         workSheet_test_report = wb.get_sheet_by_name("Test report")
#         wb.remove_sheet(workSheet_test_report)
#         workSheet_test_report = wb.create_sheet("Test report")
#
#         lastRow = 1
#         workSheet_test_report['A1'] = "Criticity"
#         # workSheet_test_report['A1'].font = fontStyle_b
#         workSheet_test_report['B1'] = "Requirements"
#         # workSheet_test_report['B1'].font = fontStyle_b
#         workSheet_test_report['C1'] = "Message"
#         # workSheet_test_report['C1'].font = fontStyle_b
#         workSheet_test_report['D1'] = "Localisation"
#         # workSheet_test_report['D1'].font = fontStyle_b
#
#         lastRow += 1
#
#         my_red = openpyxl.styles.colors.Color(rgb='00FF0000')
#         blocking_style = openpyxl.styles.fills.PatternFill(patternType='solid', fgColor=my_red)
#         my_yellow = openpyxl.styles.colors.Color(rgb='00FFFF00')
#         warning_style = openpyxl.styles.fills.PatternFill(patternType='solid', fgColor=my_yellow)
#         text_style = Font(color='FFFFFFFF')
#
#         for elem in return_list:
#
#             if elem["criticity"].casefold().strip() == "blocking":
#                 workSheet_test_report.cell(lastRow, 1, elem["criticity"]).fill = blocking_style
#             elif elem["criticity"].casefold().strip() == "warning":
#                 workSheet_test_report.cell(lastRow, 1, elem["criticity"]).fill = warning_style
#             else:
#                 workSheet_test_report.cell(lastRow, 1, elem["criticity"])
#
#             workSheet_test_report.cell(lastRow, 2, elem["testName"])
#
#             if elem["localisation"] is None or elem["localisation"] == "":
#                 workSheet_test_report.cell(lastRow, 3, "OK")
#             else:
#                 workSheet_test_report.cell(lastRow, 3, elem["message"])
#
#             if elem["localisation"] is None or elem["localisation"] == "":
#                 workSheet_test_report.cell(lastRow, 4, elem["localisation"])
#                 lastRow += 1
#
#             if elem["localisation"] is not None and elem["localisation"] != "":
#                 if isinstance(elem["localisation"][0], str):
#                     for index, element in enumerate(elem["localisation"]):
#                         workSheet_test_report.cell(lastRow + index, 4, element)
#
#                     if len(elem['localisation']) > 1:
#                         for index in range(1, len(elem["localisation"])):
#                             workSheet_test_report.cell(lastRow + index, 1, elem["criticity"]).font = text_style
#                             workSheet_test_report.cell(lastRow + index, 2, elem["testName"]).font = text_style
#
#                     lastRow += index + 1
#                 else:
#                     for index, element in enumerate(elem["localisation"]):
#                         workSheet_test_report.cell(lastRow + index, 4).value = '$' + str(list_alpha[element[2]]) + '$' + str(element[1] + 1)
#                         # workSheet_test_report.cell(lastRow + index, 4).hyperlink = '#%s %s %s!%s' % ("'" + str(element[0]).split(' ')[-3],str(element[0]).split(' ')[-2],str(element[0]).split(' ')[-1] + "'", str(list_alpha[element[2]])+str(element[1] + 1) )
#                         workSheet_test_report.cell(lastRow + index, 4).hyperlink = '#%s!%s' % ("'" + str(element[0]) + "'", str(list_alpha[element[2]]) + str(element[1] + 1))
#
#                     if len(elem['localisation']) > 1:
#                         for index in range(1, len(elem["localisation"])):
#                             workSheet_test_report.cell(lastRow + index, 1, elem["criticity"]).font = text_style
#                             workSheet_test_report.cell(lastRow + index, 2, elem["testName"]).font = text_style
#
#                     lastRow += index + 1
#
#     workSheet_test_report.column_dimensions['A'].width = 20
#     workSheet_test_report.column_dimensions['B'].width = 40
#     workSheet_test_report.column_dimensions['C'].width = 80
#     workSheet_test_report.column_dimensions['D'].width = 20
#
#     wb.save(workBook)
#
#
# def TestReturnName(criticity, testName, message, name, workBook, TSDApp):
#     return_dict = {}
#     return_dict["criticity"] = criticity
#     return_dict["testName"] = testName
#     return_dict["message"] = message
#     return_dict["localisation"] = name
#     TSDApp.return_list.append(return_dict)
#
#     if criticity.casefold() == "blocking":
#         TSDApp.criticity_blocking += 1
#     else:
#         if criticity.casefold() == "warning":
#             TSDApp.criticity_warning += 1
#         else:
#             TSDApp.criticity_information += 1
#
#     tempString = str()
#     if name is None or name == "":
#         tempString = "OK"
#     else:
#         tempString = "NOK"
#
#     textBoxText = TSDApp.tab1.textbox.toPlainText()
#     textBoxText = textBoxText + "\n" + testName + " " + tempString
#     TSDApp.tab1.textbox.setText(textBoxText)
#     TSDApp.tab1.textbox.moveCursor(QtGui.QTextCursor.End)
#
#     TSDApp.IncrementProgressBar()
#
#     # testReportSheet = workBook.Sheets("Test report")
#     # lastRow = testReportSheet.UsedRange.Rows.Count + 1
#     #
#     # testReportSheet.Cells(lastRow, 1).Value = criticity
#     # if criticity.casefold() == "blocking":
#     #     TSDApp.criticity_blocking += 1
#     # else:
#     #     if criticity.casefold() == "warning":
#     #         TSDApp.criticity_warning += 1
#     #     else:
#     #         TSDApp.criticity_information += 1
#     #
#     # ColorCell(criticity, testReportSheet.Cells(lastRow, 1))
#     # tempString = str()
#     #
#     # testReportSheet.Cells(lastRow, 2).Value = testName
#     # if name is not None:
#     #     name_len = len(name)
#     #     for index in range(1,name_len):
#     #         testReportSheet.Cells(lastRow + index, 2).Value = testName
#     #         testReportSheet.Cells(lastRow + index, 2).Font.ColorIndex = 2
#     #         testReportSheet.Cells(lastRow + index, 1).Value = criticity
#     #         testReportSheet.Cells(lastRow + index, 1).Font.ColorIndex = 2
#     #
#     # if name is None or name == "":
#     #     testReportSheet.Cells(lastRow, 3).Value = "OK"
#     #     tempString = "OK"
#     # else:
#     #     testReportSheet.Cells(lastRow, 3).Value = message
#     #     tempString = "NOK"
#     #
#     # if name is None or name == "":
#     #     testReportSheet.Cells(lastRow, 4).Value = name
#     # else:
#     #     for index, element in enumerate(name):
#     #         testReportSheet.Cells(lastRow + index, 4).Value = name
#     #
#     # textBoxText = TSDApp.tab1.textbox.toPlainText()
#     # textBoxText = textBoxText + "\n" + testName + " " + tempString
#     # TSDApp.tab1.textbox.setText(textBoxText)
#     # TSDApp.tab1.textbox.moveCursor(QtGui.QTextCursor.End)
#     #
#     # TSDApp.IncrementProgressBar()
#
# def ColorCell(criticity, cell):
#     if criticity == "Blocking":
#         cell.Interior.ColorIndex = 3
#     elif criticity == "Warning":
#         cell.Interior.ColorIndex = 12
#     elif criticity == "Information":
#         cell.Interior.ColorIndex = 6
#
# def AddTestReportSheets(workBook):
#     try:
#         workSheet = workBook.Sheets("Report information")
#         workSheet.Application.DisplayAlerts = False
#         workSheet.Delete()
#         workSheet.Application.DisplayAlerts = True
#     except:
#         pass
#     try:
#         workSheet = workBook.Sheets("Test report")
#         workSheet.Application.DisplayAlerts = False
#         workSheet.Delete()
#         workSheet.Application.DisplayAlerts = True
#     except:
#         pass
#
#     reportInfoWorkSheet = workBook.Sheets.Add(None,workBook.Sheets(workBook.Sheets.Count),1,None)
#     reportInfoWorkSheet.Name = "Report information"
#     testReportWorkSheet = workBook.Sheets.Add(None,workBook.Sheets(workBook.Sheets.Count),1,None)
#     testReportWorkSheet.Name = "Test report"
#
# def AddTestReportSheetHeader(workBook):
#     testReportWorkSheet = workBook.Sheets("Test report")
#     textList = ["Criticity", "Requirements", "Message", "Localisation"]
#     testReportWorkSheet.Range("A1:D1").Value = textList
#     testReportWorkSheet.Columns("A").ColumnWidth = 12
#     testReportWorkSheet.Columns("B").ColumnWidth = 35
#     testReportWorkSheet.Columns("C").ColumnWidth = 150
#     testReportWorkSheet.Columns("D").ColumnWidth = 12
#
#
#     #testReportWorkSheet.Range("A1:D1").Font.Bold = True
#
# def WriteReportInformationSheet(workBook, TSDApp):
#     reportInformationWorkSheet = workBook.Sheets("Report information")
#     colList = list()
#     colList.append(list(("Tool version:", TSD_Checker_V5_6.appName)))
#     colList.append(list(("Criticity configuration file:", TSDApp.DOC9Path)))
#     colList.append(list(("","")))
#     colList.append(list(("Extract CESARE file:", TSDApp.DOC8Path)))
#     colList.append(list(("Customer effects file:", TSDApp.DOC7Name)))
#     colList.append(list(("Diversity management file:", TSDApp.DOC13Path)))
#     colList.append(list(("CESARE file reference:", TSDApp.DOC8Link.split("/")[-3])))
#     colList.append(list(("Criticity configuration file reference:", TSDApp.DOC9Link.split("/")[-3])))
#     colList.append(list(("Customer effect file reference:", TSDApp.DOC7Link.split("/")[-3])))
#     colList.append(list(("Diversity management file reference:", TSDApp.DOC13Link.split("/")[-3])))
#     colList.append(list(("Check level:", TSDApp.checkLevel)))
#     colList.append(list(("","")))
#     colList.append(list(("Date of the test:", time.strftime("%x"))))
#     colList.append(list(("Time of the test:", time.strftime("%X"))))
#     colList.append(list(("Test duration:", time.strftime('%H:%M:%S', time.gmtime(TSDApp.end_time - TSDApp.start_time)))))
#     colList.append(list(("Opening duration:", time.strftime('%H:%M:%S', time.gmtime(TSDApp.opening_time - TSDApp.start_time)) )))
#     colList.append(list(("","")))
#     colList.append(list(("TSD file checked:", TSDApp.DOC3Path)))
#     colList.append(list(("TSD function file checked:", TSDApp.DOC4Path)))
#     colList.append(list(("TSD system file checked:", TSDApp.DOC5Path)))
#     colList.append(list(("","")))
#     colList.append(list(("AMDEC:",TSDApp.AMDECName)))
#     colList.append(list(("Export MedialecMatrice:",TSDApp.MedialecName)))
#     colList.append(list(("","")))
#     colList.append(list(("Status:", str(TSDApp.status))))
#     colList.append(list(("Coverage Indicator:", str(TSDApp.coverage)[0:4] + "%")))
#     colList.append(list(("Convergence Indicator:", str(TSDApp.convergence)[0:4] + "%")))
#     colList.append(list(("", "")))
#     colList.append(list(("Blocking Points", str(TSDApp.criticity_blocking))))
#     colList.append(list(("Warning Points", str(TSDApp.criticity_warning))))
#     colList.append(list(("Information Points", str(TSDApp.criticity_information))))
#     reportInformationWorkSheet.Range("A1:B26").Value = colList
#     for column in reportInformationWorkSheet.Range("A1:B26").Columns:
#         column.AutoFit()
