import TSD_Checker_V4_0
import inspect
import win32com.client as win32
from ExcelEdit import TestReturn as result
from ErrorMessages import errorMessagesDict as error
import xlrd
import xlwt
from xlutils.copy import copy

def coverageIndicator(workBook, TSDApp):
    index = 0
    for sheetname in TSDApp.WorkbookStats.sheetNames:
        if sheetname == 'tableau':
            index = TSDApp.WorkbookStats.sheetNames.index('tableau')
            break
        if sheetname == 'table':
            index = TSDApp.WorkbookStats.sheetNames.index('table')
            break

    workSheet = workBook.sheet_by_index(index)
    nrCols = workSheet.ncols
    nrRows = workSheet.nrows

    refColBase = -1
    refColDTC = -1
    refCelParam = -1
    refCelDiag = -1

    for index in range(0, TSDApp.WorkbookStats.tableLastCol):
        if str(workSheet.cell(TSDApp.tableHeaderRow,index).value).casefold().strip() == "Constituant défaillant détecté".casefold():
            refColBase = index
        if str(workSheet.cell(TSDApp.tableHeaderRow,index).value).casefold().strip() == "Code défaut".casefold():
            refColDTC = index
        if str(workSheet.cell(TSDApp.tableHeaderRow,index).value).casefold().strip() == "mesures et commandes (Mesure Parametre et Test Actionneur) / Tests de cohérence".casefold():
            refCelParam = index
        if str(workSheet.cell(TSDApp.tableHeaderRow,index).value).casefold().strip() == "DIAGNOSTIC DEBARQUE".casefold():
            refCelDiag = index


    NbComponentsOfTheFunction = 0
    NbComponentWithDiagPossible = 0
    for index in range(TSDApp.tableFirstInfoRow, nrRows):
        if workSheet.cell(index, refColBase).value is not None and workSheet.cell(index,refColBase).value != "":
            NbComponentsOfTheFunction += 1
            if (workSheet.cell(index, refColDTC).value is not None and workSheet.cell(index, refColDTC).value !="" and workSheet.cell(index,refColDTC).value != "NO DTC") or (
                    workSheet.cell(index, refCelParam).value is not None and workSheet.cell(index, refCelParam).value != "" and workSheet.cell(index,refCelParam).value != "N/A") or (
                    workSheet.cell(index, refCelDiag).value is not None and workSheet.cell(index, refCelDiag).value != "" and workSheet.cell(index,refCelDiag).value != "N/A"):
                NbComponentWithDiagPossible += 1

    return (NbComponentWithDiagPossible / NbComponentsOfTheFunction)


    # index = 0
    # for sheetname in TSDApp.WorkbookStats.sheetNames:
    #     if sheetname == 'tableau':
    #         index = TSDApp.WorkbookStats.sheetNames.index('tableau') + 1
    #         break
    #     if sheetname == 'table':
    #         index = TSDApp.WorkbookStats.sheetNames.index('table') + 1
    #         break
    #
    # workSheet = workBook.Sheets(index)
    # workSheetRange = workSheet.UsedRange
    # nrCols = workSheetRange.Columns.Count
    # nrRows = workSheetRange.Rows.Count
    # refColBase = 0
    # refColDTC = 0
    # refCelParam = 0
    # refCelDiag = 0
    #
    # for cellRow in workSheetRange.value:
    #     for cell in cellRow:
    #         if cell == "Constituant défaillant détecté":
    #             refColBase = cellRow.index(cell) + 1
    #             refRowIndex = workSheetRange.value.index(cellRow) + 1
    #         if cell == "Code défaut":
    #             refColDTC = cellRow.index(cell) + 1
    #             refRowIndex = workSheetRange.value.index(cellRow) + 1
    #         if cell == "mesures et commandes (Mesure Parametre et Test Actionneur) / Tests de cohérence":
    #             refCelParam = cellRow.index(cell) + 1
    #             refRowIndex = workSheetRange.value.index(cellRow) + 1
    #         if cell == "DIAGNOSTIC DEBARQUE":
    #             refCelDiag = cellRow.index(cell) + 1
    #             refRowIndex = workSheetRange.value.index(cellRow) + 1
    #
    #     if refColBase != 0:
    #         break
    #
    # refCellRange = workSheet.cell(refRowIndex, refColBase).MergeArea
    # nrLines = refCellRange.Rows.Count
    #
    # NbComponentsOfTheFunction = 0
    # NbComponentWithDiagPossible = 0
    # for index in range(refRowIndex + nrLines, nrRows):
    #     if workSheet.cell(index, refColBase).value is not None:
    #         NbComponentsOfTheFunction += 1
    #         if (workSheet.cell(index, refColDTC).value is not None and workSheet.cell(index, refColDTC).value != "NO DTC") or (workSheet.cell(index, refCelParam).value is not None and workSheet.cell(index, refCelParam).value != "N/A") or (workSheet.cell(index, refCelDiag).value is not None and workSheet.cell(index, refCelDiag).value != "N/A"):
    #             NbComponentWithDiagPossible += 1
    #
    # return (NbComponentWithDiagPossible / NbComponentsOfTheFunction)

def convergenceIndicator(workBook, TSDApp):

    index = -1
    for sheetname in TSDApp.WorkbookStats.sheetNames:
        if sheetname == 'tableau':
            index = TSDApp.WorkbookStats.sheetNames.index('tableau')
            break
        if sheetname == 'table':
            index = TSDApp.WorkbookStats.sheetNames.index('table')
            break

    rb_sheet = workBook.sheet_by_index(index)
    nrCols = rb_sheet.ncols
    nrRows = rb_sheet.nrows

    refColBase = -1
    refColDTC = -1
    refCelParam = -1
    refCelDiag = -1

    refSignature = -1
    refCritere = -1

    for index1 in range(0, nrRows):
        for index2 in range(0, nrCols):
            if str(rb_sheet.cell(index1, index2).value).casefold().strip() == "Critère de decision".casefold():
                refCritere = index2
                refRowIndex = index1
            if str(rb_sheet.cell(index1, index2).value).casefold().strip() == "Unique Test Signature".casefold():
                refSignature = index2
                refRowIndex = index1
        if refCritere != -1 and refSignature != -1:
            break

    for index1 in range(0, nrRows):
        for index2 in range(0, nrCols):
            if str(rb_sheet.cell(index1, index2).value).casefold().strip() == "Constituant défaillant détecté".casefold():
                refColBase = index2
                refRowIndex = index1
            if str(rb_sheet.cell(index1, index2).value).casefold().strip() == "Code défaut".casefold():
                refColDTC = index2
                refRowIndex = index1
            if str(rb_sheet.cell(index1, index2).value).casefold().strip() == "mesures et commandes (Mesure Parametre et Test Actionneur) / Tests de cohérence".casefold():
                refCelParam = index2
                refRowIndex = index1
            if str(rb_sheet.cell(index1, index2).value).casefold().strip() == "DIAGNOSTIC DEBARQUE".casefold():
                refCelDiag = index2
                refRowIndex = index1

    workBook2 = copy(workBook)
    workSheet = workBook2.get_sheet(index)


    # if refSignature == -1:
        # add column in position

    NbUniqueSignatureTests = 0
    NbAMDECLine = 0
    unique_items = []

    for index in range(TSDApp.tableFirstInfoRow, nrRows):
        if rb_sheet.cell(index, refColBase).value != "":
            NbAMDECLine += 1
            if [rb_sheet.cell(index, refColDTC).value, rb_sheet.cell(index, refCelParam).value,rb_sheet.cell(index, refCelDiag).value] not in unique_items:
                unique_items.append([rb_sheet.cell(index, refColDTC).value, rb_sheet.cell(index, refCelParam).value, rb_sheet.cell(index, refCelDiag).value])
                # workSheet.write(index, refSignature, '1')
                NbUniqueSignatureTests += 1
            else:
                # workSheet.write(index, refSignature, '0')
                pass


    # if TSDApp.DOC3Path.split('.')[-1] == "xls":
    #     workBook2.save(TSDApp.DOC3Path)
    # else:
    #     path = TSDApp.DOC3Path.split('.')[0]
    #     path2 = TSDApp.DOC3Path.split('/')[0]
    #     for i in range(1, len(TSDApp.DOC3Path.split('.')) - 1):
    #         path = path + "." + TSDApp.DOC3Path.split('.')[i]
    #     for i in range(1, len(TSDApp.DOC3Path.split('/')) - 1):
    #         path2 = path2 + "/" + TSDApp.DOC3Path.split('/')[i]
    #     path = path + ".xlsx"
    #     path2 = path2 + "/Temporal.xls"
    #
    #     workBook2.save(path2)

    # p.save_book_as(file_name=path2, dest_file_name=path)
    print("a")
    return (NbUniqueSignatureTests / NbAMDECLine)



    # index = 0
    # for sheetname in TSDApp.WorkbookStats.sheetNames:
    #     if sheetname == 'tableau':
    #         index = TSDApp.WorkbookStats.sheetNames.index('tableau') + 1
    #         break
    #     if sheetname == 'table':
    #         index = TSDApp.WorkbookStats.sheetNames.index('table') + 1
    #         break
    #
    # workSheet = workBook.Sheets(index)
    # workSheetRange = workSheet.UsedRange
    # nrCols = workSheetRange.Columns.Count
    # nrRows = workSheetRange.Rows.Count
    # refColBase = 0
    # refColDTC = 0
    # refCelParam = 0
    # refCelDiag = 0
    #
    # refSignature = 0
    # refCritere = 0
    # for cellRow in workSheetRange.value:
    #     for cell in cellRow:
    #         if cell == "Critère de decision":
    #             refCritere = cellRow.index(cell) + 1
    #             refRowIndex = workSheetRange.value.index(cellRow) + 1
    #         if cell == "Unique Test Signature":
    #             refSignature = cellRow.index(cell) + 1
    #             refRowIndex = workSheetRange.value.index(cellRow) + 1
    #
    # if refSignature != 0:
    #     workSheet.cell(refRowIndex, refSignature).EntireColumn.Delete(win32.constants.xlShiftToLeft)
    #     workSheet.cell(refRowIndex, refSignature).EntireColumn.Insert(win32.constants.xlShiftToLeft)
    #     workSheet.cell(refRowIndex, refSignature).value = "Unique Test Signature"
    # else:
    #     workSheet.cell(refRowIndex, refCritere + 1).EntireColumn.Insert(win32.constants.xlShiftToLeft)
    #     workSheet.cell(refRowIndex, refCritere + 1).value = "Unique Test Signature"
    #     refSignature = refCritere + 1
    #
    # for cellRow in workSheetRange.value:
    #     for cell in cellRow:
    #         if cell == "Constituant défaillant détecté":
    #             refColBase = cellRow.index(cell) + 1
    #             refRowIndex = workSheetRange.value.index(cellRow) + 1
    #         if cell == "Code défaut":
    #             refColDTC = cellRow.index(cell) + 1
    #             refRowIndex = workSheetRange.value.index(cellRow) + 1
    #         if cell == "mesures et commandes (Mesure Parametre et Test Actionneur) / Tests de cohérence":
    #             refCelParam = cellRow.index(cell) + 1
    #             refRowIndex = workSheetRange.value.index(cellRow) + 1
    #         if cell == "DIAGNOSTIC DEBARQUE":
    #             refCelDiag = cellRow.index(cell) + 1
    #             refRowIndex = workSheetRange.value.index(cellRow) + 1
    #
    #     if refColBase != 0:
    #         break
    #
    # refCellRange = workSheet.cell(refRowIndex, refColBase).MergeArea
    # nrLines = refCellRange.Rows.Count
    #
    # NbUniqueSignatureTests = 0
    # NbAMDECLine = 0
    # unique_items = []
    # for index in range(refRowIndex + nrLines, nrRows):
    #     if workSheet.cell(index, refColBase).value is not None:
    #         NbAMDECLine += 1
    #         if [workSheet.cell(index, refColDTC).value, workSheet.cell(index, refCelParam).value, workSheet.cell(index, refCelDiag).value] not in unique_items:
    #             unique_items.append([workSheet.cell(index, refColDTC).value, workSheet.cell(index, refCelParam).value, workSheet.cell(index, refCelDiag).value])
    #             workSheet.cell(index, refSignature).value = "1"
    #             NbUniqueSignatureTests += 1
    #         else:
    #             workSheet.cell(index, refSignature).value = "0"
    #
    # return (NbUniqueSignatureTests / NbAMDECLine)