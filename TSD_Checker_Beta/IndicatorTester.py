import TSD_Checker_V0_5_2
import inspect
import win32com.client as win32
from ExcelEdit import TestReturn as result
from ErrorMessages import errorMessagesDict as error

def coverageIndicator(workBook, TSDApp):
    index = 0
    for sheetname in TSDApp.WorkbookStats.sheetNames:
        if sheetname == 'tableau':
            index = TSDApp.WorkbookStats.sheetNames.index('tableau') + 1
            break
        if sheetname == 'table':
            index = TSDApp.WorkbookStats.sheetNames.index('table') + 1
            break

    workSheet = workBook.Sheets(index)
    workSheetRange = workSheet.UsedRange
    nrCols = workSheetRange.Columns.Count
    nrRows = workSheetRange.Rows.Count
    refColBase = 0
    refColDTC = 0
    refCelParam = 0
    refCelDiag = 0

    for cellRow in workSheetRange.Value:
        for cell in cellRow:
            if cell == "Constituant défaillant détecté":
                refColBase = cellRow.index(cell) + 1
                refRowIndex = workSheetRange.Value.index(cellRow) + 1
            if cell == "Code défaut":
                refColDTC = cellRow.index(cell) + 1
                refRowIndex = workSheetRange.Value.index(cellRow) + 1
            if cell == "mesures et commandes (Mesure Parametre et Test Actionneur) / Tests de cohérence":
                refCelParam = cellRow.index(cell) + 1
                refRowIndex = workSheetRange.Value.index(cellRow) + 1
            if cell == "DIAGNOSTIC DEBARQUE":
                refCelDiag = cellRow.index(cell) + 1
                refRowIndex = workSheetRange.Value.index(cellRow) + 1

        if refColBase != 0:
            break

    refCellRange = workSheet.Cells(refRowIndex, refColBase).MergeArea
    nrLines = refCellRange.Rows.Count

    NbComponentsOfTheFunction = 0
    NbComponentWithDiagPossible = 0
    for index in range(refRowIndex + nrLines, nrRows):
        if workSheet.Cells(index, refColBase).Value is not None:
            NbComponentsOfTheFunction += 1
            if (workSheet.Cells(index, refColDTC).Value is not None and workSheet.Cells(index, refColDTC).Value != "NO DTC") or (workSheet.Cells(index, refCelParam).Value is not None and workSheet.Cells(index, refCelParam).Value != "N/A") or (workSheet.Cells(index, refCelDiag).Value is not None and workSheet.Cells(index, refCelDiag).Value != "N/A"):
                NbComponentWithDiagPossible += 1

    return (NbComponentWithDiagPossible / NbComponentsOfTheFunction)

def convergenceIndicator(workBook, TSDApp):
    index = 0
    for sheetname in TSDApp.WorkbookStats.sheetNames:
        if sheetname == 'tableau':
            index = TSDApp.WorkbookStats.sheetNames.index('tableau') + 1
            break
        if sheetname == 'table':
            index = TSDApp.WorkbookStats.sheetNames.index('table') + 1
            break

    workSheet = workBook.Sheets(index)
    workSheetRange = workSheet.UsedRange
    nrCols = workSheetRange.Columns.Count
    nrRows = workSheetRange.Rows.Count
    refColBase = 0
    refColDTC = 0
    refCelParam = 0
    refCelDiag = 0

    refSignature = 0
    refCritere = 0
    for cellRow in workSheetRange.Value:
        for cell in cellRow:
            if cell == "Critère de decision":
                refCritere = cellRow.index(cell) + 1
            if cell == "Unique Test Signature":
                refSignature = cellRow.index(cell) + 1

    if refSignature != 0:
        workSheetRange.Columns(refSignature).EntireColumn.Delete
        workSheetRange.Range(refSignature).EntireColumn.Insert
    else:
        workSheetRange.Range(refCritere + 1).EntireColumn.Insert

    for cellRow in workSheetRange.Value:
        for cell in cellRow:
            if cell == "Constituant défaillant détecté":
                refColBase = cellRow.index(cell) + 1
                refRowIndex = workSheetRange.Value.index(cellRow) + 1
            if cell == "Code défaut":
                refColDTC = cellRow.index(cell) + 1
                refRowIndex = workSheetRange.Value.index(cellRow) + 1
            if cell == "mesures et commandes (Mesure Parametre et Test Actionneur) / Tests de cohérence":
                refCelParam = cellRow.index(cell) + 1
                refRowIndex = workSheetRange.Value.index(cellRow) + 1
            if cell == "DIAGNOSTIC DEBARQUE":
                refCelDiag = cellRow.index(cell) + 1
                refRowIndex = workSheetRange.Value.index(cellRow) + 1

        if refColBase != 0:
            break

    refCellRange = workSheet.Cells(refRowIndex, refColBase).MergeArea
    nrLines = refCellRange.Rows.Count

    NbUniqueSignatureTests = 0
    NbAMDECLine = 0
    unique_items = []
    for index in range(refRowIndex + nrLines, nrRows):
        if workSheet.Cells(index, refColBase).Value is not None:
            NbAMDECLine += 1
            if [workSheet.Cells(index, refColDTC).Value, workSheet.Cells(index, refCelParam).Value, workSheet.Cells(index, refCelDiag).Value] not in unique_items:
                unique_items.append([workSheet.Cells(index, refColDTC).Value, workSheet.Cells(index, refCelParam).Value, workSheet.Cells(index, refCelDiag).Value])
                workSheet.Cells(index, refColBase).Value = "1"
                NbUniqueSignatureTests += 1
            else:
                workSheet.Cells(index, refColBase).Value = "0"

    return (NbUniqueSignatureTests / NbAMDECLine)