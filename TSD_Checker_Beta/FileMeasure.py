import TSD_Checker_V3_1
import inspect



def getTableInfo(workBook, TSDApp):
    temp = workBook.Sheets
    sheetNames = []
    for sheet in temp:
        sheetNames.append(sheet.Name.strip().casefold())
    TSDApp.WorkbookStats.sheetNames = sheetNames
    if "tableau" in sheetNames or "table" in sheetNames:
        TSDApp.WorkbookStats.hasTable = True
        try:
            index = sheetNames.index("tableau") + 1
        except:
            index = sheetNames.index("table") + 1
        TSDApp.WorkbookStats.tableIndex = index
    else:
        TSDApp.WorkbookStats.hasTable = False

    if TSDApp.WorkbookStats.hasTable == True:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.tableIndex)
        var = 0
        ok = 0
        col_range = 0
        lastCol = 0
        tmp = 0
        ExitFromFct = 0
        TSDApp.WorkbookStats.tableLastRow = 0
        lastFilledCell = 0

        for cellRow in workSheet.Rows:
            col_range = 0
            if ExitFromFct == 1:
                break
            for cell in cellRow.Cells:
                if tmp != 0:
                    ok = 1
                    if col_range == 0:
                        if cell.Borders(9).LineStyle != -4142:
                            if cell.Value is not None:
                                lastFilledCell = cell.Row
                        else:
                            TSDApp.WorkbookStats.tableLastRow = cell.Row
                            tmp = 0
                            break
                    else:
                        break
                elif TSDApp.WorkbookStats.tableLastRow != 0:
                    ExitFromFct = 1
                    break
                if ok == 0:
                    if str(cell.Value).casefold() == "Référence".casefold().strip() or str(cell.Value).casefold().strip() == "Reference".casefold():
                        TSDApp.WorkbookStats.tableRefColIndex = cell.Column
                        TSDApp.WorkbookStats.tableRefRowIndex = cell.Row
                        indexCol = 1
                        col_range = 1
                    if col_range == 1:
                        if cell.Borders(8).LineStyle != -4142 and cell != None:
                            indexCol += 1
                            pass
                        else:
                            lastCol = cell.Column
                            tmp = 1
                            ok = 1
                            break
                else:
                    break

        if TSDApp.WorkbookStats.tableRefColIndex == 0:
            var = 1

        if var == 0:
            TSDApp.WorkbookStats.tableLastRow = lastFilledCell
            TSDApp.WorkbookStats.tableLastCol = lastCol

        else:
            TSDApp.WorkbookStats.tableLastRow = None
            TSDApp.WorkbookStats.tableLastCol = None
    else:
        TSDApp.WorkbookStats.tableLastRow = None
        TSDApp.WorkbookStats.tableLastCol = None


def getCodesDefautsInfo(workBook, TSDApp):
    temp = workBook.Sheets
    sheetNames = []
    for sheet in temp:
        sheetNames.append(sheet.Name.strip().casefold())
    TSDApp.WorkbookStats.sheetNames = sheetNames
    if "codes défauts" in sheetNames or "data trouble codes" in sheetNames:
        TSDApp.WorkbookStats.hasCode = True
        try:
            index = sheetNames.index("codes défauts") + 1
        except:
            index = sheetNames.index("data trouble codes") + 1
        TSDApp.WorkbookStats.codeIndex = index
    else:
        TSDApp.WorkbookStats.hasCode = False

    if TSDApp.WorkbookStats.hasCode == True:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.codeIndex)
        var = 0
        ok = 0
        col_range = 0
        lastCol = 0
        tmp = 0
        ExitFromFct = 0
        TSDApp.WorkbookStats.codeLastRow = 0
        lastFilledCell = 0

        for cellRow in workSheet.Rows:
            col_range = 0
            if ExitFromFct == 1:
                break
            for cell in cellRow.Cells:
                if tmp != 0:
                    ok = 1
                    if col_range == 0:
                        if cell.Borders(9).LineStyle != -4142:
                            if cell.Value is not None:
                                lastFilledCell = cell.Row
                        else:
                            TSDApp.WorkbookStats.codeLastRow = cell.Row
                            tmp = 0
                            break
                    else:
                        break
                elif TSDApp.WorkbookStats.codeLastRow != 0:
                    ExitFromFct = 1
                    break
                if ok == 0:
                    if str(cell.Value).casefold() == "Référence".casefold().strip() or str(
                            cell.Value).casefold().strip() == "Reference".casefold():
                        TSDApp.WorkbookStats.codeRefColIndex = cell.Column
                        TSDApp.WorkbookStats.codeRefRowIndex = cell.Row
                        indexCol = 1
                        col_range = 1
                    if col_range == 1:
                        if cell.Borders(8).LineStyle != -4142 and cell != None:
                            indexCol += 1
                            pass
                        else:
                            lastCol = cell.Column
                            tmp = 1
                            ok = 1
                            break
                else:
                    break

        if TSDApp.WorkbookStats.codeRefColIndex == 0:
            var = 1

        if var == 0:
            TSDApp.WorkbookStats.codeLastRow = lastFilledCell
            TSDApp.WorkbookStats.codeLastCol = lastCol

        else:
            TSDApp.WorkbookStats.codeLastRow = None
            TSDApp.WorkbookStats.codeLastCol = None
    else:
        TSDApp.WorkbookStats.codeLastRow = None
        TSDApp.WorkbookStats.codeLastCol = None


def getMesuresEtCommandesInfo(workBook, TSDApp):
    temp = workBook.Sheets
    sheetNames = []
    for sheet in temp:
        sheetNames.append(sheet.Name.strip().casefold())
    TSDApp.WorkbookStats.sheetNames = sheetNames
    if "mesures et commandes" in sheetNames:
        TSDApp.WorkbookStats.hasMeasure = True
        try:
            index = sheetNames.index("mesures et commandes") + 1
        except:
            pass
        TSDApp.WorkbookStats.measureIndex = index
    else:
        TSDApp.WorkbookStats.hasMeasure = False

    if TSDApp.WorkbookStats.hasMeasure == True:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.measureIndex)
        var = 0
        ok = 0
        col_range = 0
        lastCol = 0
        tmp = 0
        ExitFromFct = 0
        TSDApp.WorkbookStats.measureLastRow = 0
        lastFilledCell = 0

        for cellRow in workSheet.Rows:
            col_range = 0
            if ExitFromFct == 1:
                break
            for cell in cellRow.Cells:
                if tmp != 0:
                    ok = 1
                    if col_range == 0:
                        if cell.Borders(9).LineStyle != -4142:
                            if cell.Value is not None:
                                lastFilledCell = cell.Row
                        else:
                            TSDApp.WorkbookStats.measureLastRow = cell.Row
                            tmp = 0
                            break
                    else:
                        break
                elif TSDApp.WorkbookStats.measureLastRow != 0:
                    ExitFromFct = 1
                    break
                if ok == 0:
                    if str(cell.Value).casefold() == "Référence".casefold().strip() or str(
                            cell.Value).casefold().strip() == "Reference".casefold():
                        TSDApp.WorkbookStats.measureRefColIndex = cell.Column
                        TSDApp.WorkbookStats.measureRefRowIndex = cell.Row
                        indexCol = 1
                        col_range = 1
                    if col_range == 1:
                        if cell.Borders(8).LineStyle != -4142 and cell != None:
                            indexCol += 1
                            pass
                        else:
                            lastCol = cell.Column
                            tmp = 1
                            ok = 1
                            break
                else:
                    break

        if TSDApp.WorkbookStats.measureRefColIndex == 0:
            var = 1

        if var == 0:
            TSDApp.WorkbookStats.measureLastRow = lastFilledCell
            TSDApp.WorkbookStats.measureLastCol = lastCol

        else:
            TSDApp.WorkbookStats.measureLastRow = None
            TSDApp.WorkbookStats.measureLastCol = None
    else:
        TSDApp.WorkbookStats.measureLastRow = None
        TSDApp.WorkbookStats.measureLastCol = None


def getDiagnosticDebarquesInfo(workBook, TSDApp):
    temp = workBook.Sheets
    sheetNames = []
    for sheet in temp:
        sheetNames.append(sheet.Name.strip().casefold())
    TSDApp.WorkbookStats.sheetNames = sheetNames
    if "diagnostic débarqués" in sheetNames:
        TSDApp.WorkbookStats.hasDiagDeb = True
        try:
            index = sheetNames.index("diagnostic débarqués") + 1
        except:
            pass
        TSDApp.WorkbookStats.DiagDebIndex = index
    else:
        TSDApp.WorkbookStats.hasDiagDeb = False

    if TSDApp.WorkbookStats.hasDiagDeb == True:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagDebIndex)
        var = 0
        ok = 0
        col_range = 0
        lastCol = 0
        tmp = 0
        ExitFromFct = 0
        TSDApp.WorkbookStats.DiagDebLastRow = 0
        lastFilledCell = 0

        for cellRow in workSheet.Rows:
            col_range = 0
            if ExitFromFct == 1:
                break
            for cell in cellRow.Cells:
                if tmp != 0:
                    ok = 1
                    if col_range == 0:
                        if cell.Borders(9).LineStyle != -4142:
                            if cell.Value is not None:
                                lastFilledCell = cell.Row
                        else:
                            TSDApp.WorkbookStats.DiagDebLastRow = cell.Row
                            tmp = 0
                            break
                    else:
                        break
                elif TSDApp.WorkbookStats.DiagDebLastRow != 0:
                    ExitFromFct = 1
                    break
                if ok == 0:
                    if str(cell.Value).casefold() == "Référence".casefold().strip() or str(cell.Value).casefold().strip() == "Reference".casefold():
                        TSDApp.WorkbookStats.DiagDebRefColIndex = cell.Column
                        TSDApp.WorkbookStats.DiagDebRefRowIndex = cell.Row
                        indexCol = 1
                        col_range = 1
                    if col_range == 1:
                        if cell.Borders(8).LineStyle != -4142 and cell != None:
                            indexCol += 1
                            pass
                        else:
                            lastCol = cell.Column
                            tmp = 1
                            ok = 1
                            break
                else:
                    break

        if TSDApp.WorkbookStats.DiagDebRefColIndex == 0:
            var = 1

        if var == 0:
            TSDApp.WorkbookStats.DiagDebLastRow = lastFilledCell
            TSDApp.WorkbookStats.DiagDebLastCol = lastCol

        else:
            TSDApp.WorkbookStats.DiagDebLastRow = None
            TSDApp.WorkbookStats.DiagDebLastCol = None
    else:
        TSDApp.WorkbookStats.DiagDebLastRow = None
        TSDApp.WorkbookStats.DiagDebLastCol = None


def getListeMDDInfo(workBook, TSDApp):
    temp = workBook.Sheets
    sheetNames = []
    for sheet in temp:
        sheetNames.append(sheet.Name.strip().casefold())
    TSDApp.WorkbookStats.sheetNames = sheetNames
    if "degraded mode" in sheetNames or "liste mdd" in sheetNames:
        TSDApp.WorkbookStats.hasMDD = True
        try:
            index = sheetNames.index("degraded mode") + 1
        except:
            index = sheetNames.index("liste mdd") + 1
        TSDApp.WorkbookStats.MDDIndex = index
    else:
        TSDApp.WorkbookStats.hasMDD = False

    if TSDApp.WorkbookStats.hasMDD == True:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.MDDIndex)
        var = 0
        ok = 0
        col_range = 0
        lastCol = 0
        tmp = 0
        ExitFromFct = 0
        TSDApp.WorkbookStats.MDDLastRow = 0
        lastFilledCell = 0

        for cellRow in workSheet.Rows:
            col_range = 0
            if ExitFromFct == 1:
                break
            for cell in cellRow.Cells:
                if tmp != 0:
                    ok = 1
                    if col_range == 0:
                        if cell.Borders(9).LineStyle != -4142:
                            if cell.Value is not None:
                                lastFilledCell = cell.Row
                        else:
                            TSDApp.WorkbookStats.MDDLastRow = cell.Row
                            tmp = 0
                            break
                    else:
                        break
                elif TSDApp.WorkbookStats.MDDLastRow != 0:
                    ExitFromFct = 1
                    break
                if ok == 0:
                    if str(cell.Value).casefold().strip() == "Modes dégradés:".casefold() or str(cell.Value).casefold().strip() == "N°".casefold():
                        TSDApp.WorkbookStats.MDDRefColIndex = cell.Column
                        TSDApp.WorkbookStats.MDDRefRowIndex = cell.Row
                        indexCol = 1
                        col_range = 1
                    if col_range == 1:
                        if cell.Borders(8).LineStyle != -4142 and cell != None:
                            indexCol += 1
                            pass
                        else:
                            lastCol = cell.Column
                            tmp = 1
                            ok = 1
                            break
                else:
                    break

        if TSDApp.WorkbookStats.MDDRefColIndex == 0:
            var = 1

        if var == 0:
            TSDApp.WorkbookStats.MDDLastRow = lastFilledCell
            TSDApp.WorkbookStats.MDDLastCol = lastCol


        else:
            TSDApp.WorkbookStats.MDDLastRow = None
            TSDApp.WorkbookStats.MDDLastCol = None
    else:
        TSDApp.WorkbookStats.MDDLastRow = None
        TSDApp.WorkbookStats.MDDLastCol = None


def getEffetsClientsInfo(workBook, TSDApp):
    temp = workBook.Sheets
    sheetNames = []
    for sheet in temp:
        sheetNames.append(sheet.Name.strip().casefold())
    TSDApp.WorkbookStats.sheetNames = sheetNames
    if "effets clients" in sheetNames or "customer effects" in sheetNames or "customer effect" in sheetNames:
        TSDApp.WorkbookStats.hasEffClients = True
        try:
            index = sheetNames.index("effets clients") + 1
        except:
            if "customer effect" in sheetNames:
                index = sheetNames.index("customer effect") + 1
            else:
                index = sheetNames.index("customer effects") + 1
        TSDApp.WorkbookStats.EffClientsIndex = index
    else:
        TSDApp.WorkbookStats.hasEffClients = False

    if TSDApp.WorkbookStats.hasEffClients == True:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.EffClientsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0
        ok = 0
        col_range = 0
        lastCol = 0
        tmp = 0
        ExitFromFct = 0
        TSDApp.WorkbookStats.EffClientsLastRow = 0
        lastFilledCell = 0

        for cellRow in workSheet.Rows:
            col_range = 0
            if ExitFromFct == 1:
                break
            for cell in cellRow.Cells:
                if tmp != 0:
                    ok = 1
                    if col_range == 0:
                        if cell.Borders(9).LineStyle != -4142:
                            if cell.Value is not None:
                                lastFilledCell = cell.Row
                        else:
                            TSDApp.WorkbookStats.EffClientsLastRow = cell.Row
                            tmp = 0
                            break
                    else:
                        break
                elif TSDApp.WorkbookStats.EffClientsLastRow != 0:
                    ExitFromFct = 1
                    break
                if ok == 0:
                    if str(cell.Value).casefold().strip() == "Noms".casefold() or str(cell.Value).casefold().strip() == "Name".casefold():
                        refColIndex = cell.Column
                        refRowIndex = cell.Row
                        indexCol = 1
                        col_range = 1
                    if col_range == 1:
                        if cell.Borders(8).LineStyle != -4142 and cell != None:
                            indexCol += 1
                            pass
                        else:
                            lastCol = cell.Column
                            tmp = 1
                            ok = 1
                            break
                else:
                    break

        if refColIndex == 0:
            var = 1

        if var == 0:
            TSDApp.WorkbookStats.EffClientsLastRow = lastFilledCell
            TSDApp.WorkbookStats.EffClientsLastCol = lastCol
            a = 3

        else:
            TSDApp.WorkbookStats.EffClientsLastRow = None
            TSDApp.WorkbookStats.EffClientsLastCol = None
    else:
        TSDApp.WorkbookStats.EffClientsLastRow = None
        TSDApp.WorkbookStats.EffClientsLastCol = None

def getReqOfTechEffectsInfo(workBook, TSDApp):
    temp = workBook.Sheets
    sheetNames = []
    for sheet in temp:
        sheetNames.append(sheet.Name.strip().casefold())
    TSDApp.WorkbookStats.sheetNames = sheetNames
    if "req. of tech. effects" in sheetNames:
        TSDApp.WorkbookStats.hasReqTech = True
        try:
            index = sheetNames.index("req. of tech. effects") + 1
        except:
            pass
        TSDApp.WorkbookStats.ReqTechIndex = index
    else:
        TSDApp.WorkbookStats.hasReqTech = False

    if TSDApp.WorkbookStats.hasReqTech == True:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReqTechIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0
        ok = 0
        col_range = 0
        lastCol = 0
        tmp = 0
        ExitFromFct = 0
        TSDApp.WorkbookStats.ReqTechLastRow = 0
        lastFilledCell = 0

        for cellRow in workSheet.Rows:
            col_range = 0
            if ExitFromFct == 1:
                break
            for cell in cellRow.Cells:
                if tmp != 0:
                    ok = 1
                    if col_range == 0:
                        if cell.Borders(9).LineStyle != -4142:
                            if cell.Value is not None:
                                lastFilledCell = cell.Row
                        else:
                            TSDApp.WorkbookStats.ReqTechLastRow = cell.Row
                            tmp = 0
                            break
                    else:
                        break
                elif TSDApp.WorkbookStats.ReqTechLastRow != 0:
                    ExitFromFct = 1
                    break
                if ok == 0:
                    if str(cell.Value).casefold() == "Référence".casefold().strip() or str(cell.Value).casefold().strip() == "Reference".casefold():
                        refColIndex = cell.Column
                        refRowIndex = cell.Row
                        indexCol = 1
                        col_range = 1
                    if col_range == 1:
                        if cell.Borders(8).LineStyle != -4142 and cell != None:
                            indexCol += 1
                            pass
                        else:
                            lastCol = cell.Column
                            tmp = 1
                            ok = 1
                            break
                else:
                    break

        if refColIndex == 0:
            var = 1

        if var == 0:
            TSDApp.WorkbookStats.ReqTechLastRow = lastFilledCell
            TSDApp.WorkbookStats.ReqTechLastCol = lastCol

        else:
            TSDApp.WorkbookStats.ReqTechLastRow = None
            TSDApp.WorkbookStats.ReqTechLastCol = None
    else:
        TSDApp.WorkbookStats.ReqTechLastRow = None
        TSDApp.WorkbookStats.ReqTechLastCol = None


def getDataTroubleCodesInfo(workBook, TSDApp):
    temp = workBook.Sheets
    sheetNames = []
    for sheet in temp:
        sheetNames.append(sheet.Name.strip().casefold())
    TSDApp.WorkbookStats.sheetNames = sheetNames
    if "data trouble codes" in sheetNames:
        TSDApp.WorkbookStats.hasDataCodes = True
        try:
            index = sheetNames.index("data trouble codes") + 1
        except:
            pass
        TSDApp.WorkbookStats.DataCodesIndex = index
    else:
        TSDApp.WorkbookStats.hasDataCodes = False

    if TSDApp.WorkbookStats.hasDataCodes == True:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DataCodesIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0
        ok = 0
        col_range = 0
        lastCol = 0
        tmp = 0
        ExitFromFct = 0
        TSDApp.WorkbookStats.DataCodesLastRow = 0
        lastFilledCell = 0

        for cellRow in workSheet.Rows:
            col_range = 0
            if ExitFromFct == 1:
                break
            for cell in cellRow.Cells:
                if tmp != 0:
                    ok = 1
                    if col_range == 0:
                        if cell.Borders(9).LineStyle != -4142:
                            if cell.Value is not None:
                                lastFilledCell = cell.Row
                        else:
                            TSDApp.WorkbookStats.DataCodesLastRow = cell.Row
                            tmp = 0
                            break
                    else:
                        break
                elif TSDApp.WorkbookStats.DataCodesLastRow != 0:
                    ExitFromFct = 1
                    break
                if ok == 0:
                    if str(cell.Value).casefold() == "Référence".casefold().strip() or str(cell.Value).casefold().strip() == "Reference".casefold():
                        refColIndex = cell.Column
                        refRowIndex = cell.Row
                        indexCol = 1
                        col_range = 1
                    if col_range == 1:
                        if cell.Borders(8).LineStyle != -4142 and cell != None:
                            indexCol += 1
                            pass
                        else:
                            lastCol = cell.Column
                            tmp = 1
                            ok = 1
                            break
                else:
                    break

        if refColIndex == 0:
            var = 1

        if var == 0:
            TSDApp.WorkbookStats.DataCodesLastRow = lastFilledCell
            TSDApp.WorkbookStats.DataCodesLastCol = lastCol

        else:
            TSDApp.WorkbookStats.DataCodesLastRow = None
            TSDApp.WorkbookStats.DataCodesLastCol = None
    else:
        TSDApp.WorkbookStats.DataCodesLastRow = None
        TSDApp.WorkbookStats.DataCodesLastCol = None


def getReadDataIOInfo(workBook, TSDApp):
    temp = workBook.Sheets
    sheetNames = []
    for sheet in temp:
        sheetNames.append(sheet.Name.strip().casefold())
    TSDApp.WorkbookStats.sheetNames = sheetNames
    if "read data and io control" in sheetNames:
        TSDApp.WorkbookStats.hasReadDataIO = True
        try:
            index = sheetNames.index("read data and io control") + 1
        except:
            pass
        TSDApp.WorkbookStats.ReadDataIOIndex = index
    else:
        TSDApp.WorkbookStats.hasReadDataIO = False

    if TSDApp.WorkbookStats.hasReadDataIO == True:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ReadDataIOIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0
        ok = 0
        col_range = 0
        lastCol = 0
        tmp = 0
        ExitFromFct = 0
        TSDApp.WorkbookStats.ReadDataIOLastRow = 0
        lastFilledCell = 0

        for cellRow in workSheet.Rows:
            col_range = 0
            if ExitFromFct == 1:
                break
            for cell in cellRow.Cells:
                if tmp != 0:
                    ok = 1
                    if col_range == 0:
                        if cell.Borders(9).LineStyle != -4142:
                            if cell.Value is not None:
                                lastFilledCell = cell.Row
                        else:
                            TSDApp.WorkbookStats.ReadDataIOLastRow = cell.Row
                            tmp = 0
                            break
                    else:
                        break
                elif TSDApp.WorkbookStats.ReadDataIOLastRow != 0:
                    ExitFromFct = 1
                    break
                if ok == 0:
                    if str(cell.Value).casefold() == "Référence".casefold().strip() or str(cell.Value).casefold().strip() == "Reference".casefold():
                        refColIndex = cell.Column
                        refRowIndex = cell.Row
                        indexCol = 1
                        col_range = 1
                    if col_range == 1:
                        if cell.Borders(8).LineStyle != -4142 and cell != None:
                            indexCol += 1
                            pass
                        else:
                            lastCol = cell.Column
                            tmp = 1
                            ok = 1
                            break
                else:
                    break

        if refColIndex == 0:
            var = 1

        if var == 0:
            TSDApp.WorkbookStats.ReadDataIOLastRow = lastFilledCell
            TSDApp.WorkbookStats.ReadDataIOLastCol = lastCol

        else:
            TSDApp.WorkbookStats.ReadDataIOLastRow = None
            TSDApp.WorkbookStats.ReadDataIOLastCol = None
    else:
        TSDApp.WorkbookStats.ReadDataIOLastRow = None
        TSDApp.WorkbookStats.ReadDataIOLastCol = None

def getNotEmbeddedDiagnosisInfo(workBook, TSDApp):
    temp = workBook.Sheets
    sheetNames = []
    for sheet in temp:
        sheetNames.append(sheet.Name.strip().casefold())
    TSDApp.WorkbookStats.sheetNames = sheetNames
    if "not embedded diagnosis" in sheetNames:
        TSDApp.WorkbookStats.hasNotEmbDiag = True
        try:
            index = sheetNames.index("not embedded diagnosis") + 1
        except:
            pass
        TSDApp.WorkbookStats.NotEmbDiagIndex = index
    else:
        TSDApp.WorkbookStats.hasNotEmbDiag = False

    if TSDApp.WorkbookStats.hasNotEmbDiag == True:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.NotEmbDiagIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0
        ok = 0
        col_range = 0
        lastCol = 0
        tmp = 0
        ExitFromFct = 0
        TSDApp.WorkbookStats.NotEmbDiagLastRow = 0
        lastFilledCell = 0

        for cellRow in workSheet.Rows:
            col_range = 0
            if ExitFromFct == 1:
                break
            for cell in cellRow.Cells:
                if tmp != 0:
                    ok = 1
                    if col_range == 0:
                        if cell.Borders(9).LineStyle != -4142:
                            if cell.Value is not None:
                                lastFilledCell = cell.Row
                        else:
                            TSDApp.WorkbookStats.NotEmbDiagLastRow = cell.Row
                            tmp = 0
                            break
                    else:
                        break
                elif TSDApp.WorkbookStats.NotEmbDiagLastRow != 0:
                    ExitFromFct = 1
                    break
                if ok == 0:
                    if str(cell.Value).casefold() == "Référence".casefold().strip() or str(cell.Value).casefold().strip() == "Reference".casefold():
                        refColIndex = cell.Column
                        refRowIndex = cell.Row
                        indexCol = 1
                        col_range = 1
                    if col_range == 1:
                        if cell.Borders(8).LineStyle != -4142 and cell != None:
                            indexCol += 1
                            pass
                        else:
                            lastCol = cell.Column
                            tmp = 1
                            ok = 1
                            break
                else:
                    break

        if refColIndex == 0:
            var = 1

        if var == 0:
            TSDApp.WorkbookStats.NotEmbDiagLastRow = lastFilledCell
            TSDApp.WorkbookStats.NotEmbDiagLastCol = lastCol

        else:
            TSDApp.WorkbookStats.NotEmbDiagLastRow = None
            TSDApp.WorkbookStats.NotEmbDiagLastCol = None
    else:
        TSDApp.WorkbookStats.NotEmbDiagLastRow = None
        TSDApp.WorkbookStats.NotEmbDiagLastCol = None

def DOC3Info(workBook, TSDApp):
    getTableInfo(workBook, TSDApp)
    getCodesDefautsInfo(workBook, TSDApp)
    getMesuresEtCommandesInfo(workBook, TSDApp)
    getDiagnosticDebarquesInfo(workBook, TSDApp)
    getListeMDDInfo(workBook, TSDApp)
    getEffetsClientsInfo(workBook, TSDApp)

def DOC4Info(workBook, TSDApp):
    getTableInfo(workBook, TSDApp)
    getEffetsClientsInfo(workBook, TSDApp)
    getReqOfTechEffectsInfo(workBook, TSDApp)

def DOC5Info(workBook, TSDApp):
    getTableInfo(workBook, TSDApp)
    getCodesDefautsInfo(workBook, TSDApp)
    getDataTroubleCodesInfo(workBook, TSDApp)
    getReadDataIOInfo(workBook, TSDApp)
    getEffetsClientsInfo(workBook, TSDApp)
    getListeMDDInfo(workBook, TSDApp)
    getNotEmbeddedDiagnosisInfo(workBook, TSDApp)