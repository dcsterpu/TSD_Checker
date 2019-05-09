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
                        TSDApp.WorkbookStats.ReqTechRefColIndex = cell.Column
                        TSDApp.WorkbookStats.ReqTechRefRowIndex = cell.Row
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

        if TSDApp.WorkbookStats.ReqTechRefColIndex == 0:
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

def getConstituants(workBook, TSDApp):
    temp = workBook.Sheets
    sheetNames = []
    for sheet in temp:
        sheetNames.append(sheet.Name.strip().casefold())
    TSDApp.WorkbookStats.sheetNames = sheetNames
    if "constituants" in sheetNames:
        TSDApp.WorkbookStats.hasConstituants = True
        try:
            index = sheetNames.index("constituants") + 1
        except:
            pass
        TSDApp.WorkbookStats.constituantsIndex = index
    else:
        TSDApp.WorkbookStats.hasConstituants = False

    if TSDApp.WorkbookStats.hasConstituants == True:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.constituantsIndex)
        var = 0
        ok = 0
        col_range = 0
        lastCol = 0
        tmp = 0
        ExitFromFct = 0
        TSDApp.WorkbookStats.constituantsLastRow = 0
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
                            TSDApp.WorkbookStats.constituantsLastRow = cell.Row
                            tmp = 0
                            break
                    else:
                        break
                elif TSDApp.WorkbookStats.constituantsLastRow != 0:
                    ExitFromFct = 1
                    break
                if ok == 0:
                    if str(cell.Value).casefold().strip() == "Noms".casefold():
                        TSDApp.WorkbookStats.constituantsRefColIndex = cell.Column
                        TSDApp.WorkbookStats.constituantsRefRowIndex = cell.Row
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
            TSDApp.WorkbookStats.constituantsLastRow = lastFilledCell
            TSDApp.WorkbookStats.constituantsLastCol = lastCol

        else:
            TSDApp.WorkbookStats.constituantsLastRow = None
            TSDApp.WorkbookStats.constituantsLastCol = None
    else:
        TSDApp.WorkbookStats.constituantsLastRow = None
        TSDApp.WorkbookStats.constituantsLastCol = None

def getERInfo(workBook, TSDApp):
    temp = workBook.Sheets
    sheetNames = []
    for sheet in temp:
        sheetNames.append(sheet.Name.strip().casefold())
    TSDApp.WorkbookStats.sheetNames = sheetNames
    if "er" in sheetNames:
        TSDApp.WorkbookStats.hasER = True
        try:
            index = sheetNames.index("er") + 1
        except:
            pass
        TSDApp.WorkbookStats.ERIndex = index
    else:
        TSDApp.WorkbookStats.hasER = False

    if TSDApp.WorkbookStats.hasER == True:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.ERIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0
        ok = 0
        col_range = 0
        lastCol = 0
        tmp = 0
        ExitFromFct = 0
        TSDApp.WorkbookStats.ERLastRow = 0
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
                            TSDApp.WorkbookStats.ERLastRow = cell.Row
                            tmp = 0
                            break
                    else:
                        break
                elif TSDApp.WorkbookStats.ERLastRow != 0:
                    ExitFromFct = 1
                    break
                if ok == 0:
                    if str(cell.Value).casefold().strip() == "nom".casefold():
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
            TSDApp.WorkbookStats.ERLastRow = lastFilledCell
            TSDApp.WorkbookStats.ERLastCol = lastCol

        else:
            TSDApp.WorkbookStats.ERLastRow = None
            TSDApp.WorkbookStats.ERLastCol = None
    else:
        TSDApp.WorkbookStats.ERLastRow = None
        TSDApp.WorkbookStats.ERLastCol = None

def getSituationDeVieInfo(workBook, TSDApp):
    temp = workBook.Sheets
    sheetNames = []
    for sheet in temp:
        sheetNames.append(sheet.Name.strip().casefold())
    TSDApp.WorkbookStats.sheetNames = sheetNames
    if "situations de vie" in sheetNames:
        TSDApp.WorkbookStats.hasSitDeVie = True
        try:
            index = sheetNames.index("situations de vie") + 1
        except:
            pass
        TSDApp.WorkbookStats.SitDeVieIndex = index
    else:
        TSDApp.WorkbookStats.hasSitDeVie = False

    if TSDApp.WorkbookStats.hasSitDeVie == True:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.SitDeVieIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0
        ok = 0
        col_range = 0
        lastCol = 0
        tmp = 0
        ExitFromFct = 0
        TSDApp.WorkbookStats.SitDeVieLastRow = 0
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
                            TSDApp.WorkbookStats.SitDeVieLastRow = cell.Row
                            tmp = 0
                            break
                    else:
                        break
                elif TSDApp.WorkbookStats.SitDeVieLastRow != 0:
                    ExitFromFct = 1
                    break
                if ok == 0:
                    if str(cell.Value).casefold().strip() == "Situations de vie".casefold():
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
            TSDApp.WorkbookStats.SitDeVieLastRow = lastFilledCell
            TSDApp.WorkbookStats.SitDeVieLastCol = lastCol

        else:
            TSDApp.WorkbookStats.SitDeVieLastRow = None
            TSDApp.WorkbookStats.SitDeVieLastCol = None
    else:
        TSDApp.WorkbookStats.SitDeVieLastRow = None
        TSDApp.WorkbookStats.SitDeVieLastCol = None

def getDiagnosticNeedsInfo(workBook, TSDApp):
    temp = workBook.Sheets
    sheetNames = []
    for sheet in temp:
        sheetNames.append(sheet.Name.strip().casefold())
    TSDApp.WorkbookStats.sheetNames = sheetNames
    if "diagnostic needs" in sheetNames:
        TSDApp.WorkbookStats.hasDiagNeeds = True
        try:
            index = sheetNames.index("diagnostic needs") + 1
        except:
            pass
        TSDApp.WorkbookStats.DiagNeedsIndex = index
    else:
        TSDApp.WorkbookStats.hasDiagNeeds = False

    if TSDApp.WorkbookStats.hasDiagNeeds == True:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0
        ok = 0
        col_range = 0
        lastCol = 0
        tmp = 0
        ExitFromFct = 0
        TSDApp.WorkbookStats.DiagNeedsLastRow = 0
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
                            TSDApp.WorkbookStats.DiagNeedsLastRow = cell.Row
                            tmp = 0
                            break
                    else:
                        break
                elif TSDApp.WorkbookStats.DiagNeedsLastRow != 0:
                    ExitFromFct = 1
                    break
                if ok == 0:
                    if str(cell.Value).casefold().strip() == "Reference".casefold():
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
            TSDApp.WorkbookStats.DiagNeedsLastRow = lastFilledCell
            TSDApp.WorkbookStats.DiagNeedsLastCol = lastCol

        else:
            TSDApp.WorkbookStats.DiagNeedsLastRow = None
            TSDApp.WorkbookStats.DiagNeedsLastCol = None
    else:
        TSDApp.WorkbookStats.DiagNeedsLastRow = None
        TSDApp.WorkbookStats.DiagNeedsLastCol = None

def getFearedEventInfo(workBook, TSDApp):
    temp = workBook.Sheets
    sheetNames = []
    for sheet in temp:
        sheetNames.append(sheet.Name.strip().casefold())
    TSDApp.WorkbookStats.sheetNames = sheetNames
    if "feared events" in sheetNames:
        TSDApp.WorkbookStats.hasDiagNeeds = True
        try:
            index = sheetNames.index("feared events") + 1
        except:
            pass
        TSDApp.WorkbookStats.DiagNeedsIndex = index
    else:
        TSDApp.WorkbookStats.hasDiagNeeds = False

    if TSDApp.WorkbookStats.hasDiagNeeds == True:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.DiagNeedsIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0
        ok = 0
        col_range = 0
        lastCol = 0
        tmp = 0
        ExitFromFct = 0
        TSDApp.WorkbookStats.FearedEventLastRow = 0
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
                            TSDApp.WorkbookStats.FearedEventLastRow = cell.Row
                            tmp = 0
                            break
                    else:
                        break
                elif TSDApp.WorkbookStats.FearedEventLastRow != 0:
                    ExitFromFct = 1
                    break
                if ok == 0:
                    if str(cell.Value).casefold().strip() == "Description".casefold():
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
            TSDApp.WorkbookStats.FearedEventLastRow = lastFilledCell
            TSDApp.WorkbookStats.FearedEventLastCol = lastCol

        else:
            TSDApp.WorkbookStats.FearedEventLastRow = None
            TSDApp.WorkbookStats.FearedEventLastCol = None
    else:
        TSDApp.WorkbookStats.FearedEventLastRow = None
        TSDApp.WorkbookStats.FearedEventLastCol = None

def getSystemInfo(workBook, TSDApp):
    temp = workBook.Sheets
    sheetNames = []
    for sheet in temp:
        sheetNames.append(sheet.Name.strip().casefold())
    TSDApp.WorkbookStats.sheetNames = sheetNames
    if "system" in sheetNames:
        TSDApp.WorkbookStats.hasSystem = True
        try:
            index = sheetNames.index("system") + 1
        except:
            pass
        TSDApp.WorkbookStats.SystemIndex = index
    else:
        TSDApp.WorkbookStats.hasSystem = False

    if TSDApp.WorkbookStats.hasSystem == True:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.SystemIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0
        ok = 0
        col_range = 0
        lastCol = 0
        tmp = 0
        ExitFromFct = 0
        TSDApp.WorkbookStats.SystemLastRow = 0
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
                            TSDApp.WorkbookStats.SystemLastRow = cell.Row
                            tmp = 0
                            break
                    else:
                        break
                elif TSDApp.WorkbookStats.SystemLastRow != 0:
                    ExitFromFct = 1
                    break
                if ok == 0:
                    if str(cell.Value).casefold().strip() == "Label".casefold():
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
            TSDApp.WorkbookStats.SystemLastRow = lastFilledCell
            TSDApp.WorkbookStats.SystemLastCol = lastCol

        else:
            TSDApp.WorkbookStats.SystemLastRow = None
            TSDApp.WorkbookStats.SystemLastCol = None
    else:
        TSDApp.WorkbookStats.SystemLastRow = None
        TSDApp.WorkbookStats.SystemLastCol = None

def getOperationSituationInfo(workBook, TSDApp):
    temp = workBook.Sheets
    sheetNames = []
    for sheet in temp:
        sheetNames.append(sheet.Name.strip().casefold())
    TSDApp.WorkbookStats.sheetNames = sheetNames
    if "operation situation" in sheetNames:
        TSDApp.WorkbookStats.hasOpSit = True
        try:
            index = sheetNames.index("operation situation") + 1
        except:
            pass
        TSDApp.WorkbookStats.OpSitIndex = index
    else:
        TSDApp.WorkbookStats.hasOpSit = False

    if TSDApp.WorkbookStats.hasOpSit == True:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.OpSitIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0
        ok = 0
        col_range = 0
        lastCol = 0
        tmp = 0
        ExitFromFct = 0
        TSDApp.WorkbookStats.OpSitLastRow = 0
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
                            TSDApp.WorkbookStats.OpSitLastRow = cell.Row
                            tmp = 0
                            break
                    else:
                        break
                elif TSDApp.WorkbookStats.OpSitLastRow != 0:
                    ExitFromFct = 1
                    break
                if ok == 0:
                    if str(cell.Value).casefold().strip() == "Description".casefold():
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
            TSDApp.WorkbookStats.OpSitLastRow = lastFilledCell
            TSDApp.WorkbookStats.OpSitLastCol = lastCol

        else:
            TSDApp.WorkbookStats.OpSitLastRow = None
            TSDApp.WorkbookStats.OpSitLastCol = None
    else:
        TSDApp.WorkbookStats.OpSitLastRow = None
        TSDApp.WorkbookStats.OpSitLastCol = None

def getTechnicalEffectInfo(workBook, TSDApp):
    temp = workBook.Sheets
    sheetNames = []
    for sheet in temp:
        sheetNames.append(sheet.Name.strip().casefold())
    TSDApp.WorkbookStats.sheetNames = sheetNames
    if "technical effect" in sheetNames or "effets techniques"in sheetNames:
        TSDApp.WorkbookStats.hasTechEff = True
        try:
            index = sheetNames.index("technical effect") + 1
        except:
            index = sheetNames.index("effets techniques") + 1
        TSDApp.WorkbookStats.TechEffIndex = index
    else:
        TSDApp.WorkbookStats.hasTechEff = False

    if TSDApp.WorkbookStats.hasTechEff == True:
        workSheet = workBook.Sheets(TSDApp.WorkbookStats.TechEffIndex)
        refColIndex = 0
        refRowIndex = 0
        var = 0
        ok = 0
        col_range = 0
        lastCol = 0
        tmp = 0
        ExitFromFct = 0
        TSDApp.WorkbookStats.TechEffLastRow = 0
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
                            TSDApp.WorkbookStats.TechEffLastRow = cell.Row
                            tmp = 0
                            break
                    else:
                        break
                elif TSDApp.WorkbookStats.TechEffLastRow != 0:
                    ExitFromFct = 1
                    break
                if ok == 0:
                    if str(cell.Value).casefold().strip() == "Name".casefold():
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
            TSDApp.WorkbookStats.TechEffLastRow = lastFilledCell
            TSDApp.WorkbookStats.TechEffLastCol = lastCol

        else:
            TSDApp.WorkbookStats.TechEffLastRow = None
            TSDApp.WorkbookStats.TechEffLastCol = None
    else:
        TSDApp.WorkbookStats.TechEffLastRow = None
        TSDApp.WorkbookStats.TechEffLastCol = None



def DOC3Info(workBook, TSDApp):
    getTableInfo(workBook, TSDApp)
    getCodesDefautsInfo(workBook, TSDApp)
    getMesuresEtCommandesInfo(workBook, TSDApp)
    getDiagnosticDebarquesInfo(workBook, TSDApp)
    getListeMDDInfo(workBook, TSDApp)
    getEffetsClientsInfo(workBook, TSDApp)
    getConstituants(workBook, TSDApp)
    getERInfo(workBook, TSDApp)
    getSituationDeVieInfo(workBook, TSDApp)

def DOC4Info(workBook, TSDApp):
    getTableInfo(workBook, TSDApp)
    getEffetsClientsInfo(workBook, TSDApp)
    getReqOfTechEffectsInfo(workBook, TSDApp)
    getDiagnosticNeedsInfo(workBook, TSDApp)
    getFearedEventInfo(workBook, TSDApp)
    getSystemInfo(workBook, TSDApp)
    getOperationSituationInfo(workBook, TSDApp)

def DOC5Info(workBook, TSDApp):
    getTableInfo(workBook, TSDApp)
    getCodesDefautsInfo(workBook, TSDApp)
    getDataTroubleCodesInfo(workBook, TSDApp)
    getReadDataIOInfo(workBook, TSDApp)
    getEffetsClientsInfo(workBook, TSDApp)
    getListeMDDInfo(workBook, TSDApp)
    getNotEmbeddedDiagnosisInfo(workBook, TSDApp)
    getTechnicalEffectInfo(workBook, TSDApp)