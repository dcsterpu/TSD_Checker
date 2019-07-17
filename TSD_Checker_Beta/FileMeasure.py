import TSD_Checker_V6_0
import inspect


def getTableInfo(workBook, TSDApp):
    sheetNames = []
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())

    if "tableau" in sheetNames or "table" in sheetNames:
        TSDApp.WorkbookStats.hasTable = True
        try:
            index = sheetNames.index("tableau")
        except:
            index = sheetNames.index("table")
        TSDApp.WorkbookStats.tableIndex = index
    else:
        TSDApp.WorkbookStats.hasTable = False

    if TSDApp.WorkbookStats.hasTable == True:

        TSDApp.WorkbookStats.tableLastRow = workBook.sheet_by_index(index).nrows
        TSDApp.WorkbookStats.tableLastCol = workBook.sheet_by_index(index).ncols


def getCodesDefautsInfo(workBook, TSDApp):
    sheetNames = []
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())
    if "codes défauts" in sheetNames or "data trouble codes" in sheetNames:
        TSDApp.WorkbookStats.hasCode = True
        try:
            index = sheetNames.index("codes défauts")
        except:
            index = sheetNames.index("data trouble codes")
        TSDApp.WorkbookStats.codeIndex = index
    else:
        TSDApp.WorkbookStats.hasCode = False

    if TSDApp.WorkbookStats.hasCode == True:
        TSDApp.WorkbookStats.codeLastRow = workBook.sheet_by_index(index).nrows
        TSDApp.WorkbookStats.codeLastCol = workBook.sheet_by_index(index).ncols


def getMesuresEtCommandesInfo(workBook, TSDApp):
    sheetNames = []
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())
    if "mesures et commandes" in sheetNames:
        TSDApp.WorkbookStats.hasMeasure = True
        try:
            index = sheetNames.index("mesures et commandes")
        except:
            pass
        TSDApp.WorkbookStats.measureIndex = index
    else:
        TSDApp.WorkbookStats.hasMeasure = False

    if TSDApp.WorkbookStats.hasMeasure == True:
        TSDApp.WorkbookStats.measureLastRow = workBook.sheet_by_index(index).nrows
        TSDApp.WorkbookStats.measureLastCol = workBook.sheet_by_index(index).ncols


def getDiagnosticDebarquesInfo(workBook, TSDApp):
    sheetNames = []
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())
    if "diagnostic débarqués" in sheetNames:
        TSDApp.WorkbookStats.hasDiagDeb = True
        try:
            index = sheetNames.index("diagnostic débarqués")
        except:
            pass
        TSDApp.WorkbookStats.DiagDebIndex = index
    else:
        TSDApp.WorkbookStats.hasDiagDeb = False

    if TSDApp.WorkbookStats.hasDiagDeb == True:
        TSDApp.WorkbookStats.DiagDebLastRow = workBook.sheet_by_index(index).nrows
        TSDApp.WorkbookStats.DiagDebLastCol = workBook.sheet_by_index(index).ncols


def getListeMDDInfo(workBook, TSDApp):
    sheetNames = []
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())
    if "degraded mode" in sheetNames or "liste mdd" in sheetNames:
        TSDApp.WorkbookStats.hasMDD = True
        try:
            index = sheetNames.index("degraded mode")
        except:
            index = sheetNames.index("liste mdd")
        TSDApp.WorkbookStats.MDDIndex = index
    else:
        TSDApp.WorkbookStats.hasMDD = False

    if TSDApp.WorkbookStats.hasMDD == True:
        TSDApp.WorkbookStats.MDDLastRow = workBook.sheet_by_index(index).nrows
        TSDApp.WorkbookStats.MDDLastCol = workBook.sheet_by_index(index).ncols


def getEffetsClientsInfo(workBook, TSDApp):
    sheetNames = []
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())
    if "effets clients" in sheetNames or "customer effects" in sheetNames or "customer effect" in sheetNames:
        TSDApp.WorkbookStats.hasEffClients = True
        try:
            index = sheetNames.index("effets clients")
        except:
            if "customer effect" in sheetNames:
                index = sheetNames.index("customer effect")
            else:
                index = sheetNames.index("customer effects")
        TSDApp.WorkbookStats.EffClientsIndex = index
    else:
        TSDApp.WorkbookStats.hasEffClients = False

    if TSDApp.WorkbookStats.hasEffClients == True:
        TSDApp.WorkbookStats.EffClientsLastRow = workBook.sheet_by_index(index).nrows
        TSDApp.WorkbookStats.EffClientsLastCol = workBook.sheet_by_index(index).ncols


def getReqOfTechEffectsInfo(workBook, TSDApp):
    sheetNames = []
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())
    if "req. of tech. effects" in sheetNames:
        TSDApp.WorkbookStats.hasReqTech = True
        try:
            index = sheetNames.index("req. of tech. effects")
        except:
            pass
        TSDApp.WorkbookStats.ReqTechIndex = index
    else:
        TSDApp.WorkbookStats.hasReqTech = False

    if TSDApp.WorkbookStats.hasReqTech == True:
        TSDApp.WorkbookStats.ReqTechLastRow = workBook.sheet_by_index(index).nrows
        TSDApp.WorkbookStats.ReqTechLastCol = workBook.sheet_by_index(index).ncols


def getDataTroubleCodesInfo(workBook, TSDApp):
    sheetNames = []
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())
    if "data trouble codes" in sheetNames:
        TSDApp.WorkbookStats.hasDataCodes = True
        try:
            index = sheetNames.index("data trouble codes")
        except:
            pass
        TSDApp.WorkbookStats.DataCodesIndex = index
    else:
        TSDApp.WorkbookStats.hasDataCodes = False

    if TSDApp.WorkbookStats.hasDataCodes == True:
        TSDApp.WorkbookStats.DataCodesLastRow = workBook.sheet_by_index(index).nrows
        TSDApp.WorkbookStats.DataCodesLastCol = workBook.sheet_by_index(index).ncols


def getReadDataIOInfo(workBook, TSDApp):
    sheetNames = []
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())
    if "read data and io control" in sheetNames:
        TSDApp.WorkbookStats.hasReadDataIO = True
        try:
            index = sheetNames.index("read data and io control")
        except:
            pass
        TSDApp.WorkbookStats.ReadDataIOIndex = index
    else:
        TSDApp.WorkbookStats.hasReadDataIO = False

    if TSDApp.WorkbookStats.hasReadDataIO == True:
        TSDApp.WorkbookStats.ReadDataIOLastRow = workBook.sheet_by_index(index).nrows
        TSDApp.WorkbookStats.ReadDataIOLastCol = workBook.sheet_by_index(index).ncols


def getNotEmbeddedDiagnosisInfo(workBook, TSDApp):
    sheetNames = []
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())
    if "not embedded diagnosis" in sheetNames:
        TSDApp.WorkbookStats.hasNotEmbDiag = True
        try:
            index = sheetNames.index("not embedded diagnosis")
        except:
            pass
        TSDApp.WorkbookStats.NotEmbDiagIndex = index
    else:
        TSDApp.WorkbookStats.hasNotEmbDiag = False

    if TSDApp.WorkbookStats.hasNotEmbDiag == True:
        TSDApp.WorkbookStats.NotEmbDiagLastRow = workBook.sheet_by_index(index).nrows
        TSDApp.WorkbookStats.NotEmbDiagLastCol = workBook.sheet_by_index(index).ncols


def getConstituants(workBook, TSDApp):
    sheetNames = []
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())
    if "constituants" in sheetNames:
        TSDApp.WorkbookStats.hasConstituants = True
        try:
            index = sheetNames.index("constituants")
        except:
            pass
        TSDApp.WorkbookStats.constituantsIndex = index
    else:
        TSDApp.WorkbookStats.hasConstituants = False

    if TSDApp.WorkbookStats.hasConstituants == True:
        TSDApp.WorkbookStats.constituantsLastRow = workBook.sheet_by_index(index).nrows
        TSDApp.WorkbookStats.constituantsLastCol = workBook.sheet_by_index(index).ncols


def getERInfo(workBook, TSDApp):
    sheetNames = []
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())
    if "er" in sheetNames:
        TSDApp.WorkbookStats.hasER = True
        try:
            index = sheetNames.index("er")
        except:
            pass
        TSDApp.WorkbookStats.ERIndex = index
    else:
        TSDApp.WorkbookStats.hasER = False

    if TSDApp.WorkbookStats.hasER == True:
        TSDApp.WorkbookStats.ERLastRow = workBook.sheet_by_index(index).nrows
        TSDApp.WorkbookStats.ERLastCol = workBook.sheet_by_index(index).ncols


def getSituationDeVieInfo(workBook, TSDApp):
    sheetNames = []
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())
    if "situations de vie" in sheetNames:
        TSDApp.WorkbookStats.hasSitDeVie = True
        try:
            index = sheetNames.index("situations de vie")
        except:
            pass
        TSDApp.WorkbookStats.SitDeVieIndex = index
    else:
        TSDApp.WorkbookStats.hasSitDeVie = False

    if TSDApp.WorkbookStats.hasSitDeVie == True:
        TSDApp.WorkbookStats.SitDeVieLastRow = workBook.sheet_by_index(index).nrows
        TSDApp.WorkbookStats.SitDeVieLastCol = workBook.sheet_by_index(index).ncols


def getDiagnosticNeedsInfo(workBook, TSDApp):
    sheetNames = []
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())
    if "diagnostic needs" in sheetNames:
        TSDApp.WorkbookStats.hasDiagNeeds = True
        try:
            index = sheetNames.index("diagnostic needs")
        except:
            pass
        TSDApp.WorkbookStats.DiagNeedsIndex = index
    else:
        TSDApp.WorkbookStats.hasDiagNeeds = False

    if TSDApp.WorkbookStats.hasDiagNeeds == True:
        TSDApp.WorkbookStats.DiagNeedsLastRow = workBook.sheet_by_index(index).nrows
        TSDApp.WorkbookStats.DiagNeedsLastCol = workBook.sheet_by_index(index).ncols


def getFearedEventInfo(workBook, TSDApp):
    sheetNames = []
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())
    if "feared events" in sheetNames:
        TSDApp.WorkbookStats.hasFearedEvent = True
        try:
            index = sheetNames.index("feared events")
        except:
            pass
        TSDApp.WorkbookStats.FearedEventIndex = index
    else:
        TSDApp.WorkbookStats.hasFearedEvent = False

    if TSDApp.WorkbookStats.hasFearedEvent == True:
        TSDApp.WorkbookStats.FearedEventLastRow = workBook.sheet_by_index(index).nrows
        TSDApp.WorkbookStats.FearedEventLastCol = workBook.sheet_by_index(index).ncols


def getSystemInfo(workBook, TSDApp):
    sheetNames = []
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())
    if "system" in sheetNames:
        TSDApp.WorkbookStats.hasSystem = True
        try:
            index = sheetNames.index("system")
        except:
            index = sheetNames.index("système")
        TSDApp.WorkbookStats.SystemIndex = index
    else:
        TSDApp.WorkbookStats.hasSystem = False

    if TSDApp.WorkbookStats.hasSystem == True:
        TSDApp.WorkbookStats.SystemLastRow = workBook.sheet_by_index(index).nrows
        TSDApp.WorkbookStats.SystemLastCol = workBook.sheet_by_index(index).ncols


def getOperationSituationInfo(workBook, TSDApp):
    sheetNames = []
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())
    if "operation situation" in sheetNames:
        TSDApp.WorkbookStats.hasOpSit = True
        try:
            index = sheetNames.index("operation situation")
        except:
            pass
        TSDApp.WorkbookStats.OpSitIndex = index
    else:
        TSDApp.WorkbookStats.hasOpSit = False

    if TSDApp.WorkbookStats.hasOpSit == True:
        TSDApp.WorkbookStats.OpSitLastRow = workBook.sheet_by_index(index).nrows
        TSDApp.WorkbookStats.OpSitLastCol = workBook.sheet_by_index(index).ncols


def getTechnicalEffectInfo(workBook, TSDApp):
    sheetNames = []
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())
    if "technical effect" in sheetNames or "effets techniques" in sheetNames:
        TSDApp.WorkbookStats.hasTechEff = True
        try:
            index = sheetNames.index("technical effect")
        except:
            index = sheetNames.index("effets techniques")
        TSDApp.WorkbookStats.TechEffIndex = index
    else:
        TSDApp.WorkbookStats.hasTechEff = False

    if TSDApp.WorkbookStats.hasTechEff == True:
        TSDApp.WorkbookStats.TechEffLastRow = workBook.sheet_by_index(index).nrows
        TSDApp.WorkbookStats.TechEffLastCol = workBook.sheet_by_index(index).ncols


def getPartsInfo(workBook, TSDApp):
    sheetNames = []
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())
    if "parts" in sheetNames:
        TSDApp.WorkbookStats.hasParts = True
        try:
            index = sheetNames.index("parts")
        except:
            pass
        TSDApp.WorkbookStats.PartsIndex = index
    else:
        TSDApp.WorkbookStats.hasParts = False

    if TSDApp.WorkbookStats.hasParts == True:
        TSDApp.WorkbookStats.PartsLastRow = workBook.sheet_by_index(index).nrows
        TSDApp.WorkbookStats.PartsLastCol = workBook.sheet_by_index(index).ncols


def getVariantInfo(workBook, TSDApp):
    sheetNames = []
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())
    if "variant" in sheetNames:
        TSDApp.WorkbookStats.hasVariant = True
        try:
            index = sheetNames.index("variant")
        except:
            pass
        TSDApp.WorkbookStats.VariantIndex = index
    else:
        TSDApp.WorkbookStats.hasVariant = False

    if TSDApp.WorkbookStats.hasVariant == True:
        TSDApp.WorkbookStats.VariantLastRow = workBook.sheet_by_index(index).nrows
        TSDApp.WorkbookStats.VariantLastCol = workBook.sheet_by_index(index).ncols


def getSituationInfo(workBook, TSDApp):
    sheetNames = []
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())
    if "situation" in sheetNames:
        TSDApp.WorkbookStats.hasSituation = True
        try:
            index = sheetNames.index("situation")
        except:
            pass
        TSDApp.WorkbookStats.SituationIndex = index
    else:
        TSDApp.WorkbookStats.hasSituation = False

    if TSDApp.WorkbookStats.hasSituation == True:
        TSDApp.WorkbookStats.SituationLastRow = workBook.sheet_by_index(index).nrows
        TSDApp.WorkbookStats.SituationLastCol = workBook.sheet_by_index(index).ncols


def getDegradedModeInfo(workBook, TSDApp):
    sheetNames = []
    for sheet in workBook.sheet_names():
        sheetNames.append(sheet.casefold())
    if "degraded mode" in sheetNames:
        TSDApp.WorkbookStats.hasDegradedMode = True
        try:
            index = sheetNames.index("degraded mode")
        except:
            pass
        TSDApp.WorkbookStats.DegradedModeIndex = index
    else:
        TSDApp.WorkbookStats.hasDegradedMode = False

    if TSDApp.WorkbookStats.hasDegradedMode == True:
        TSDApp.WorkbookStats.DegradedModeLastRow = workBook.sheet_by_index(index).nrows
        TSDApp.WorkbookStats.DegradedModeLastCol = workBook.sheet_by_index(index).ncols


def DOC3Info1(workBook, TSDApp):
    getTableInfo(workBook, TSDApp)
    getCodesDefautsInfo(workBook, TSDApp)
    getMesuresEtCommandesInfo(workBook, TSDApp)
    getDiagnosticDebarquesInfo(workBook, TSDApp)
    getListeMDDInfo(workBook, TSDApp)
    getEffetsClientsInfo(workBook, TSDApp)
    getConstituants(workBook, TSDApp)
    getERInfo(workBook, TSDApp)
    getSituationDeVieInfo(workBook, TSDApp)


def DOC4Info1(workBook, TSDApp):
    getTableInfo(workBook, TSDApp)
    getEffetsClientsInfo(workBook, TSDApp)
    getReqOfTechEffectsInfo(workBook, TSDApp)
    getDiagnosticNeedsInfo(workBook, TSDApp)
    getFearedEventInfo(workBook, TSDApp)
    getSystemInfo(workBook, TSDApp)
    getOperationSituationInfo(workBook, TSDApp)


def DOC5Info1(workBook, TSDApp):
    getTableInfo(workBook, TSDApp)
    getCodesDefautsInfo(workBook, TSDApp)
    getDataTroubleCodesInfo(workBook, TSDApp)
    getReadDataIOInfo(workBook, TSDApp)
    getEffetsClientsInfo(workBook, TSDApp)
    getNotEmbeddedDiagnosisInfo(workBook, TSDApp)
    getTechnicalEffectInfo(workBook, TSDApp)
    getFearedEventInfo(workBook, TSDApp)
    getPartsInfo(workBook, TSDApp)
    getVariantInfo(workBook, TSDApp)
    getSituationInfo(workBook, TSDApp)
    getDegradedModeInfo(workBook, TSDApp)