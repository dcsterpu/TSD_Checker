import TSD_Checker_V4_0
from lxml import etree, objectify


def DOC9Parser(TSDApp, ExcelApp, DOC9Path):

    try:
        DOC9 = ExcelApp.Workbooks.Open(DOC9Path)
    except:
        TSDApp.tab1.textbox.setText("ERROR: when trying to parse the criticity file " + DOC9Path.split('/')[-1])
        return
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

def DOC13Parser(TSDApp, ExcelApp, DOC13Path):
    try:
        DOC13 = ExcelApp.Workbooks.Open(DOC13Path)
    except:
        TSDApp.tab1.textbox.setText("ERROR: when trying to parse the diversity referential file " + DOC13Path.split('/')[-1])
        return
    workSheet = DOC13.Sheets("Liste EC")
    workSheetRangeValuesTuple = workSheet.UsedRange.Value
    testCol = 0
    temp_list = []
    final_list = []

    for rowTuple in workSheetRangeValuesTuple:
        if 'Nom CF /\nNom CO PLM (CF_CO)' in rowTuple:
            testCol = rowTuple.index("Nom CF /\nNom CO PLM (CF_CO)")
            break
    for rowTuple in workSheetRangeValuesTuple:
        temp_list.append(rowTuple[testCol])

    for elem in temp_list:
        if elem not in [None, "", 'Nom CF /\nNom CO PLM (CF_CO)']:
            final_list.append(elem)
    return final_list

def DOC8Parser(TSDApp ,ExcelApp, DOC8Path):
    try:
        DOC8 = ExcelApp.Workbooks.Open(DOC8Path)
    except:
        TSDApp.tab1.textbox.setText("ERROR: when trying to parse the CESARE file " + DOC8Path.split('/')[-1])
        return
    cnt = 0
    temp = DOC8.Sheets
    for sheet in temp:
        cnt = cnt + 1
        if "sous familles Cesare" in sheet.Name:
            index = cnt
            break
    workSheet = DOC8.Sheets(index)
    workSheetRangeValuesTuple = workSheet.UsedRange.Value
    testCol = 0
    temp_list = []
    final_list = []

    for rowTuple in workSheetRangeValuesTuple:
        #temp = tuple(tuple(b.strip() for b in a) for a in rowTuple)
        if '\xa0Nom de la sous famille\xa0' in rowTuple:
            testCol = rowTuple.index("\xa0Nom de la sous famille\xa0")
            break
    for rowTuple in workSheetRangeValuesTuple:
        temp_list.append(rowTuple[testCol])

    for elem in temp_list:
        if elem not in [None, "", '\xa0Nom de la sous famille\xa0']:
            final_list.append(elem.replace(u'\xa0', u''))
    return final_list

def DOC15Parser(TSDApp ,DOC15Path):
    if DOC15Path.endswith('.odx'):
        parser = etree.XMLParser(remove_comments=True)
        try:
            tree = objectify.parse(DOC15Path, parser=parser)
        except:
            TSDApp.tab1.textbox.setText("ERROR: when trying to parse the diagnostic messagerie file " + DOC15Path.split('/')[-1])
            return None, None
        root = tree.getroot()
        subfamily = root.find(".//BASE-VARIANT")
        subfamily_name = subfamily.attrib['ID']
        dids = root.findall(".//DATA-OBJECT-PROP")
        returnList = []
        for did in dids:
            returnList.append(did.attrib['ID'])
        return subfamily_name, returnList
    else:
        return None, None

def DOC10Coherence(ExcelApp, TSDApp, DOC10Path):
    try:
        DOC10 = ExcelApp.Workbooks.Open(DOC10Path)
    except:
        TSDApp.tab1.textbox.setText("ERROR: when trying to parse the Plan type Synthese diagnosticabilite file " + DOC10Path.split('/')[-1])
    sheets =  DOC10.Sheets
    sheetNames = list()
    for sheet in sheets:
        sheetNames.append(sheet.Name.strip().casefold())
    dict10 = []

    if "tableau" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("tableau") + 1
        workSheet = DOC10.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.tableLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict10.append(dictionary)

    if "codes défauts" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("codes défauts") + 1
        workSheet = DOC10.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.codeLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict10.append(dictionary)

    if "mesures et commandes" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("mesures et commandes") + 1
        workSheet = DOC10.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.measureLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict10.append(dictionary)

    if "diagnostic débarqués" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("diagnostic débarqués") + 1
        workSheet = DOC10.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.DiagDebLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict10.append(dictionary)

    if "effets clients" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("effets clients") + 1
        workSheet = DOC10.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.EffClientsLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict10.append(dictionary)

    if "er" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("er") + 1
        workSheet = DOC10.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.ERLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict10.append(dictionary)

    if "constituants" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("constituants") + 1
        workSheet = DOC10.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.constituantsLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict10.append(dictionary)

    if "situations de vie" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("situations de vie") + 1
        workSheet = DOC10.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.SitDeVieLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict10.append(dictionary)

    if "liste mdd" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("liste mdd") + 1
        workSheet = DOC10.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.MDDLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict10.append(dictionary)


    # dict10 = {}
    # dict10['02043_18_04939_WHOLENESS_1600'] = True
    # dict10['02043_18_04939_WHOLENESS_1601'] = True
    # dict10['02043_18_04939_WHOLENESS_1602'] = True
    # dict10['02043_18_04939_WHOLENESS_1603'] = True
    # dict10['02043_18_04939_WHOLENESS_1604'] = True
    # dict10['02043_18_04939_WHOLENESS_1605'] = True
    # dict10['02043_18_04939_WHOLENESS_1606'] = True
    # dict10['02043_18_04939_WHOLENESS_1607'] = True
    # dict10['02043_18_04939_WHOLENESS_1608'] = True
    # dict10['02043_18_04939_WHOLENESS_1609'] = True
    # dict10['02043_18_04939_WHOLENESS_1610'] = True
    # dict10['02043_18_04939_WHOLENESS_1611'] = False
    # dict10['02043_18_04939_WHOLENESS_1612'] = False
    # dict10['02043_18_04939_WHOLENESS_1613'] = True
    # dict10['02043_18_04939_WHOLENESS_1615'] = True
    # dict10['02043_18_04939_WHOLENESS_1616'] = True
    # dict10['02043_18_04939_WHOLENESS_1617'] = False
    # dict10['02043_18_04939_WHOLENESS_1618'] = True
    # dict10['02043_18_04939_WHOLENESS_1619'] = True
    # dict10['02043_18_04939_WHOLENESS_1620'] = True
    # dict10['02043_18_04939_WHOLENESS_1621'] = True
    # dict10['02043_18_04939_WHOLENESS_1622'] = False
    # dict10['02043_18_04939_WHOLENESS_1623'] = True
    # dict10['02043_18_04939_WHOLENESS_1624'] = False
    # dict10['02043_18_04939_WHOLENESS_1625'] = True
    # dict10['02043_18_04939_WHOLENESS_1626'] = False
    # dict10['02043_18_04939_WHOLENESS_1627'] = False
    # dict10['02043_18_04939_WHOLENESS_1628'] = False
    # dict10['02043_18_04939_WHOLENESS_1629'] = False
    # dict10['02043_18_04939_WHOLENESS_1630'] = False
    # dict10['02043_18_04939_WHOLENESS_1631'] = False
    # dict10['02043_18_04939_WHOLENESS_1632'] = False
    # dict10['02043_18_04939_WHOLENESS_1633'] = True
    # dict10['02043_18_04939_WHOLENESS_1650'] = True
    # dict10['02043_18_04939_WHOLENESS_1651'] = True
    # dict10['02043_18_04939_WHOLENESS_1652'] = True
    # dict10['02043_18_04939_WHOLENESS_1653'] = True
    # dict10['02043_18_04939_WHOLENESS_1654'] = True
    # dict10['02043_18_04939_WHOLENESS_1655'] = True
    # dict10['02043_18_04939_WHOLENESS_1656'] = True
    # dict10['02043_18_04939_WHOLENESS_1657'] = True
    # dict10['02043_18_04939_WHOLENESS_1658'] = True
    # dict10['02043_18_04939_WHOLENESS_1659'] = True
    # dict10['02043_18_04939_WHOLENESS_1660'] = True
    # dict10['02043_18_04939_WHOLENESS_1661'] = False
    # dict10['02043_18_04939_WHOLENESS_1662'] = False
    # dict10['02043_18_04939_WHOLENESS_1663'] = True
    # dict10['02043_18_04939_WHOLENESS_1664'] = True
    # dict10['02043_18_04939_WHOLENESS_1684'] = True
    # dict10['02043_18_04939_WHOLENESS_1685'] = True
    # dict10['02043_18_04939_WHOLENESS_1686'] = False
    # dict10['02043_18_04939_WHOLENESS_1687'] = False
    # dict10['02043_18_04939_WHOLENESS_1688'] = True
    # dict10['02043_18_04939_WHOLENESS_1689'] = True
    # dict10['02043_18_04939_WHOLENESS_1690'] = True
    # dict10['02043_18_04939_WHOLENESS_1691'] = True
    # dict10['02043_18_04939_WHOLENESS_1692'] = True
    # dict10['02043_18_04939_WHOLENESS_1693'] = False
    # dict10['02043_18_04939_WHOLENESS_1694'] = False
    # dict10['02043_18_04939_WHOLENESS_1700'] = True
    # dict10['02043_18_04939_WHOLENESS_1701'] = True
    # dict10['02043_18_04939_WHOLENESS_1702'] = True
    # dict10['02043_18_04939_WHOLENESS_1703'] = True
    # dict10['02043_18_04939_WHOLENESS_1704'] = True
    # dict10['02043_18_04939_WHOLENESS_1705'] = True
    # dict10['02043_18_04939_WHOLENESS_1706'] = True
    # dict10['02043_18_04939_WHOLENESS_1707'] = False
    # dict10['02043_18_04939_WHOLENESS_1708'] = True
    # dict10['02043_18_04939_WHOLENESS_1709'] = True
    # dict10['02043_18_04939_WHOLENESS_1710'] = True
    # dict10['02043_18_04939_WHOLENESS_1711'] = False
    # dict10['02043_18_04939_WHOLENESS_1712'] = False
    # dict10['02043_18_04939_WHOLENESS_1713'] = True
    # dict10['02043_18_04939_WHOLENESS_1714'] = True
    # dict10['02043_18_04939_WHOLENESS_1715'] = True
    # dict10['02043_18_04939_WHOLENESS_1716'] = True
    # dict10['02043_18_04939_WHOLENESS_1717'] = True
    # dict10['02043_18_04939_WHOLENESS_1718'] = False
    # dict10['02043_18_04939_WHOLENESS_1719'] = False
    # dict10['02043_18_04939_WHOLENESS_1720'] = True
    # dict10['02043_18_04939_WHOLENESS_1750'] = True
    # dict10['02043_18_04939_WHOLENESS_1751'] = True
    # dict10['02043_18_04939_WHOLENESS_1752'] = True
    # dict10['02043_18_04939_WHOLENESS_1753'] = True
    # dict10['02043_18_04939_WHOLENESS_1754'] = False
    # dict10['02043_18_04939_WHOLENESS_1755'] = False
    # dict10['02043_18_04939_WHOLENESS_1756'] = True
    # dict10['02043_18_04939_WHOLENESS_1757'] = True
    # dict10['02043_18_04939_WHOLENESS_1758'] = False
    # dict10['02043_18_04939_WHOLENESS_1759'] = False
    # dict10['02043_18_04939_WHOLENESS_1800'] = True
    # dict10['02043_18_04939_WHOLENESS_1801'] = True
    # dict10['02043_18_04939_WHOLENESS_1802'] = True
    # dict10['02043_18_04939_WHOLENESS_1803'] = False
    # dict10['02043_18_04939_WHOLENESS_1810'] = True
    # dict10['02043_18_04939_WHOLENESS_1811'] = True
    # dict10['02043_18_04939_WHOLENESS_1812'] = True
    # dict10['02043_18_04939_WHOLENESS_1813'] = True
    # dict10['02043_18_04939_WHOLENESS_1814'] = False
    # dict10['02043_18_04939_WHOLENESS_1815'] = False
    # dict10['02043_18_04939_WHOLENESS_1820'] = True
    # dict10['02043_18_04939_WHOLENESS_1821'] = True
    # dict10['02043_18_04939_WHOLENESS_1822'] = False
    # dict10['02043_18_04939_WHOLENESS_1823'] = False
    # dict10['02043_18_04939_WHOLENESS_1824'] = True
    # dict10['02043_18_04939_WHOLENESS_1825'] = False
    # dict10['02043_18_04939_WHOLENESS_1830'] = True
    # dict10['02043_18_04939_WHOLENESS_1831'] = False
    # dict10['02043_18_04939_WHOLENESS_1840'] = True
    # dict10['02043_18_04939_WHOLENESS_1841'] = False
    return dict10

def DOC11Coherence(ExcelApp, TSDApp, DOC11Path):
    try:
        DOC11 = ExcelApp.Workbooks.Open(DOC11Path)
    except:
        TSDApp.tab1.textbox.setText("ERROR: when trying to parse the Plan type AMDE TSD fonction vehicule file " + DOC11Path.split('/')[-1])
    sheets = DOC11.Sheets
    sheetNames = list()
    for sheet in sheets:
        sheetNames.append(sheet.Name.strip().casefold())

    dict11 = []

    if "table" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("table") + 1
        workSheet = DOC11.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.tableLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict11.append(dictionary)

    if "diagnostic needs" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("diagnostic needs") + 1
        workSheet = DOC11.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.DiagNeedsLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict11.append(dictionary)

    if "customer effects" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("customer effects") + 1
        workSheet = DOC11.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.EffClientsLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict11.append(dictionary)

    if "feared events" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("feared events") + 1
        workSheet = DOC11.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.FearedEventLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict11.append(dictionary)

    if "system" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("system") + 1
        workSheet = DOC11.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.SystemLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict11.append(dictionary)

    if "operation situation" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("operation situation") + 1
        workSheet = DOC11.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.OpSitLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict11.append(dictionary)
    return dict11
    # doc11 = {}
    # doc11['02043_18_04939_WHOLENESS_1300'] = True
    # doc11['02043_18_04939_WHOLENESS_1301'] = True
    # doc11['02043_18_04939_WHOLENESS_1302'] = True
    # doc11['02043_18_04939_WHOLENESS_1303'] = True
    # doc11['02043_18_04939_WHOLENESS_1304'] = True
    # doc11['02043_18_04939_WHOLENESS_1305'] = True
    # doc11['02043_18_04939_WHOLENESS_1306'] = True
    # doc11['02043_18_04939_WHOLENESS_1307'] = True
    # doc11['02043_18_04939_WHOLENESS_1308'] = True
    # doc11['02043_18_04939_WHOLENESS_1309'] = True
    # doc11['02043_18_04939_WHOLENESS_1310'] = True
    # doc11['02043_18_04939_WHOLENESS_1311'] = True
    # doc11['02043_18_04939_WHOLENESS_1312'] = True
    # doc11['02043_18_04939_WHOLENESS_1313'] = True
    # doc11['02043_18_04939_WHOLENESS_1314'] = True
    # doc11['02043_18_04939_WHOLENESS_1315'] = True
    # doc11['02043_18_04939_WHOLENESS_1316'] = False
    # doc11['02043_18_04939_WHOLENESS_1317'] = True
    # doc11['02043_18_04939_WHOLENESS_1318'] = True
    # doc11['02043_18_04939_WHOLENESS_1319'] = False
    # doc11['02043_18_04939_WHOLENESS_1320'] = False
    # doc11['02043_18_04939_WHOLENESS_1321'] = False
    # doc11['02043_18_04939_WHOLENESS_1322'] = False
    # doc11['02043_18_04939_WHOLENESS_1323'] = True
    # doc11['02043_18_04939_WHOLENESS_1324'] = True
    # doc11['02043_18_04939_WHOLENESS_1325'] = False
    # doc11['02043_18_04939_WHOLENESS_1326'] = False
    # doc11['02043_18_04939_WHOLENESS_1327'] = True
    # doc11['02043_18_04939_WHOLENESS_1328'] = True
    # doc11['02043_18_04939_WHOLENESS_1329'] = True
    # doc11['02043_18_04939_WHOLENESS_1330'] = False
    # doc11['02043_18_04939_WHOLENESS_1331'] = True
    # doc11['02043_18_04939_WHOLENESS_1332'] = False
    # doc11['02043_18_04939_WHOLENESS_1333'] = True
    # doc11['02043_18_04939_WHOLENESS_1334'] = False
    # doc11['02043_18_04939_WHOLENESS_1350'] = True
    # doc11['02043_18_04939_WHOLENESS_1351'] = True
    # doc11['02043_18_04939_WHOLENESS_1352'] = True
    # doc11['02043_18_04939_WHOLENESS_1353'] = True
    # doc11['02043_18_04939_WHOLENESS_1354'] = True
    # doc11['02043_18_04939_WHOLENESS_1355'] = True
    # doc11['02043_18_04939_WHOLENESS_1356'] = True
    # doc11['02043_18_04939_WHOLENESS_1357'] = True
    # doc11['02043_18_04939_WHOLENESS_1358'] = True
    # doc11['02043_18_04939_WHOLENESS_1359'] = True
    # doc11['02043_18_04939_WHOLENESS_1360'] = False
    # doc11['02043_18_04939_WHOLENESS_1361'] = True
    # doc11['02043_18_04939_WHOLENESS_1400'] = True
    # doc11['02043_18_04939_WHOLENESS_1401'] = True
    # doc11['02043_18_04939_WHOLENESS_1402'] = True
    # doc11['02043_18_04939_WHOLENESS_1403'] = False
    # doc11['02043_18_04939_WHOLENESS_1430'] = True
    # doc11['02043_18_04939_WHOLENESS_1431'] = True
    # doc11['02043_18_04939_WHOLENESS_1432'] = True
    # doc11['02043_18_04939_WHOLENESS_1433'] = True
    # doc11['02043_18_04939_WHOLENESS_1434'] = True
    # doc11['02043_18_04939_WHOLENESS_1435'] = True
    # doc11['02043_18_04939_WHOLENESS_1450'] = True
    # doc11['02043_18_04939_WHOLENESS_1451'] = True
    # doc11['02043_18_04939_WHOLENESS_1452'] = True
    # doc11['02043_18_04939_WHOLENESS_1453'] = True
    # doc11['02043_18_04939_WHOLENESS_1454'] = True
    # doc11['02043_18_04939_WHOLENESS_1455'] = False
    # doc11['02043_18_04939_WHOLENESS_1456'] = False
    # doc11['02043_18_04939_WHOLENESS_1500'] = True
    # doc11['02043_18_04939_WHOLENESS_1501'] = True
    # doc11['02043_18_04939_WHOLENESS_1550'] = True
    # doc11['02043_18_04939_WHOLENESS_1551'] = True
    # doc11['02043_18_04939_WHOLENESS_1552'] = False
    # return doc11


def DOC12Coherence(ExcelApp, TSDApp, DOC12Path):
    try:
        DOC12 = ExcelApp.Workbooks.Open(DOC12Path)
    except:
        TSDApp.tab1.textbox.setText("ERROR: when trying to parse the Plan type TSD Systeme file " + DOC12Path.split('/')[-1])
    sheets = DOC12.Sheets
    sheetNames = list()
    for sheet in sheets:
        sheetNames.append(sheet.Name.strip().casefold())

    dict12 = []

    if "table" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("table") + 1
        workSheet = DOC12.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.tableLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict12.append(dictionary)

    if "data trouble codes" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("data trouble codes") + 1
        workSheet = DOC12.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.codeLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict12.append(dictionary)

    if "read data and io control" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("read data and io control") + 1
        workSheet = DOC12.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.ReadDataIOLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict12.append(dictionary)

    if "not embedded diagnosis" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("not embedded diagnosis") + 1
        workSheet = DOC12.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.NotEmbDiagLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict12.append(dictionary)

    if "technical effect" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("technical effect") + 1
        workSheet = DOC12.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.TechEffLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict12.append(dictionary)

    if "customer effect" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("customer effect") + 1
        workSheet = DOC12.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.EffClientsLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict12.append(dictionary)

    if "feared events" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("feared events") + 1
        workSheet = DOC12.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.FearedEventLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict12.append(dictionary)

    if "parts" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("parts") + 1
        workSheet = DOC12.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.PartsLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict12.append(dictionary)

    if "variant" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("variant") + 1
        workSheet = DOC12.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.VariantLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict12.append(dictionary)

    if "situation" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("situation") + 1
        workSheet = DOC12.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.SituationLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict12.append(dictionary)

    if "degraded mode" in sheetNames:
        col = 1
        row = 0
        index = sheetNames.index("Degraded mode") + 1
        workSheet = DOC12.Sheets(index)
        for index in range(1, 20):
            if workSheet.Cells(index, 1).Value == "Requirement N°":
                row = index
                break
        if row > 0:
            for index in range(col + 1, TSDApp.WorkbookStats.DegradedModeLastCol + 1):
                dictionary = {}
                if workSheet.Cells(row, index).Value is not None:
                    dictionary["name"] = workSheet.Cells(row, index).Value
                    dictionary["value"] = workSheet.Cells(row + 1, index).Value
                    if dictionary["value"] == "Oui":
                        dictionary["value"] = True
                    else:
                        dictionary["value"] = False
                    dict12.append(dictionary)

    return dict12

    # doc12 = {}
    # doc12['02043_18_04939_WHOLENESS_1900'] = True
    # doc12['02043_18_04939_WHOLENESS_1901'] = True
    # doc12['02043_18_04939_WHOLENESS_1902'] = True
    # doc12['02043_18_04939_WHOLENESS_1903'] = True
    # doc12['02043_18_04939_WHOLENESS_1904'] = True
    # doc12['02043_18_04939_WHOLENESS_1905'] = True
    # doc12['02043_18_04939_WHOLENESS_1906'] = True
    # doc12['02043_18_04939_WHOLENESS_1907'] = True
    # doc12['02043_18_04939_WHOLENESS_1908'] = True
    # doc12['02043_18_04939_WHOLENESS_1909'] = True
    # doc12['02043_18_04939_WHOLENESS_1910'] = False
    # doc12['02043_18_04939_WHOLENESS_1911'] = True
    # doc12['02043_18_04939_WHOLENESS_1912'] = True
    # doc12['02043_18_04939_WHOLENESS_1913'] = True
    # doc12['02043_18_04939_WHOLENESS_1914'] = True
    # doc12['02043_18_04939_WHOLENESS_1915'] = False
    # doc12['02043_18_04939_WHOLENESS_1916'] = False
    # doc12['02043_18_04939_WHOLENESS_1917'] = True
    # doc12['02043_18_04939_WHOLENESS_1918'] = True
    # doc12['02043_18_04939_WHOLENESS_1919'] = True
    # doc12['02043_18_04939_WHOLENESS_1920'] = True
    # doc12['02043_18_04939_WHOLENESS_1921'] = True
    # doc12['02043_18_04939_WHOLENESS_1922'] = False
    # doc12['02043_18_04939_WHOLENESS_1923'] = True
    # doc12['02043_18_04939_WHOLENESS_1924'] = False
    # doc12['02043_18_04939_WHOLENESS_1925'] = False
    # doc12['02043_18_04939_WHOLENESS_1926'] = False
    # doc12['02043_18_04939_WHOLENESS_1927'] = False
    # doc12['02043_18_04939_WHOLENESS_1928'] = True
    # doc12['02043_18_04939_WHOLENESS_1950'] = True
    # doc12['02043_18_04939_WHOLENESS_1951'] = True
    # doc12['02043_18_04939_WHOLENESS_1952'] = True
    # doc12['02043_18_04939_WHOLENESS_1953'] = True
    # doc12['02043_18_04939_WHOLENESS_1954'] = True
    # doc12['02043_18_04939_WHOLENESS_1955'] = True
    # doc12['02043_18_04939_WHOLENESS_1956'] = True
    # doc12['02043_18_04939_WHOLENESS_1957'] = True
    # doc12['02043_18_04939_WHOLENESS_1958'] = True
    # doc12['02043_18_04939_WHOLENESS_1959'] = True
    # doc12['02043_18_04939_WHOLENESS_1960'] = False
    # doc12['02043_18_04939_WHOLENESS_1961'] = True
    # doc12['02043_18_04939_WHOLENESS_1962'] = True
    # doc12['02043_18_04939_WHOLENESS_1963'] = False
    # doc12['02043_18_04939_WHOLENESS_1964'] = True
    # doc12['02043_18_04939_WHOLENESS_1965'] = True
    # doc12['02043_18_04939_WHOLENESS_1966'] = True
    # doc12['02043_18_04939_WHOLENESS_1967'] = True
    # doc12['02043_18_04939_WHOLENESS_1968'] = True
    # doc12['02043_18_04939_WHOLENESS_1969'] = True
    # doc12['02043_18_04939_WHOLENESS_2000'] = True
    # doc12['02043_18_04939_WHOLENESS_2001'] = True
    # doc12['02043_18_04939_WHOLENESS_2002'] = True
    # doc12['02043_18_04939_WHOLENESS_2003'] = True
    # doc12['02043_18_04939_WHOLENESS_2004'] = True
    # doc12['02043_18_04939_WHOLENESS_2005'] = True
    # doc12['02043_18_04939_WHOLENESS_2006'] = True
    # doc12['02043_18_04939_WHOLENESS_2007'] = True
    # doc12['02043_18_04939_WHOLENESS_2008'] = True
    # doc12['02043_18_04939_WHOLENESS_2009'] = True
    # doc12['02043_18_04939_WHOLENESS_2010'] = True
    # doc12['02043_18_04939_WHOLENESS_2011'] = True
    # doc12['02043_18_04939_WHOLENESS_2050'] = True
    # doc12['02043_18_04939_WHOLENESS_2051'] = True
    # doc12['02043_18_04939_WHOLENESS_2052'] = True
    # doc12['02043_18_04939_WHOLENESS_2053'] = True
    # doc12['02043_18_04939_WHOLENESS_2054'] = True
    # doc12['02043_18_04939_WHOLENESS_2055'] = True
    # doc12['02043_18_04939_WHOLENESS_2056'] = True
    # doc12['02043_18_04939_WHOLENESS_2060'] = True
    # doc12['02043_18_04939_WHOLENESS_2061'] = True
    # doc12['02043_18_04939_WHOLENESS_2062'] = True
    # doc12['02043_18_04939_WHOLENESS_2070'] = True
    # doc12['02043_18_04939_WHOLENESS_2071'] = True
    # doc12['02043_18_04939_WHOLENESS_2072'] = True
    # doc12['02043_18_04939_WHOLENESS_2080'] = True
    # doc12['02043_18_04939_WHOLENESS_2081'] = True
    # doc12['02043_18_04939_WHOLENESS_2082'] = True
    # doc12['02043_18_04939_WHOLENESS_2083'] = False
    # doc12['02043_18_04939_WHOLENESS_2084'] = False
    # doc12['02043_18_04939_WHOLENESS_2090'] = True
    # doc12['02043_18_04939_WHOLENESS_2091'] = True
    # doc12['02043_18_04939_WHOLENESS_2092'] = True
    # doc12['02043_18_04939_WHOLENESS_2100'] = True
    # doc12['02043_18_04939_WHOLENESS_2101'] = True
    # doc12['02043_18_04939_WHOLENESS_2102'] = True
    # doc12['02043_18_04939_WHOLENESS_2110'] = True
    # doc12['02043_18_04939_WHOLENESS_2111'] = True
    # doc12['02043_18_04939_WHOLENESS_2112'] = False
    # doc12['02043_18_04939_WHOLENESS_2120'] = True
    # doc12['02043_18_04939_WHOLENESS_2121'] = True
    # return doc12

