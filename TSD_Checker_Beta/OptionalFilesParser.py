import TSD_Checker_V5_0
from lxml import etree, objectify
import xlrd

def DOC9Parser(TSDApp, ExcelApp, DOC9Path):
    try:
        DOC9 = xlrd.open_workbook(DOC9Path, on_demand=True)
    except:
        TSDApp.tab1.textbox.setText("ERROR: when trying to parse the criticity file " + DOC9Path.split('/')[-1])
        return

    workSheet = DOC9.sheet_by_name("Configuration")
    rows = workSheet.nrows
    cols = workSheet.ncols
    refRow = 0
    refCol = 0
    flag = False

    for index1 in range(0, rows ):
        for index2 in range(0, cols ):
            if workSheet.cell(index1, index2).value.casefold().strip() == "Requirements".casefold():
                refRow = index1
                refCol = index2
                flag = True
                break
        if flag == True:
            break

    if flag is True:
        for index in range(0,cols):
            if workSheet.cell(refRow, index).value.casefold().strip() == "previsionnal":
                prev = index
            elif workSheet.cell(refRow, index).value.casefold().strip() == "consolidated":
                cons = index
            elif workSheet.cell(refRow, index).value.casefold().strip() == "validated":
                val = index


    DOC9Dict = dict()
    for index in range(refRow + 1, rows):
        if workSheet.cell(index1, index2).value.strip() is not None:
            tempDict = dict()
            tempDict["previsional"] = workSheet.cell(index, prev).value.strip()
            tempDict["consolidated"] = workSheet.cell(index, cons).value.strip()
            tempDict["validated"] = workSheet.cell(index, val).value.strip()
            try:
                testName = "Test_" + workSheet.cell(index, refCol).value.strip()
                DOC9Dict[testName] = tempDict
            except:
                pass

    return DOC9Dict
    # try:
    #     DOC9 = ExcelApp.Workbooks.Open(DOC9Path)
    # except:
    #     TSDApp.tab1.textbox.setText("ERROR: when trying to parse the criticity file " + DOC9Path.split('/')[-1])
    #     return
    # workSheet = DOC9.Sheets("Configuration")
    # workSheetRangeValuesTuple = workSheet.UsedRange.Value
    # fillDictFlag = False
    # DOC9Dict = dict()
    #
    # for rowTuple in workSheetRangeValuesTuple:
    #     if fillDictFlag == True:
    #         tempDict = dict()
    #         tempDict["previsional"] = rowTuple[prevCol]
    #         tempDict["consolidated"] = rowTuple[consCol]
    #         tempDict["validated"] = rowTuple[valiCol]
    #         try:
    #             testName = "Test_" + rowTuple[testCol].strip()
    #             DOC9Dict[testName] = tempDict
    #         except:
    #             pass
    #
    #     elif "Requirements" in rowTuple:
    #         testCol = rowTuple.index("Requirements")
    #         prevCol = rowTuple.index("Previsionnal")
    #         consCol = rowTuple.index("Consolidated")
    #         valiCol = rowTuple.index("Validated")
    #         fillDictFlag = True
    #
    # DOC9.Close()
    #
    # return DOC9Dict

def DOC13Parser(TSDApp, ExcelApp, DOC13Path):
    try:
        DOC13 = xlrd.open_workbook(DOC13Path, on_demand=True)
    except:
        TSDApp.tab1.textbox.setText("ERROR: when trying to parse the diversity referential file " + DOC13Path.split('/')[-1])
        return

    workSheet = DOC13.sheet_by_name("Liste EC")
    rows = workSheet.nrows
    cols = workSheet.ncols
    flag = False
    final_list = []

    for index1 in range(0, rows):
        for index2 in range(0, cols):
            if workSheet.cell(index1, index2).value.casefold().strip() == 'Nom CF /\nNom CO PLM (CF_CO)'.casefold():
                refRow = index1
                refCol = index2
                flag = True
                break
        if flag is True:
            break

    for index in range(refRow + 1, rows):
        if workSheet.cell(index, refCol).value is not None and workSheet.cell(index, refCol).value != "":
            final_list.append(workSheet.cell(index,refCol).value.strip())

    return final_list

    # try:
    #     DOC13 = ExcelApp.Workbooks.Open(DOC13Path)
    # except:
    #     TSDApp.tab1.textbox.setText("ERROR: when trying to parse the diversity referential file " + DOC13Path.split('/')[-1])
    #     return
    # workSheet = DOC13.Sheets("Liste EC")
    # workSheetRangeValuesTuple = workSheet.UsedRange.Value
    # testCol = 0
    # temp_list = []
    # final_list = []
    #
    # for rowTuple in workSheetRangeValuesTuple:
    #     if 'Nom CF /\nNom CO PLM (CF_CO)' in rowTuple:
    #         testCol = rowTuple.index("Nom CF /\nNom CO PLM (CF_CO)")
    #         break
    # for rowTuple in workSheetRangeValuesTuple:
    #     temp_list.append(rowTuple[testCol])
    #
    # for elem in temp_list:
    #     if elem not in [None, "", 'Nom CF /\nNom CO PLM (CF_CO)']:
    #         final_list.append(elem)
    # return final_list


def DOC8Parser(TSDApp ,ExcelApp, DOC8Path):
    try:
        DOC8 = xlrd.open_workbook(DOC8Path, on_demand=True)
    except:
        TSDApp.tab1.textbox.setText("ERROR: when trying to parse the CESARE file " + DOC8Path.split('/')[-1])
        return

    sheet_list = DOC8.sheet_names()
    cnt = 0
    index = -1
    for sheet in sheet_list:
        cnt += 1
        if "sous familles Cesare" in sheet:
            index = cnt - 1
            break

    workSheet = DOC8.sheet_by_index(index)
    rows = workSheet.nrows
    cols = workSheet.ncols
    flag = False
    final_list = []

    for index1 in range(0, rows):
        for index2 in range(0, cols):
            try:
                if workSheet.cell(index1, index2).value.casefold().strip() == 'Nom de la sous famille'.casefold():
                    refRow = index1
                    refCol = index2
                    flag = True
                    break
            except:
                pass
        if flag is True:
            break

    for index in range(refRow + 1, rows):
        if workSheet.cell(index, refCol).value is not None and workSheet.cell(index, refCol).value != "":
            final_list.append(workSheet.cell(index, refCol).value.replace(u'\xa0', u''))

    return final_list

    # try:
    #     DOC8 = ExcelApp.Workbooks.Open(DOC8Path)
    # except:
    #     TSDApp.tab1.textbox.setText("ERROR: when trying to parse the CESARE file " + DOC8Path.split('/')[-1])
    #     return
    # cnt = 0
    # temp = DOC8.Sheets
    # for sheet in temp:
    #     cnt = cnt + 1
    #     if "sous familles Cesare" in sheet.Name:
    #         index = cnt
    #         break
    # workSheet = DOC8.Sheets(index)
    # workSheetRangeValuesTuple = workSheet.UsedRange.Value
    # testCol = 0
    # temp_list = []
    # final_list = []
    #
    # for rowTuple in workSheetRangeValuesTuple:
    #     #temp = tuple(tuple(b.strip() for b in a) for a in rowTuple)
    #     if '\xa0Nom de la sous famille\xa0' in rowTuple:
    #         testCol = rowTuple.index("\xa0Nom de la sous famille\xa0")
    #         break
    # for rowTuple in workSheetRangeValuesTuple:
    #     temp_list.append(rowTuple[testCol])
    #
    # for elem in temp_list:
    #     if elem not in [None, "", '\xa0Nom de la sous famille\xa0']:
    #         final_list.append(elem.replace(u'\xa0', u''))
    # return final_list

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
