import TSD_Checker_V6_0
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


    cnt = 0
    for index1 in range(0, rows):
        for index2 in range(0, cols):
            if workSheet.cell(index1, index2).value.casefold().strip() == 'Nom CF /\nNom CO PLM (CF_CO)'.casefold():
                refRow = index1
                refCol = index2
                cnt += 1
            if workSheet.cell(index1, index2).value.casefold().strip() == 'EC name /\nDesignationFR CF PLM'.casefold():
                refRowEC = index1
                refColEC = index2
                cnt += 1
            if workSheet.cell(index1, index2).value.casefold().strip() == 'Values'.casefold():
                refRowVal = index1
                refColVal = index2
                cnt += 1
            if cnt == 3:
                break
        if cnt == 3:
            break

    for index in range(refRow + 1, rows):
        if workSheet.cell(index, refCol).value is not None and workSheet.cell(index, refCol).value != "":
            final_list.append(workSheet.cell(index,refCol).value.strip())

    ECs = dict()
    for index in range(refRow + 1, rows):
        if workSheet.cell(index, refColEC).value != "" and workSheet.cell(index, refColVal).value == "":
            values = []
            for index1 in range(index + 1, rows):
                if workSheet.cell(index1, refColEC).value == workSheet.cell(index, refColEC).value:
                    values.append(workSheet.cell(index1,refColVal).value)
                else:
                    break

            EC = workSheet.cell(index, refColEC).value
            ECs[EC] = values

    return final_list, ECs



def DOC8Parser(TSDApp ,ExcelApp, DOC8Path):
    try:
        DOC8 = xlrd.open_workbook(DOC8Path, on_demand=True)
    except:
        TSDApp.tab1.textbox.setText("ERROR: when trying to parse the CESARE file " + DOC8Path.split('/')[-1])
        return

    sheet_list = DOC8.sheet_names()
    cnt = 0
    index_1 = -1
    index_2 = -1
    for sheet in sheet_list:
        cnt += 1
        if "sous familles Cesare" in sheet:
            index_1 = cnt - 1
        if "ECU Exception" in sheet:
            index_2 = cnt - 1

    workSheet = DOC8.sheet_by_index(index_1)
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

    workSheet2 = DOC8.sheet_by_index(index_2)
    rows2 = workSheet2.nrows
    cols2 = workSheet2.ncols
    flag = False

    for index1 in range(0,rows2):
        for index2 in range(0,cols2):
            try:
                if workSheet2.cell(index1,index2).value.casefold().strip() == 'Nom de la sous famille'.casefold():
                    refRow = index1
                    refCol = index2
                    flag = True
                    break
            except:
                pass
        if flag is True:
            break

    for index in range(refRow + 1, rows2):
        if workSheet2.cell(index, refCol).value is not None and workSheet2.cell(index, refCol).value != "":
            final_list.append(workSheet2.cell(index, refCol).value.replace(u'\xa0', u''))

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
