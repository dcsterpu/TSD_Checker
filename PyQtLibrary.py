import sys
from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QPushButton, QAction, QMenu, QApplication, QMessageBox, QComboBox, QLabel, QLineEdit, QInputDialog, QFileDialog
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot, QObject
from PyQt5 import QtCore, QtGui, QtWidgets, QtWebEngineWidgets
import pandas as pd
import openpyxl as xl



class Application(QWidget):

    def __init__(self):
        super().__init__()
        self.left = 800
        self.top = 800
        self.width = 800
        self.height = 650
        self.initUI()
        self.excelExtension = str()

    def openFileNameDialog1(self):
        fileName1, _filter = QtWidgets.QFileDialog.getOpenFileName(self, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.myTextBox1.setText(fileName1)
        self.textbox.setText("next file")

    def openFileNameDialog2(self):
        fileName2, _filter = QtWidgets.QFileDialog.getOpenFileName(self, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.myTextBox2.setText(fileName2)



    def openFileNameDialog3(self):
        fileName3, _filter = QtWidgets.QFileDialog.getOpenFileName(self, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.myTextBox3.setText(fileName3)


    def openFileNameDialog4(self):
        fileName4, _filter = QtWidgets.QFileDialog.getOpenFileName(self, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.myTextBox4.setText(fileName4)


    def openFileNameDialog5(self):
        fileName5, _filter = QtWidgets.QFileDialog.getOpenFileName(self, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.myTextBox5.setText(fileName5)


    def openFileNameDialog6(self):
        fileName6, _filter = QtWidgets.QFileDialog.getOpenFileName(self, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.myTextBox6.setText(fileName6)


    def openFileNameDialog7(self):
        fileName7, _filter = QtWidgets.QFileDialog.getOpenFileName(self, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.myTextBox7.setText(fileName7)


    def openFileNameDialog8(self):
        fileName8, _filter = QtWidgets.QFileDialog.getOpenFileName(self, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.myTextBox8.setText(fileName8)

    def openFileNameDialog9(self):
        fileName9, _filter = QtWidgets.QFileDialog.getOpenFileName(self, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.myTextBox9.setText(fileName9)

    def openFileNameDialog10(self):
        fileName10, _filter = QtWidgets.QFileDialog.getOpenFileName(self, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.myTextBox10.setText(fileName10)


    def buttonClicked(self):
        fileName = self.myTextBox1.toPlainText()
        self.readExcel(fileName)
        fileName = self.myTextBox2.toPlainText()
        self.writeExcel(fileName)

    def readExcel(self,fileName):
        if fileName.endswith(".xls"):
            import xlrd as reader
            def openExcel(filename):
                return reader.open_workbook(filename)

            def getSheetNames(workbook):
                return workbook.sheet_names()
            self.excelExtension = ".xls"
            self.readXLSFileName = fileName

        elif fileName.endswith((".xlsx", ".xlsm")):
            import openpyxl.reader as reader
            from openpyxl import load_workbook
            def openExcel(filename):
                if filename.endswith(".xlsm"):
                    return load_workbook(filename, keep_vba=True)
                else:
                    return  load_workbook(filename)

            def getSheetNames(workbook):
                return workbook.sheetnames
            self.excelExtension = fileName[-5:]

        else:
            self.textbox.setText("Invalid file format")
            return

        workBook = openExcel(fileName)
        self.excelFile = workBook
        sheetNamesList = getSheetNames(workBook)
        sheetNamesString = str()
        for sheetNames in sheetNamesList:
            sheetNamesString = sheetNamesString + " " + sheetNames

        self.textbox.setText(sheetNamesString)
        return

    def writeExcel(self, fileName):
        if fileName.endswith(".xls") and self.excelExtension==".xls":
            import pandas as pd
            from xlutils.copy import copy
            import xlrd
            def editExcel(workbook,fileName):
                worksheetDF = pd.read_excel(self.excelFile, engine="xlrd", sheet_name="Sheet1")

                worksheetDF = pd.concat([worksheetDF, pd.DataFrame({"A": 1})], ignore_index=True)
                excelFileNew = pd.ExcelWriter(fileName, engine="xlwt")
                workbook_xlwt = copy(self.excelFile)
                sheet1_index = workbook_xlwt.sheet_index("sheet1")
                workbook_xlwt._Workbook__worksheet_idx_from_name.pop("sheet1")
                workbook_xlwt._Workbook__worksheets.pop(sheet1_index)
                excelFileNew.book = workbook_xlwt
                worksheetDF.to_excel(excelFileNew, sheet_name="Sheet1", index=False, header=False)

                excelFileNew.save()

        elif fileName.endswith(".xlsx") and self.excelExtension == ".xlsx":
            import openpyxl as xl
            def editExcel(workbook,fileName):
                worksheet = workbook["Sheet1"]
                worksheet.append(["blabla"])
                workbook.save(fileName)
                workbook.close()
        elif fileName.endswith(".xlsm") and self.excelExtension == ".xlsm":
            import openpyxl as xl
            def editExcel(workbook,fileName):
                worksheet = workbook["Sheet1"]
                worksheet.append(["blabla"])
                workbook.save(fileName)
                workbook.close()

        else:
            self.textbox.setText("No File")
            return

        editExcel(self.excelFile, fileName)

    def initUI(self):

    # Create a textbox
        message = "message"
        self.textbox = QtWidgets.QTextEdit(self)
        self.textbox.setText(message)
        self.textbox.move(10, 500)
        self.textbox.resize(700, 60)
        self.textbox.setReadOnly(True)



    #Create a drop down list
        self.lbl = QLabel("Check level", self)

        combo = QComboBox(self)
        combo.addItem("   Option1   ")
        combo.addItem("   Option2   ")
        combo.addItem("   Option3   ")
        combo.addItem("   Option4   ")
        combo.addItem("   Option5   ")
        combo.resize(508, 20.4)  #rezise the drop down list
        combo.move(200, 430)
        self.lbl.move(5, 436)
        combo.activated[str].connect(self.onActivated)

    # Create a drop down list
        self.lbl = QLabel("Project name", self)

        combo = QComboBox(self)
        combo.addItem("   Option1   ")
        combo.addItem("   Option2   ")
        combo.addItem("   Option3   ")
        combo.addItem("   Option4   ")
        combo.addItem("   Option5   ")
        combo.resize(508, 20.4)  # rezise the drop down list
        combo.move(200, 460)
        self.lbl.move(5, 466)
        combo.activated[str].connect(self.onActivated)

        self.setGeometry(self.left, self.top, self.width, self.height)
        self.setWindowTitle('TSD Checker')


        #File Selectiom Dialog1
        self.lbl2 = QLabel("TSD File:", self)
        self.lbl2.move(5,15)
        self.myTextBox1 = QtWidgets.QTextEdit(self)
        self.myTextBox1.resize(460, 25)
        self.myTextBox1.move(200, 10)
        self.myTextBox1.setReadOnly(True)
        button1 = QPushButton('...',self)
        button1.clicked.connect(self.openFileNameDialog1)
        button1.move(660, 10)
        button1.resize(45,22)

    # File Selectiom Dialog2
        self.lbl3 = QLabel("TSD vehicle Function file:", self)
        self.lbl3.move(5, 45)
        self.myTextBox2 = QtWidgets.QTextEdit(self)
        self.myTextBox2.resize(460, 25)
        self.myTextBox2.move(200, 40)
        self.myTextBox2.setReadOnly(True)
        button2 = QPushButton('...', self)
        button2.clicked.connect(self.openFileNameDialog2)
        button2.move(660, 40)
        button2.resize(45, 22)

    # File Selectiom Dialog3
        self.lbl4 = QLabel("TSD system file:", self)
        self.lbl4.move(5, 75)
        self.myTextBox3 = QtWidgets.QTextEdit(self)
        self.myTextBox3.resize(460, 25)
        self.myTextBox3.move(200, 70)
        self.myTextBox3.setReadOnly(True)
        button3 = QPushButton('...', self)
        button3.clicked.connect(self.openFileNameDialog3)
        button3.move(660, 70)
        button3.resize(45, 22)

    # File Selectiom Dialog4
        self.lbl5 = QLabel("TSD configuration file:", self)
        self.lbl5.move(5, 105)
        self.myTextBox4 = QtWidgets.QTextEdit(self)
        self.myTextBox4.resize(460, 25)
        self.myTextBox4.move(200, 100)
        self.myTextBox4.setReadOnly(True)


        self.link1 = QLabel('''<a href='http://www.google.com'>Google</a>''', self)
        self.link1.setOpenExternalLinks(True)
        self.link1.move(200, 130)


        button4 = QPushButton('...', self)
        button4.clicked.connect(self.openFileNameDialog4)
        button4.move(660, 100)
        button4.resize(45, 22)

    # File Selectiom Dialog5
        self.lbl6 = QLabel("Famille/Sous-Famille list export(CESARE):", self)
        self.lbl6.move(5, 165)
        self.myTextBox5 = QtWidgets.QTextEdit(self)
        self.myTextBox5.resize(460, 25)
        self.myTextBox5.move(200, 160)
        self.myTextBox5.setReadOnly(True)

        self.link2 = QLabel('''<a href='http://www.google.com'>Google</a>''', self)
        self.link2.setOpenExternalLinks(True)
        self.link2.move(200, 190)

        button5 = QPushButton('...', self)
        button5.clicked.connect(self.openFileNameDialog5)
        button5.move(660, 160)
        button5.resize(45, 22)

    # File Selectiom Dialog6
        self.lbl7 = QLabel("Customer effect file:", self)
        self.lbl7.move(5, 225)
        self.myTextBox6 = QtWidgets.QTextEdit(self)
        self.myTextBox6.resize(460, 25)
        self.myTextBox6.move(200, 220)
        self.myTextBox6.setReadOnly(True)

        self.link3 = QLabel('''<a href='http://www.google.com'>Google</a>''', self)
        self.link3.setOpenExternalLinks(True)
        self.link3.move(200, 250)


        button6 = QPushButton('...', self)
        button6.clicked.connect(self.openFileNameDialog6)
        button6.move(660,220)
        button6.resize(45, 22)

    # File Selectiom Dialog7
        self.lbl8 = QLabel("AMDEC:", self)
        self.lbl8.move(5, 285)
        self.myTextBox7 = QtWidgets.QTextEdit(self)
        self.myTextBox7.resize(460, 25)
        self.myTextBox7.move(200, 280)
        self.myTextBox7.setReadOnly(True)
        button7 = QPushButton('...', self)
        button7.clicked.connect(self.openFileNameDialog7)
        button7.move(660, 280)
        button7.resize(45, 22)

    # File Selectiom Dialog8
        self.lbl9 = QLabel("export MedialecMatrice:", self)
        self.lbl9.move(5, 317)
        self.myTextBox8 = QtWidgets.QTextEdit(self)
        self.myTextBox8.resize(460, 25)
        self.myTextBox8.move(200, 310)
        self.myTextBox8.setReadOnly(True)
        button8 = QPushButton('...', self)
        button8.clicked.connect(self.openFileNameDialog8)
        button8.move(660, 310)
        button8.resize(45, 22)

    # File Selectiom Dialog9
        self.lbl10 = QLabel("Diversity management file:", self)
        self.lbl10.move(5, 347)
        self.myTextBox9 = QtWidgets.QTextEdit(self)
        self.myTextBox9.resize(460, 25)
        self.myTextBox9.move(200, 340)
        self.myTextBox9.setReadOnly(True)

        self.link4 = QLabel('''<a href='http://docinfogroupe.inetpsa.com/ead/doc/ref.02016_11_04964/v.vc/fiche'>http://docinfogroupe.inetpsa.com/ead/doc/ref.02016_11_04964/v.vc/fiche</a>''', self)
        self.link4.setOpenExternalLinks(True)
        self.link4.move(200, 368)

        button9 = QPushButton('...', self)
        button9.clicked.connect(self.openFileNameDialog9)
        button9.move(660, 340)
        button9.resize(45, 22)

    # File Selectiom Dialog10
        self.lbl11 = QLabel("Diagnostic matrix file:", self)
        self.lbl11.move(5, 397)
        self.myTextBox10 = QtWidgets.QTextEdit(self)
        self.myTextBox10.resize(460, 25)
        self.myTextBox10.move(200, 390)
        self.myTextBox10.setReadOnly(True)
        button10 = QPushButton('...', self)
        button10.clicked.connect(self.openFileNameDialog10)
        button10.move(660, 390)
        button10.resize(45, 22)


    # Check button
        button = QPushButton('Check', self)
        button.move(350, 590)
        button.resize(90,25)
        button.clicked.connect(self.buttonClicked)
        button.setStyleSheet('QPushButton {background-color: white; color: black;}')

        self.show()


    def onActivated(self, text):
        self.lbl.setText(text)
        self.lbl.adjustSize()



if __name__ == '__main__':
    app = QApplication(sys.argv)
    apel = Application()
    myQLabel = QLabel()
    sys.exit(app.exec_())



