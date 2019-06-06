import sys
from PyQt5.QtWidgets import QWidget, QPushButton, QApplication, QComboBox, QLabel, QLineEdit,  QTabWidget, QVBoxLayout, QProgressBar, QRadioButton
from PyQt5 import QtCore, QtWidgets
import win32com.client as win32
import os
import io
import requests
from ctypes import windll
import OptionalFilesParser
import GeneralStructureTester
import FileMeasure
from timeit import default_timer as timer
import ExcelEdit
import WholenessTester
import Coherence_checksTester
import IndicatorTester
import time
import xlwt
import xlrd
from xlrd import open_workbook
from xlwt import easyxf

appName = "TSD Checker V4.0"
pBarIncrement = 100/159

class Application(QWidget):

    def __init__(self):
        super().__init__()
        self.left = 200
        self.top = 200
        self.width = 900
        self.height = 600
        self.tabs = QTabWidget()
        self.tab1 = QWidget()
        self.tab2 = QWidget()
        self.DOC8Link = '''https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02043_18_05471/v.vc/pj'''
        self.DOC9Link = '''https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02043_18_05474/v.vc/pj'''
        self.DOC7Link = '''https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02043_18_05499/v.vc/pj'''
        self.DOC13Link = '''https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02016_11_04964/v.vc/pj'''
        self.DOC3Link = '''https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.AEEV_IAEE07_0033/v.vc/pj'''
        self.DOC4Link = '''https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02043_12_01665/v.vc/pj'''
        self.DOC5Link = '''https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02043_12_01666/v.vc/pj'''
        self.DOC14Link = "https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02043_19_00392/v.vc/pj"
        self.tabs.addTab(self.tab1, "TSD Checker")
        self.tabs.addTab(self.tab2, "Options")
        self.initUI(self.tab1)
        self.initUIOptions(self.tab2)
        self.layout = QVBoxLayout()
        self.layout.addWidget(self.tabs)
        self.setLayout(self.layout)
        self.setWindowTitle(appName)
        self.coverage = ""
        self.convergence = ""
        self.status = ""
        self.username = os.environ['USERNAME']
        self.fileFolder = "C:/Users/" + self.username + "/AppData/Local/Temp/TSD_Checker/"
        self.pBarValue = 0

        self.criticity_blocking = 0
        self.criticity_warning = 0
        self.criticity_information = 0
        self.start_time = 0
        self.end_time = 0
        self.opening_time = 0

        self.return_list = []

    def ToggleLink(self):
        if self.tab2.RadioButtonInternet.isChecked() == True:
            self.DOC8Link = '''https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02043_18_05471/v.vc/pj'''
            self.DOC9Link = '''https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02043_18_05474/v.vc/pj'''
            self.DOC7Link = '''https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02043_18_05499/v.vc/pj'''
            self.DOC13Link = '''https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02016_11_04964/v.vc/pj'''
            self.DOC3Link = '''https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.AEEV_IAEE07_0033/v.vc/pj'''
            self.DOC4Link = '''https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02043_12_01665/v.vc/pj'''
            self.DOC5Link = '''https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02043_12_01666/v.vc/pj'''
            self.DOC9Link = "https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02043_18_05474/v.vc/pj"
            self.DOC14Link = "https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02043_19_00392/v.vc/pj"
            self.tab2.link2.setText('''<a href=''' + self.DOC8Link + '''>DocInfo Reference: 02043_18_05471</a>''')
            self.tab2.link1.setText('''<a href=''' + self.DOC9Link + '''>DocInfo Reference: 02043_18_05474</a>''')
            self.tab2.link3.setText('''<a href=''' + self.DOC7Link + '''>DocInfo Reference: 02043_18_05499</a>''')
            self.tab2.link4.setText('''<a href=''' + self.DOC13Link + '''>DocInfo Reference: 02016_11_04964</a>''')
        elif self.tab2.RadioButtonIntranet.isChecked() == True:
            self.DOC8Link = "http://docinfogroupe.inetpsa.com/ead/doc/ref.02043_18_05471/v.vc/pj"
            self.DOC9Link = "http://docinfogroupe.inetpsa.com/ead/doc/ref.02043_18_05474/v.vc/pj"
            self.DOC7Link = "http://docinfogroupe.inetpsa.com/ead/doc/ref.02043_18_05499/v.vc/pj"
            self.DOC13Link = "http://docinfogroupe.inetpsa.com/ead/doc/ref.02016_11_04964/v.vc/pj"
            self.DOC3Link = '''http://docinfogroupe.inetpsa.com/ead/doc/ref.AEEV_IAEE07_0033/v.vc/pj'''
            self.DOC4Link = '''http://docinfogroupe.inetpsa.com/ead/doc/ref.02043_12_01665/v.vc/pj'''
            self.DOC5Link = '''http://docinfogroupe.inetpsa.com/ead/doc/ref.02043_12_01666/v.vc/pj'''
            self.DOC9Link = "http://docinfogroupe.inetpsa.com/ead/doc/ref.02043_18_05474/v.vc/pj"
            self.DOC14Link = "http://docinfogroupe.inetpsa.com/ead/doc/ref.02043_19_00392/v.vc/pj"
            self.tab2.link2.setText('''<a href=''' + self.DOC8Link + '''>DocInfo Reference: 02043_18_05471</a>''')
            self.tab2.link1.setText('''<a href=''' + self.DOC9Link + '''>DocInfo Reference: 02043_18_05474</a>''')
            self.tab2.link3.setText('''<a href=''' + self.DOC7Link + '''>DocInfo Reference: 02043_18_05499</a>''')
            self.tab2.link4.setText('''<a href=''' + self.DOC13Link + '''>DocInfo Reference: 02016_11_04964</a>''')

    def openFileNameDialog1(self):
        fileName1, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab1, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab1.myTextBox1.setText(fileName1)
        # self.tab1.textbox.setText("next file")

    def openFileNameDialog2(self):
        fileName2, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab1, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab1.myTextBox2.setText(fileName2)

    def openFileNameDialog3(self):
        fileName3, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab1, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab1.myTextBox3.setText(fileName3)

    def openFileNameDialog4(self):
        fileName4, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab1, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab1.myTextBox4.setText(fileName4)

    def openFileNameDialog5(self):
        fileName5, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab1, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab1.myTextBox5.setText(fileName5)

    def openFileNameDialog6(self):
        fileName6, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab1, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab1.myTextBox6.setText(fileName6)

    def openFileNameDialog7(self):
        fileName7, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab2, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab1.myTextBox4.setText(fileName7)

    def openFileNameDialog8(self):
        fileName8, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab2, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab1.myTextBox5.setText(fileName8)

    def openFileNameDialog9(self):
        fileName9, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab2, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab2.myTextBox9.setText(fileName9)

    def openFileNameDialog10(self):
        fileName10, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab2, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab2.myTextBox10.setText(fileName10)

    def openFileNameDialog20(self):
        fileName20, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab1, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab1.myTextBox6.setText(fileName20)

    def openFileNameDialog30(self):
        fileName30, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab2, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab2.myTextBox7.setText(fileName30)

    def openFileNameDialog40(self):
        fileName40, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab2, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab2.myTextBox8.setText(fileName40)

    def initUI(self, tab):

        # Create coverage textbox
        tab.lbl_coverage = QLabel("Coverage Indicator:", tab)
        tab.lbl_coverage.move(5, 450)
        tab.message = ""
        tab.textbox_coverage = QtWidgets.QTextEdit(self.tab1)
        tab.textbox_coverage.setText(tab.message)
        tab.textbox_coverage.move(140, 450)
        tab.textbox_coverage.resize(70, 20)
        tab.textbox_coverage.setReadOnly(True)

        # Create convergence textbox
        tab.lbl_coverage = QLabel("Convergence Indicator:", tab)
        tab.lbl_coverage.move(300, 450)
        tab.message = ""
        tab.textbox_convergence = QtWidgets.QTextEdit(self.tab1)
        tab.textbox_convergence.setText(tab.message)
        tab.textbox_convergence.move(460, 450)
        tab.textbox_convergence.resize(70, 20)
        tab.textbox_convergence.setReadOnly(True)

        # Create a textbox
        tab.message = ""
        tab.textbox = QtWidgets.QTextEdit(self.tab1)
        tab.textbox.setText(tab.message)
        tab.textbox.move(10, 290)
        tab.textbox.resize(700, 130)
        tab.textbox.setReadOnly(True)



        # create a progress bar
        tab.pbar = QProgressBar(self.tab1)
        tab.pbar.setGeometry(10, 310, 700, 20)
        tab.pbar.setAlignment(QtCore.Qt.AlignCenter)
        tab.pbar.setValue(0)
        tab.pbar.move(10, 420)

        # Create a color textbox1
        tab.colorTextBox1 = QtWidgets.QTextEdit(self.tab1)
        tab.colorTextBox1.setStyleSheet(" background-color: grey ")
        tab.colorTextBox1.resize(20, 20)
        tab.colorTextBox1.move(710, 10)

        # Create a color textbox2
        tab.colorTextBox2 = QtWidgets.QTextEdit(self.tab1)
        tab.colorTextBox2.setStyleSheet(" background-color: grey ")
        tab.colorTextBox2.resize(20, 20)
        tab.colorTextBox2.move(710, 40)

        # Create a color textbox3
        tab.colorTextBox3 = QtWidgets.QTextEdit(self.tab1)
        tab.colorTextBox3.setStyleSheet(" background-color: grey ")
        tab.colorTextBox3.resize(20, 20)
        tab.colorTextBox3.move(710, 70)

        # Create a color textbox4
        tab.colorTextBox4 = QtWidgets.QTextEdit(self.tab1)
        tab.colorTextBox4.setStyleSheet(" background-color: grey ")
        tab.colorTextBox4.resize(20, 20)
        tab.colorTextBox4.move(710, 100)

        # Create a color textbox5
        tab.colorTextBox5 = QtWidgets.QTextEdit(self.tab1)
        tab.colorTextBox5.setStyleSheet(" background-color: grey ")
        tab.colorTextBox5.resize(20, 20)
        tab.colorTextBox5.move(710, 130)

        # Create a color textbox6
        tab.colorTextBox6 = QtWidgets.QTextEdit(self.tab1)
        tab.colorTextBox6.setStyleSheet(" background-color: grey ")
        tab.colorTextBox6.resize(20, 20)
        tab.colorTextBox6.move(710, 160)

        # Create a drop down list
        tab.lbl = QLabel("Check level", tab)

        tab.combo = QComboBox(tab)
        tab.combo.addItem("Previsional")
        tab.combo.addItem("Consolidated")
        tab.combo.addItem("Validated")
        tab.combo.resize(508, 20.4)  # rezise the drop down list
        tab.combo.move(200, 200)
        tab.lbl.move(5, 205)
        tab.combo.activated[str].connect(self.onActivated)

        # Create a drop down list
        tab.lbl1 = QLabel("Project name", tab)

        tab.combo1 = QComboBox(tab)
        tab.combo1.addItem("   Generic   ")
        tab.combo1.addItem("   All   ")
        tab.combo1.resize(330, 20.4)  # rezise the drop down list
        tab.combo1.move(200, 260)
        tab.lbl1.move(5, 265)
        tab.combo1.activated[str].connect(self.onActivated)

        # Create a dropdown list
        tab.lbl2 = QLabel("Architecture type", tab)

        tab.combo2 = QComboBox(tab)
        tab.combo2.addItem("Archi 2010")
        tab.combo2.addItem("Archi NEA R1")
        tab.combo2.addItem("Archi NEA R2")
        tab.combo2.resize(508, 20.4)
        tab.combo2.move(200, 230)
        tab.lbl2.move(5, 235)
        tab.combo2.activated[str].connect(self.onActivated)

        self.setGeometry(self.left, self.top, self.width, self.height)
        self.setWindowTitle('TSD Checker')

        tab.importNames = QPushButton(tab)
        tab.importNames.setText("Import Project Names")
        tab.importNames.resize(160, 20.4)
        tab.importNames.move(550, 260)

        # File Selectiom Dialog1
        tab.lbl2 = QLabel("TSD File:", tab)
        tab.lbl2.move(5, 15)
        tab.myTextBox1 = QtWidgets.QTextEdit(tab)
        tab.myTextBox1.resize(460, 25)
        tab.myTextBox1.move(200, 10)
        tab.myTextBox1.setReadOnly(True)
        tab.myTextBox1.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        tab.button1 = QPushButton('...', tab)
        tab.button1.clicked.connect(self.openFileNameDialog1)
        tab.button1.move(660, 10)
        tab.button1.resize(45, 22)

        # File Selectiom Dialog2
        tab.lbl3 = QLabel("TSD vehicle Function file:", tab)
        tab.lbl3.move(5, 45)
        tab.myTextBox2 = QtWidgets.QTextEdit(tab)
        tab.myTextBox2.resize(460, 25)
        tab.myTextBox2.move(200, 40)
        tab.myTextBox2.setReadOnly(True)
        tab.myTextBox2.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        tab.button2 = QPushButton('...', tab)
        tab.button2.clicked.connect(self.openFileNameDialog2)
        tab.button2.move(660, 40)
        tab.button2.resize(45, 22)

        # File Selectiom Dialog3
        tab.lbl4 = QLabel("TSD system file:", tab)
        tab.lbl4.move(5, 75)
        tab.myTextBox3 = QtWidgets.QTextEdit(tab)
        tab.myTextBox3.resize(460, 25)
        tab.myTextBox3.move(200, 70)
        tab.myTextBox3.setReadOnly(True)
        tab.myTextBox3.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        tab.button3 = QPushButton('...', tab)
        tab.button3.clicked.connect(self.openFileNameDialog3)
        tab.button3.move(660, 70)
        tab.button3.resize(45, 22)

        # File Selectiom Dialog4
        tab.lbl8 = QLabel("AMDEC:", tab)
        tab.lbl8.move(5, 105)
        tab.myTextBox4 = QtWidgets.QTextEdit(tab)
        tab.myTextBox4.resize(460, 25)
        tab.myTextBox4.move(200, 100)
        tab.myTextBox4.setReadOnly(True)
        tab.myTextBox4.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        tab.button4 = QPushButton('...', tab)
        tab.button4.clicked.connect(self.openFileNameDialog7)
        tab.button4.move(660, 100)
        tab.button4.resize(45, 22)

        # File Selectiom Dialog5
        tab.lbl9 = QLabel("export MedialecMatrice:", tab)
        tab.lbl9.move(5, 135)
        tab.myTextBox5 = QtWidgets.QTextEdit(tab)
        tab.myTextBox5.resize(460, 25)
        tab.myTextBox5.move(200, 130)
        tab.myTextBox5.setReadOnly(True)
        tab.button5 = QPushButton('...', tab)
        tab.button5.clicked.connect(self.openFileNameDialog8)
        tab.button5.move(660, 130)
        tab.button5.resize(45, 22)
        tab.myTextBox5.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)

        # File Selectiom Dialog6
        tab.lbl11 = QLabel("Diagnostic matrix file:", tab)
        tab.lbl11.move(5, 165)
        tab.myTextBox6 = QtWidgets.QTextEdit(tab)
        tab.myTextBox6.resize(460, 25)
        tab.myTextBox6.move(200, 160)
        tab.myTextBox6.setReadOnly(True)
        tab.myTextBox6.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        tab.button6 = QPushButton('...', tab)
        tab.button6.clicked.connect(self.openFileNameDialog20)
        tab.button6.move(660, 160)
        tab.button6.resize(45, 22)

        # Check button
        tab.button = QPushButton('Check', tab)
        tab.button.move(310, 470)
        tab.button.resize(90, 25)
        tab.button.clicked.connect(self.buttonClicked)
        #button.setStyleSheet('QPushButton {background-color: white; color: black;}')
        tab.buttonNew = QPushButton("Open \nReport", tab)
        tab.buttonNew.resize(90, 60)
        tab.buttonNew.move(710, 310)
        tab.buttonNew.setEnabled(False)
        tab.buttonNew.clicked.connect(self.ButtonReportClick)

        self.show()

    def ButtonReportClick(self):

        self.excel = win32.gencache.EnsureDispatch('Excel.Application')

        if self.tab1.myTextBox1.toPlainText():
           fileName = self.tab1.myTextBox1.toPlainText()
           self.excel.Visible = True
           self.excel.Workbooks.Open(fileName)

        if self.tab1.myTextBox2.toPlainText():
           fileName = self.tab1.myTextBox2.toPlainText()
           self.excel.Visible = True
           self.excel.Workbooks.Open(fileName)

        if self.tab1.myTextBox3.toPlainText():
           fileName = self.tab1.myTextBox3.toPlainText()
           self.excel.Visible = True
           self.excel.Workbooks.Open(fileName)

    def initUIOptions(self, tab):

        tab.lblUser = QLabel("USER:", tab)
        tab.lblUser.move(165,25)
        tab.TextBoxUser = QtWidgets.QLineEdit(tab)
        tab.TextBoxUser.resize(200,25)
        tab.TextBoxUser.move(220, 20)
        tab.TextBoxUser.setText("E518720")


        tab.lblPass = QLabel("PASSWORD:", tab)
        tab.lblPass.move(450,25)
        tab.TextBoxPass = QtWidgets.QLineEdit(tab)
        tab.TextBoxPass.resize(180,25)
        tab.TextBoxPass.move(550, 20)
        tab.TextBoxPass.setEchoMode((QLineEdit.Password))
        tab.TextBoxPass.setText("cST12122")


        # File Selectiom Dialog5
        tab.lbl6 = QLabel("Family list export(CESARE):", tab)
        tab.lbl6.move(5, 145)
        tab.myTextBox7 = QtWidgets.QTextEdit(tab)
        tab.myTextBox7.resize(460, 25)
        tab.myTextBox7.move(210, 140)
        tab.myTextBox7.setReadOnly(True)

        tab.link2 = QLabel('''<a href=''' + self.DOC8Link + '''>DocInfo Reference: 02043_18_05471</a>''', tab)
        tab.link2.setOpenExternalLinks(True)
        tab.link2.move(220, 167)


        tab.button7 = QPushButton('...', tab)
        tab.button7.move(670, 141)
        tab.button7.resize(45, 22)
        tab.button7.clicked.connect(self.openFileNameDialog30)



        # File Selectiom Dialog4
        tab.lbl5 = QLabel("Criticity configuration file:", tab)
        tab.lbl5.move(5,215)
        tab.myTextBox8 = QtWidgets.QTextEdit(tab)
        tab.myTextBox8.resize(460, 25)
        tab.myTextBox8.move(210, 210)
        tab.myTextBox8.setReadOnly(True)

        tab.link1 = QLabel('''<a href='''+self.DOC9Link+'''>DocInfo Reference: 02043_18_05474</a>''', tab)
        tab.link1.setOpenExternalLinks(True)
        tab.link1.move(220, 237)

        tab.button8 = QPushButton('...', tab)
        tab.button8.clicked.connect(self.openFileNameDialog40)
        tab.button8.move(670, 211)
        tab.button8.resize(45, 22)



        # File Selectiom Dialog6
        tab.lbl7 = QLabel("Customer effect file:", tab)
        tab.lbl7.move(5, 275)
        tab.myTextBox9 = QtWidgets.QTextEdit(tab)
        tab.myTextBox9.resize(460, 25)
        tab.myTextBox9.move(210, 270)
        tab.myTextBox9.setReadOnly(True)

        tab.link3 = QLabel('''<a href=''' + self.DOC7Link + '''>DocInfo Reference: 02043_18_05499</a>''', tab)
        tab.link3.setOpenExternalLinks(True)
        tab.link3.move(220, 297)

        tab.button9 = QPushButton('...', tab)
        tab.button9.clicked.connect(self.openFileNameDialog9)
        tab.button9.move(670, 271)
        tab.button9.resize(45, 22)

        # File Selectiom Dialog9
        tab.lbl10 = QLabel("Diversity management file:", tab)
        tab.lbl10.move(5, 335)
        tab.myTextBox10 = QtWidgets.QTextEdit(tab)
        tab.myTextBox10.resize(460, 25)
        tab.myTextBox10.move(210,330)
        tab.myTextBox10.setReadOnly(True)

        tab.link4 = QLabel('''<a href=''' + self.DOC13Link + '''>DocInfo Reference: 02016_11_04964</a>''', tab)
        tab.link4.setOpenExternalLinks(True)
        tab.link4.move(220, 357)


        tab.button10 = QPushButton('...', tab)
        tab.button10.clicked.connect(self.openFileNameDialog10)
        tab.button10.move(670, 331)
        tab.button10.resize(45, 22)

        tab.labelInternetAndIntranet = QLabel("Network Type:", tab)
        tab.labelInternetAndIntranet.move(130, 60)
        tab.RadioButtonInternet = QRadioButton(self.tab2)
        tab.RadioButtonInternet.setText("Internet link")
        tab.RadioButtonInternet.setChecked(True)
        tab.RadioButtonIntranet = QRadioButton(self.tab2)
        tab.RadioButtonIntranet.setText("Intranet link")
        tab.RadioButtonInternet.toggled.connect(self.ToggleLink)
        tab.RadioButtonIntranet.toggled.connect(self.ToggleLink)
        tab.RadioButtonInternet.move(230, 58)
        tab.RadioButtonIntranet.move(230, 90)

    def download_file(self, url):
        user = self.tab2.TextBoxUser.text()
        user = str(user)
        password = self.tab2.TextBoxPass.text()
        password = str(password)
        if not user or not password:
            self.tab1.textbox.setText("Missing Username or Password")
            return "False"
        try:
            os.stat(self.fileFolder)
        except:
            os.mkdir(self.fileFolder)
        try:
            response = requests.get(url, stream=True, auth=(user, password))
        except:
            return "Error"
        status = response.status_code
        if status == 401:
            self.tab1.textbox.setText("Username or Password Incorrect")
            return "False"

        FileName = response.headers['Content-Disposition'].split('"')[1]
        FilePath = self.fileFolder + FileName
        success_download = self.tab1.textbox.toPlainText()
        success_download = success_download + "\nfile " + FileName + " has been successfully downloaded\n=======================\n"
        self.tab1.textbox.setText(success_download)
        with open(FilePath, 'wb') as f:
            for chunk in response.iter_content(chunk_size=128):
                f.write(chunk)
        return FilePath

    def onActivated(self):
        return

    def buttonClicked(self):
        return


class Test(Application):

    def __init__(self):
        super().__init__()

        #Tested Files COM Objects
        self.DOC3Workbook = None
        self.DOC4Workbook = None
        self.DOC5Workbook = None
        self.AMDECWorkbook = str()
        self.MedialecWorkbook = str()
        self.DiagnosticWorkbook = str()

        #Tested Files Paths
        self.DOC3Path = str()
        self.DOC4Path = str()
        self.DOC5Path = str()
        self.AMDECPath = str()
        self.MedialecPath = str()
        self.DiagnosticPath = str()

        #Tested Files Paths
        self.DOC3Name = str()
        self.DOC4Name = str()
        self.DOC5Name = str()
        self.AMDECName = str()
        self.MedialecName = str()
        self.DiagnosticName = str()


        # Optional Files Paths
        self.DOC8Path = str() # CESARE
        self.DOC9Path = str() # TSD Config
        self.DOC7Path = str() # Customer effect
        self.DOC13Path = str() # Diversity mng

        # Optional Files Names
        self.DOC8Name = str() # CESARE
        self.DOC9Name = str() # TSD Config
        self.DOC7Name = str() # Customer effect
        self.DOC13Name = str() # Diversity mng
        self.DOC14Name = str()

        # Optional Files Content
        self.DOC9Dict = dict()
        self.DOC13List = []

        # COM Object
        self.excelApp = None

        #Tests Parameters
        self.checkLevel = str()
        self.WorkbookStats = GeneralStructureTester.WorkbookProperties()




        try:
            os.stat(self.fileFolder)
            for file in os.listdir(self.fileFolder):
                try:
                    os.remove(self.fileFolder + file)
                except:
                    os.system("taskkill /f /im EXCEL.EXE")
                    for file in os.listdir(self.fileFolder):
                        os.remove(self.fileFolder + file)
                    break
        except:
            os.mkdir(self.fileFolder)


    def IncrementProgressBar(self):
        self.pBarValue += pBarIncrement
        self.tab1.pbar.setValue(self.pBarValue)


    def buttonClicked(self):


        self.tab1.textbox.setText("File analyzation starting...")
        self.start_time = time.time()
        os.system("taskkill /f /im EXCEL.EXE")
        self.checkLevel = str(self.tab1.combo.currentText()).strip().casefold()
        if self.excelApp is None:
            self.excelApp = win32.gencache.EnsureDispatch('Excel.Application')
        self.excelApp.Visible = False

        self.tab1.colorTextBox1.setStyleSheet(" background-color: grey ")
        self.tab1.colorTextBox2.setStyleSheet(" background-color: grey ")
        self.tab1.colorTextBox3.setStyleSheet(" background-color: grey ")

        self.tab1.textbox.setText("")
        self.tab1.pbar.setValue(0)
        if self.tab1.myTextBox6.toPlainText():
            self.Doc15Path = self.tab1.myTextBox6.toPlainText()

        else:
            self.Doc15Path = None
        if not self.tab2.myTextBox7.toPlainText():
            self.DOC8Path = self.download_file(self.DOC8Link)
        else:
            self.Doc8Path = self.tab2.myTextBox7.toPlainText()
        if self.DOC8Path == "Error":
            self.tab1.textbox.setText(
                "ERROR: No network available\nTo continue, please select files for field in the Options tab ")
            return
        if self.DOC8Path == "False":
            return
        if not self.tab2.myTextBox8.toPlainText():
            self.DOC9Path = self.download_file(self.DOC9Link)
        else:
            self.DOC9Path = self.tab2.myTextBox8.toPlainText()
        if not self.tab2.myTextBox9.toPlainText():
            self.DOC7Path = self.download_file(self.DOC7Link)
        else:
            self.DOC7Path = self.tab2.myTextBox9.toPlainText()
        if not self.tab2.myTextBox10.toPlainText():
            self.DOC13Path = self.download_file(self.DOC13Link)
        else:
            self.DOC13Path = self.tab2.myTextBox10.toPlainText()
        self.DOC9Dict = OptionalFilesParser.DOC9Parser(self, self.excelApp, self.DOC9Path)
        if self.DOC9Dict == None:
            return
        self.DOC13List = OptionalFilesParser.DOC13Parser(self, self.excelApp, self.DOC13Path)
        if self.DOC13List == None:
            return
        self.DOC8List = OptionalFilesParser.DOC8Parser(self, self.excelApp, self.DOC8Path)
        if self.DOC8List == None:
            return
        if self.Doc15Path is not None:
            self.subfamily_name, self.Doc15List = OptionalFilesParser.DOC15Parser(self ,self.Doc15Path)
            if self.subfamily_name == None or self.Doc15List == None:
                return
        else:
            self.Doc15List = None
            self.subfamily_name = None

        #self.DOC8Name = self.download_file(self.DOC8Link)

        self.DOC14Name = self.download_file(self.DOC14Link)
        self.DOC7Name = self.download_file(self.DOC7Link)
        archi_type = self.tab1.combo2.currentText()


        if self.tab1.myTextBox1.toPlainText():
            self.DOC3Name = self.download_file(self.DOC3Link)

            self.DOC3Path = self.tab1.myTextBox1.toPlainText()
            try:
                self.DOC3Workbook = self.excelApp.Workbooks.Open(self.DOC3Path)
            except:
                self.tab1.textbox.setText("ERROR: when trying to parse the plan type Tableau de synthèse diagnosticabilité file " + self.DOC3Path.split('/')[-1])
                return
            if self.DOC3Workbook == None:
                return
            ExcelEdit.AddTestReportSheets(self.DOC3Workbook)
            ExcelEdit.AddTestReportSheetHeader(self.DOC3Workbook)
            check = False
            check_indicator = False
            ok_indicator = 0
            ok = 0

            FileMeasure.DOC3Info1(self.DOC3Workbook, self)

            self.opening_time = time.time()

        #GeneralStructure

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0000"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0000(self.DOC3Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0005"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0005(self.DOC3Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0010"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0010(self.DOC3Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0011"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0011(self.DOC3Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0020"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0020(self.DOC3Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0025"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0025(self.DOC3Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0030"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0030(self.DOC3Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0035"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0035(self.DOC3Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0040"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0040(self.DOC3Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0051"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0051(self.DOC3Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0052"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0052(self.DOC3Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0053"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0053(self.DOC3Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0054"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0054(self.DOC3Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0055"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0055(self.DOC3Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0056"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0056(self.DOC3Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0057"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0057(self.DOC3Workbook, self)


            if archi_type == "Archi 2010":
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0058"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0058(self.DOC3Workbook, self)
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0061"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0061(self.DOC3Workbook, self)
            elif archi_type == "Archi NEA R1":
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0059"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0059(self.DOC3Workbook, self)
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0060"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0060(self.DOC3Workbook, self)
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0062"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0062(self.DOC3Workbook, self)
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0063"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0063(self.DOC3Workbook, self)
            elif archi_type == "Archi NEA R2":
                pass

            # DOC3

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0100"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0100(self.DOC3Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0110"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0110(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0120"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0120(self.DOC3Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0130"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0130(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0140"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0140(self.DOC3Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0150"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0150(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0160"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0160(self.DOC3Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0170"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0170(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0180"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0180(self.DOC3Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0190"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0190(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0200"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0200(self.DOC3Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0210"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0210(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0220"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0220(self.DOC3Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0230"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0230(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0240"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0240(self.DOC3Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0250"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0250(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0260"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0260(self.DOC3Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0270"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0270(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)

        # Wholeness
            if ok == 0 or ok == 1:
                FileMeasure.DOC3Info2(self.DOC3Workbook, self)

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1000"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1000(self.DOC3Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1001"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1001(self.DOC3Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1010"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1010(self.DOC3Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1011"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1011(self.DOC3Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1020"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1020(self.DOC3Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1021"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1021(self.DOC3Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1030"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1030(self.DOC3Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1031"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1031(self.DOC3Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                '''check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1040(self.DOC3Workbook, self)
                if check_indicator == True:
                    ok_indicator = 1'''

                '''check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1041(self.DOC3Workbook, self)
                if check_indicator == True:
                    ok_indicator = 1'''

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1050"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1050(self.DOC3Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1055"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1055(self.DOC3Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1060"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1060(self.DOC3Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1061"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1061(self.DOC3Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1062"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1062(self.DOC3Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1070"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1070(self.DOC3Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1080"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1080(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1090"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1090(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1100"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1100(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1110"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1110(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1120"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1120(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1130"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1130(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1140"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1140(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1150"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1150(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1160"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1160(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1170"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1170(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1180"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1180(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1190"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1190(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1200"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1200(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1210"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1210(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1220"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1220(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1230"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1230(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1240"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1240(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1600"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1600(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1600"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1601(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1602"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1602(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1603"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1603(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1604"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1604(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1605"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1605(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1606"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1606(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1607"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1607(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1608"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1608(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1609"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1609(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1610"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1610(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1611"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1611(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1612"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1612(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1613"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1613(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1615"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1615(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1616"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1616(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1617"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1617(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1618"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1618(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1619"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1619(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1620"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1620(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1621"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1621(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1622"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1622(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1623"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1623(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1624"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1624(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1625"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1625(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1626"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1626(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1627"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1627(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1628"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1628(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1629"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1629(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1630"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1630(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1631"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1631(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1632"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1632(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1650"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1650(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1651"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1651(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1652"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1652(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1653"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1653(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1654"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1654(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1655"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1655(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1656"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1656(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1657"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1657(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1658"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1658(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1659"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1659(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1660"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1660(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1661"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1661(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1662"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1662(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1663"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1663(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1664"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1664(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1684"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1684(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1685"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1685(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1686"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1686(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1687"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1687(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1688"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1688(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1689"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1689(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1690"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1690(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1691"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1691(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1692"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1692(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1693"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1693(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1700"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1700(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1701"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1701(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1702"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1702(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1703"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1703(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1704"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1704(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1705"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1705(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1706"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1706(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1707"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1707(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1708"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1708(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1709"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1709(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1710"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1710(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1711"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1711(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1712"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1712(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1713"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1713(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1714"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1714(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1715"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1715(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1716"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1716(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1717"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1717(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1718"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1718(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1719"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1719(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1750"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1750(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1751"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1751(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1752"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1752(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1753"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1753(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1754"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1754(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1755"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1755(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1756"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1756(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1757"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1757(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1758"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1758(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1759"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1759(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1800"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1800(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1801"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1801(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1802"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1802(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1803"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1803(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1810"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1810(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1811"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1811(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1812"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1812(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1813"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1813(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1814"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1814(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1815"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1815(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1820"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1820(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1821"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1821(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1822"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1822(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1823"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1823(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1824"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1824(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1825"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1825(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1830"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1830(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1831"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1831(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1840"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1840(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1841"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1841(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                #
                # #  Coherence checks
                # if self.DOC9Dict["Test_02043_18_04939_COH_2000"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2000(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_COH_2001"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2001(self.DOC3Workbook, self, self.DOC8List)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_COH_2002"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2002(self.DOC3Workbook, self, self.DOC8List)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_COH_2005"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2005(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_COH_2006"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2006(self.DOC3Workbook, self, self.DOC8List)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_COH_2007"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2007(self.excelApp, self.DOC3Workbook, self, self.DOC14Name)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_COH_2010"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2010(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_COH_2020"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2020(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_COH_2030"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2030(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_COH_2040"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2040(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_COH_2050"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2050(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_COH_2060"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2060(self.excelApp, self.DOC3Workbook, self, self.DOC7Path)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # #check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2091(self.DOC3Workbook, self)
                #
                # if self.DOC9Dict["Test_02043_18_04939_COH_2100"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2100(self.DOC3Workbook, self, self.DOC8List)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_COH_2110"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2110(self.DOC3Workbook, self, self.DOC8List)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_COH_2140"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2140(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_COH_2150"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2150(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_COH_2160"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2160(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_COH_2190"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2190(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_COH_2210"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2210(self.DOC3Workbook, self)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_COH_2230"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2230(self.DOC3Workbook, self, self.subfamily_name, self.Doc15List)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_COH_2240"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2240(self.DOC3Workbook, self, self.DOC13List)
                #     if check_indicator == True:
                #         ok_indicator = 1
                #
                # if self.DOC9Dict["Test_02043_18_04939_COH_2250"][self.checkLevel].casefold().strip() != "n/a":
                #     check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2250(self.DOC3Workbook, self,self.DOC13List)
                #     if check_indicator == True:
                #         ok_indicator = 1


                self.coverage = IndicatorTester.coverageIndicator(self.DOC3Workbook, self) * 100
                self.tab1.textbox_coverage.setText(str(self.coverage)[0:4] + "%")
                self.IncrementProgressBar()

                self.convergence = IndicatorTester.convergenceIndicator(self.DOC3Workbook, self) * 100
                self.tab1.textbox_convergence.setText(str(self.convergence)[0:4] + "%")
                self.IncrementProgressBar()

                if ok_indicator == 1:
                    self.tab1.colorTextBox1.setStyleSheet("background-color: red")
                    self.status = "Failed"
                    self.tab1.buttonNew.setEnabled(True)
                else:
                    self.tab1.colorTextBox1.setStyleSheet("background-color: green")
                    self.status = "Passed"
                    self.tab1.buttonNew.setEnabled(True)

                self.end_time = time.time()

                ExcelEdit.WriteReportInformationSheet(self.DOC3Workbook, self)

                self.DOC3Workbook.Save()

                ExcelEdit.ExcelWrite(self.return_list, self.DOC3Path)


                # self.DOC3Workbook.Save()


            elif ok == 1:
                self.tab1.colorTextBox1.setStyleSheet("background-color: red")
                self.status = "Fail"
                self.tab1.buttonNew.setEnabled(True)
                self.tab1.pbar.setValue(100)
                self.end_time = time.time()
                ExcelEdit.WriteReportInformationSheet(self.DOC3Workbook, self)
                self.DOC3Workbook.Save()



        if self.tab1.myTextBox2.toPlainText():
            self.DOC4Name = self.download_file(self.DOC4Link)

            self.DOC4Path = self.tab1.myTextBox2.toPlainText()
            try:
                self.DOC4Workbook = self.excelApp.Workbooks.Open(self.DOC4Path)
            except:
                self.tab1.textbox.setText("ERROR: when trying to parse the plan type TSD Fonction véhicule file " + self.DOC4Path.split('/')[-1])
                return
            if self.DOC4Workbook == None:
                return
            ExcelEdit.AddTestReportSheets(self.DOC4Workbook)
            ExcelEdit.AddTestReportSheetHeader(self.DOC4Workbook)
            check = False
            check_indicator = False
            ok_indicator = 0
            ok = 0

            self.DOC4Workbook.Save()

            FileMeasure.DOC4Info1(self.DOC4Workbook, self)

            # GeneralStructure

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0000"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0000(self.DOC4Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0005"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0005(self.DOC4Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0010"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0010(self.DOC4Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0011"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0011(self.DOC4Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0020"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0020(self.DOC4Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0025"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0025(self.DOC4Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0030"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0030(self.DOC4Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0035"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0035(self.DOC4Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0040"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0040(self.DOC4Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0051"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0051(self.DOC4Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0052"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0052(self.DOC4Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0053"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0053(self.DOC4Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0054"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0054(self.DOC4Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0055"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0055(self.DOC4Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0056"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0056(self.DOC4Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0057"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0057(self.DOC4Workbook, self)

            if archi_type == "Archi 2010":
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0058"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0058(self.DOC4Workbook, self)
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0061"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0061(self.DOC4Workbook, self)
            elif archi_type == "Archi NEA R1":
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0059"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0059(self.DOC4Workbook, self)
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0060"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0060(self.DOC4Workbook, self)
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0062"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0062(self.DOC4Workbook, self)
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0063"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0063(self.DOC4Workbook, self)
            elif archi_type == "Archi NEA R2":
                pass

        # DOC4
            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0400"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0400(self.DOC4Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0410"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0410(self.excelApp, self.DOC4Workbook, self, self.DOC4Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0420"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0420(self.DOC4Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0430"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0430(self.excelApp, self.DOC4Workbook, self, self.DOC4Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0440"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0440(self.DOC4Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0450"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0450(self.excelApp, self.DOC4Workbook, self, self.DOC4Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0460"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0460(self.DOC4Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0470"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0470(self.excelApp, self.DOC4Workbook, self, self.DOC4Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0480"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0480(self.DOC4Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0490"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0490(self.excelApp, self.DOC4Workbook, self, self.DOC4Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0500"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0500(self.DOC4Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0510"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0510(self.excelApp, self.DOC4Workbook, self, self.DOC4Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0520"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0520(self.DOC4Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0530"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0530(self.excelApp, self.DOC4Workbook, self, self.DOC4Name)

            # Wholeness

            if ok == 1 or ok == 0:
                FileMeasure.DOC4Info2(self.DOC4Workbook, self)

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1000"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1000(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1001"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1001(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1010"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1010(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1011"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1011(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1020"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1020(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1021"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1021(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1030"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1030(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1031"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1031(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1040"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1040(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1041"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1041(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1050"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1050(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1055"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1055(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1060"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1060(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1061"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1061(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1062"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1062(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1070"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1070(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1080"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1080(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1090"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1090(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1100"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1100(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1110"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1110(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1120"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1120(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1130"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1130(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1140"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1140(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1150"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1150(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1160"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1160(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1170"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1170(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1180"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1180(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1190"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1190(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1200"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1200(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1210"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1210(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1220"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1220(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1230"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1230(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1240"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1240(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1300"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1300(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1301"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1301(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1302"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1302(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1303"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1303(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1304"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1304(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1305"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1305(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1306"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1306(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1307"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1307(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1308"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1308(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1309"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1309(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1310"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1310(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1311"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1311(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1312"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1312(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1313"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1313(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1314"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1314(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1315"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1315(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1316"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1316(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1317"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1317(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1318"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1318(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1319"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1319(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1320"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1320(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1321"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1321(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1322"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1322(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1323"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1323(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1324"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1324(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1325"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1325(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1326"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1326(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1327"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1327(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1328"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1328(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1329"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1329(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1330"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1330(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1331"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1331(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1332"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1332(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1333"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1333(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1334"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1334(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1350"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1350(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1351"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1351(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1352"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1352(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1353"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1353(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1354"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1354(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1355"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1355(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1356"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1356(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1357"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1357(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1358"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1358(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1359"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1359(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1360"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1360(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1361"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1361(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1400"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1400(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1401"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1401(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1402"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1402(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1403"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1403(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1430"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1430(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1431"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1431(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1432"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1432(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1433"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1433(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1434"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1434(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1435"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1435(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1450"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1450(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1451"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1451(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1452"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1452(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1453"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1453(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1454"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1454(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1455"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1455(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1456"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1456(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1500"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1500(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1501"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1501(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1550"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1550(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1551"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1551(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1552"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1552(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

            # Coherence checks

                if self.DOC9Dict["Test_02043_18_04939_COH_2000"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2000(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2001"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2001(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2005"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2005(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2007"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2007(self.excelApp, self.DOC4Workbook, self, self.DOC14Name)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2010"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2010(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2020"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2020(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2030"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2030(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2040"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2040(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2050"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2050(self.DOC4Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2070"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2070(self.excelApp, self.DOC4Workbook, self, self.DOC7Path)
                    if check_indicator == True:
                        ok_indicator = 1

                #check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2091(self.DOC4Workbook, self)

                if self.DOC9Dict["Test_02043_18_04939_COH_2120"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2120(self.excelApp, self.DOC4Workbook, self, self.DOC5Name)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2241"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2241(self.DOC4Workbook, self, self.DOC13List)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2251"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2251(self.DOC4Workbook, self, self.DOC13List)
                    if check_indicator == True:
                        ok_indicator = 1

                self.coverage = IndicatorTester.coverageIndicator(self.DOC4Workbook, self) * 100
                self.tab1.textbox_coverage.setText(str(self.coverage)[0:4] + "%")

                self.convergence = IndicatorTester.convergenceIndicator(self.DOC4Workbook, self) * 100
                self.tab1.textbox_convergence.setText(str(self.convergence)[0:4] + "%")

                if ok_indicator == 1:
                    self.tab1.colorTextBox2.setStyleSheet("background-color: red")
                    self.status = "Pass"
                    self.tab1.buttonNew.setEnabled(True)
                else:
                    self.tab1.colorTextBox2.setStyleSheet("background-color: green")
                    self.status = "Fail"
                    self.tab1.buttonNew.setEnabled(True)

                ExcelEdit.WriteReportInformationSheet(self.DOC4Workbook, self)
                self.DOC4Workbook.Save()
            elif ok == 1:
                self.tab1.colorTextBox2.setStyleSheet("background-color: red")
                self.status = "Fail"
                self.tab1.buttonNew.setEnabled(True)
                self.tab1.pbar.setValue(100)
                ExcelEdit.WriteReportInformationSheet(self.DOC4Workbook, self)
                self.DOC4Workbook.Save()


        if self.tab1.myTextBox3.toPlainText():
            self.DOC5Name = self.download_file(self.DOC5Link)

            self.DOC5Path = self.tab1.myTextBox3.toPlainText()
            try:
                self.DOC5Workbook = self.excelApp.Workbooks.Open(self.DOC5Path)
            except:
                self.tab1.textbox.setText("ERROR: when trying to parse the plan type TSD Système file " + self.DOC5Path.split('/')[-1])
                return
            if self.DOC5Workbook == None:
                return
            ExcelEdit.AddTestReportSheets(self.DOC5Workbook)
            ExcelEdit.AddTestReportSheetHeader(self.DOC5Workbook)
            check = False
            check_indicator = False
            ok_indicator = 0
            ok = 0

            FileMeasure.DOC5Info1(self.DOC5Workbook, self)

            # GeneralStructure

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0000"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0000(self.DOC5Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0005"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0005(self.DOC5Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0010"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0010(self.DOC5Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0011"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0011(self.DOC5Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0020"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0020(self.DOC5Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0025"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0025(self.DOC5Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0030"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0030(self.DOC5Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0035"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0035(self.DOC5Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0040"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0040(self.DOC5Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0051"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0051(self.DOC5Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0052"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0052(self.DOC5Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0053"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0053(self.DOC5Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0054"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0054(self.DOC5Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0055"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0055(self.DOC5Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0056"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0056(self.DOC5Workbook, self)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0057"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0057(self.DOC5Workbook, self)


            if archi_type == "Archi 2010":
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0058"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0058(self.DOC5Workbook, self)
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0061"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0061(self.DOC5Workbook, self)
            elif archi_type == "Archi NEA R1":
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0059"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0059(self.DOC5Workbook, self)
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0060"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0060(self.DOC5Workbook, self)
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0062"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0062(self.DOC5Workbook, self)
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0063"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0063(self.DOC5Workbook, self)
            elif archi_type == "Archi NEA R2":
                pass

            # DOC5
            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0700"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0700(self.DOC5Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0710"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0710(self.excelApp, self.DOC5Workbook, self, self.DOC5Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0720"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0720(self.DOC5Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0730"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0730(self.excelApp, self.DOC5Workbook, self, self.DOC5Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0740"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0740(self.DOC5Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0750"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0750(self.excelApp, self.DOC5Workbook, self, self.DOC5Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0760"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0760(self.DOC5Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0770"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0770(self.excelApp, self.DOC5Workbook, self, self.DOC5Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0780"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0780(self.DOC5Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0790"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0790(self.excelApp, self.DOC5Workbook, self, self.DOC5Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0800"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0800(self.DOC5Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0810"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0810(self.excelApp, self.DOC5Workbook, self, self.DOC5Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0820"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0820(self.DOC5Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0830"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0830(self.excelApp, self.DOC5Workbook, self, self.DOC5Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0840"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0840(self.DOC5Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0850"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0850(self.excelApp, self.DOC5Workbook, self, self.DOC5Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0860"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0860(self.DOC5Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0870"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0870(self.excelApp, self.DOC5Workbook, self, self.DOC5Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0880"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0880(self.DOC5Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0890"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0890(self.excelApp, self.DOC5Workbook, self, self.DOC5Name)

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0900"][self.checkLevel].casefold().strip() != "n/a":
                check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0900(self.DOC5Workbook, self)
                if check == True:
                    ok = 1

            if self.DOC9Dict["Test_02043_18_04939_STRUCT_0910"][self.checkLevel].casefold().strip() != "n/a":
                GeneralStructureTester.Test_02043_18_04939_STRUCT_0910(self.excelApp, self.DOC5Workbook, self, self.DOC5Name)


            # Wholeness
            if ok == 0 or ok == 1:
                FileMeasure.DOC5Info2(self.DOC5Workbook, self)

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1000"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1000(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1001"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1001(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1010"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1010(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1011"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1011(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1020"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1020(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1000"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1021(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1030"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1030(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1031"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1031(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1040"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1040(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1041"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1041(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1050"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1050(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1055"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1055(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1060"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1060(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1061"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1061(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1062"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1062(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1070"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1070(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1080"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1080(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1090"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1090(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1100"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1100(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1110"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1110(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1120"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1120(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1130"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1130(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1140"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1140(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1150"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1150(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1160"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1160(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1170"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1170(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1180"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1180(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1190"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1190(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1200"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1200(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1210"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1210(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1220"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1220(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1230"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1230(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1230"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1230(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1240"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1240(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1900"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1900(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1901"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1901(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1902"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1902(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1903"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1903(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1904"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1904(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1905"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1905(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1906"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1906(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1907"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1907(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1908"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1908(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1909"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1909(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1910"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1910(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1911"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1911(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1912"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1912(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1913"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1913(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1914"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1914(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1915"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1915(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1916"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1916(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1917"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1917(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1918"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1918(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1919"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1919(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1920"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1920(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1921"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1921(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1922"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1922(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1923"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1923(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1924"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1924(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1925"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1925(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1926"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1926(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1927"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1927(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1950"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1950(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1951"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1951(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1952"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1952(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1953"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1953(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1954"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1954(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1955"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1955(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1956"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1956(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1957"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1957(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1958"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1958(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1959"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1959(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1960"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1960(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1961"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1961(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1962"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1962(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1963"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1963(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1964"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1964(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1965"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1965(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1966"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1966(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1967"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1967(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1968"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1968(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1969"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1969(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2000"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2000(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2001"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2001(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2002"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2002(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2003"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2003(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2004"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2004(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2005"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2005(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2006"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2006(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2007"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2007(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2008"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2008(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2009"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2009(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2010"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2010(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2011"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2011(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2050"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2050(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2051"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2051(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2052"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2052(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2053"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2053(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2054"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2054(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2055"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2055(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2056"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2056(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2060"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2060(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2061"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2061(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2062"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2062(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2070"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2070(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2071"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2071(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2072"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2072(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2080"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2080(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2081"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2081(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2082"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2082(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2083"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2083(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2084"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2084(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2090"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2090(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2091"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2091(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2092"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2092(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2100"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2100(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2101"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2101(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2102"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2102(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2110"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2110(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2111"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2111(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2112"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2112(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2120"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2120(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2121"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2121(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                # Coherence checks
                if self.DOC9Dict["Test_02043_18_04939_COH_2000"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2000(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2001"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2001(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2002"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2002(self.DOC5Workbook, self, self.DOC8List)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2005"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2005(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2006"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2006(self.DOC5Workbook, self, self.DOC8Name)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2007"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2007(self.excelApp, self.DOC5Workbook, self, self.DOC14Name)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2010"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2010(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2020"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2020(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2030"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2030(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2040"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2040(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2050"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2050(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2080"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2080(self.excelApp, self.DOC5Workbook, self, self.DOC7Path)
                    if check_indicator == True:
                        ok_indicator = 1

                #check_indicator =  Coherence_checksTester.Test_02043_18_04939_COH_2091(self.DOC5Workbook, self)

                if self.DOC9Dict["Test_02043_18_04939_COH_2130"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2130(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2170"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2170(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2180"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2180(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2200"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2200(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2220"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2220(self.DOC5Workbook, self)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2230"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2230(self.DOC5Workbook, self,self.subfamily_name, self.Doc15List)
                    if check_indicator == True:
                       ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2240"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2240(self.DOC5Workbook, self, self.DOC13List)
                    if check_indicator == True:
                        ok_indicator = 1

                if self.DOC9Dict["Test_02043_18_04939_COH_2250"][self.checkLevel].casefold().strip() != "n/a":
                    check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2250(self.DOC5Workbook, self, self.DOC13List)
                    if check_indicator == True:
                        ok_indicator = 1

                self.coverage = IndicatorTester.coverageIndicator(self.DOC5Workbook, self) * 100
                self.tab1.textbox_coverage.setText(str(self.coverage)[0:4] + "%")

                self.convergence = IndicatorTester.convergenceIndicator(self.DOC5Workbook, self) * 100
                self.tab1.textbox_convergence.setText(str(self.convergence)[0:4] + "%")

                if ok_indicator == 1:
                    self.tab1.colorTextBox3.setStyleSheet("background-color: red")
                    self.status = "Fail"
                    self.tab1.buttonNew.setEnabled(True)
                else:
                    self.tab1.colorTextBox3.setStyleSheet("background-color: green")
                    self.status = "Pass"
                    self.tab1.buttonNew.setEnabled(True)

                ExcelEdit.WriteReportInformationSheet(self.DOC5Workbook, self)
                self.DOC5Workbook.Save()
            elif ok == 1:
                self.tab1.colorTextBox3.setStyleSheet("background-color: red")
                self.status = "Fail"
                self.tab1.buttonNew.setEnabled(True)
                self.tab1.pbar.setValue(100)
                ExcelEdit.WriteReportInformationSheet(self.DOC5Workbook, self)
                self.DOC5Workbook.Save()


        self.excelApp.Quit()
        self.excelApp = None

        i = 5


if __name__ == '__main__':


    try:
        FindWindow(None, appName)
        windll.user32.MessageBoxW(0, "Application already running", "Warning", 0|48)

    except:
        app = QApplication(sys.argv)
        apel = Test()
        myQLabel = QLabel()
        sys.exit(app.exec_())
