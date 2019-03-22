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
from timeit import default_timer as timer
import ExcelEdit
import WholenessTester
import Coherence_checksTester


appName = "TSD Checker V0.5.3"
pBarIncrement = 100/85

class Application(QWidget):

    def __init__(self):
        super().__init__()
        self.left = 200
        self.top = 200
        self.width = 900
        self.height = 550
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
        self.username = os.environ['USERNAME']
        self.fileFolder = "C:/Users/" + self.username + "/AppData/Local/Temp/TSD_Checker/"
        self.pBarValue = 0


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
            self.tab2.link1.setText('''<a href=''' + self.DOC9Link + '''>DocInfo Reference: 02043_18_05472</a>''')
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
            self.tab2.link1.setText('''<a href=''' + self.DOC9Link + '''>DocInfo Reference: 02043_18_05472</a>''')
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
        self.tab2.myTextBox7.setText(fileName7)


    def openFileNameDialog8(self):
        fileName8, _filter = QtWidgets.QFileDialog.getOpenFileName(self.ta21, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab2.myTextBox8.setText(fileName8)


    def openFileNameDialog9(self):
        fileName9, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab2, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab2.myTextBox9.setText(fileName9)


    def openFileNameDialog10(self):
        fileName10, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab2, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab2.myTextBox10.setText(fileName10)


    def initUI(self, tab):

        # Create a textbox
        tab.message = ""
        tab.textbox = QtWidgets.QTextEdit(self.tab1)
        tab.textbox.setText(tab.message)
        tab.textbox.move(10, 270)
        tab.textbox.resize(700, 130)
        tab.textbox.setReadOnly(True)

        sb = tab.textbox.verticalScrollBar()
        sb.setValue(sb.minimum())

        # create a progress bar
        tab.pbar = QProgressBar(self.tab1)
        tab.pbar.setGeometry(10, 310, 700, 20)
        tab.pbar.setAlignment(QtCore.Qt.AlignCenter)
        tab.pbar.setValue(0)
        tab.pbar.move(10, 410)

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
        tab.combo1.resize(378, 20.4)  # rezise the drop down list
        tab.combo1.move(200, 230)
        tab.lbl1.move(5, 235)
        tab.combo1.activated[str].connect(self.onActivated)

        self.setGeometry(self.left, self.top, self.width, self.height)
        self.setWindowTitle('TSD Checker')

        tab.importNames = QPushButton(tab)
        tab.importNames.setText("Import Project Names")
        tab.importNames.resize(120, 20.4)
        tab.importNames.move(585, 230)

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
        tab.button6 = QPushButton('...', tab)
        tab.button6.clicked.connect(self.openFileNameDialog10)
        tab.button6.move(660, 160)
        tab.button6.resize(45, 22)
        tab.myTextBox6.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)

        # Check button
        button = QPushButton('Check', tab)
        button.move(310, 470)
        button.resize(90, 25)
        button.clicked.connect(self.buttonClicked)
        button.setStyleSheet('QPushButton {background-color: white; color: black;}')
        buttonNew = QPushButton("Open \nReport", tab)
        buttonNew.resize(90, 60)
        buttonNew.move(710, 310)
        buttonNew.clicked.connect(self.ButtonReportClick)

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
        tab.TextBoxUser.move(200, 20)
        tab.TextBoxUser.setText("E518720")


        tab.lblPass = QLabel("PASSWORD:", tab)
        tab.lblPass.move(450,25)
        tab.TextBoxPass = QtWidgets.QLineEdit(tab)
        tab.TextBoxPass.resize(180,25)
        tab.TextBoxPass.move(520, 20)
        tab.TextBoxPass.setEchoMode((QLineEdit.Password))
        tab.TextBoxPass.setText("Cst67677")


        # File Selectiom Dialog5
        tab.lbl6 = QLabel("Famille/Sous-Famille list export(CESARE):", tab)
        tab.lbl6.move(5, 145)
        tab.myTextBox7 = QtWidgets.QTextEdit(tab)
        tab.myTextBox7.resize(460, 25)
        tab.myTextBox7.move(200, 140)
        tab.myTextBox7.setReadOnly(True)

        tab.link2 = QLabel('''<a href=''' + self.DOC8Link + '''>DocInfo Reference: 02043_18_05471</a>''', tab)
        tab.link2.setOpenExternalLinks(True)
        tab.link2.move(720, 145)


        tab.button7 = QPushButton('...', tab)
        tab.button7.move(660, 140)
        tab.button7.resize(45, 22)
        tab.button7.clicked.connect(self.openFileNameDialog7)



        # File Selectiom Dialog4
        tab.lbl5 = QLabel("TSD configuration file:", tab)
        tab.lbl5.move(5,185)
        tab.myTextBox8 = QtWidgets.QTextEdit(tab)
        tab.myTextBox8.resize(460, 25)
        tab.myTextBox8.move(200, 180)
        tab.myTextBox8.setReadOnly(True)

        tab.link1 = QLabel('''<a href='''+self.DOC9Link+'''>DocInfo Reference: 02043_18_05472</a>''', tab)
        tab.link1.setOpenExternalLinks(True)
        tab.link1.move(720, 185)

        tab.button8 = QPushButton('...', tab)
        tab.button8.clicked.connect(self.openFileNameDialog8)
        tab.button8.move(660, 180)
        tab.button8.resize(45, 22)



        # File Selectiom Dialog6
        tab.lbl7 = QLabel("Customer effect file:", tab)
        tab.lbl7.move(5, 225)
        tab.myTextBox9 = QtWidgets.QTextEdit(tab)
        tab.myTextBox9.resize(460, 25)
        tab.myTextBox9.move(200, 220)
        tab.myTextBox9.setReadOnly(True)

        tab.link3 = QLabel('''<a href=''' + self.DOC7Link + '''>DocInfo Reference: 02043_18_05499</a>''', tab)
        tab.link3.setOpenExternalLinks(True)
        tab.link3.move(720, 225)



        tab.button9 = QPushButton('...', tab)
        tab.button9.clicked.connect(self.openFileNameDialog9)
        tab.button9.move(660, 220)
        tab.button9.resize(45, 22)

        # File Selectiom Dialog9
        tab.lbl10 = QLabel("Diversity management file:", tab)
        tab.lbl10.move(5, 265)
        tab.myTextBox10 = QtWidgets.QTextEdit(tab)
        tab.myTextBox10.resize(460, 25)
        tab.myTextBox10.move(200,260)
        tab.myTextBox10.setReadOnly(True)

        tab.link4 = QLabel('''<a href=''' + self.DOC13Link + '''>DocInfo Reference: 02016_11_04964</a>''', tab)
        tab.link4.setOpenExternalLinks(True)
        tab.link4.move(720, 265)


        tab.button10 = QPushButton('...', tab)
        tab.button10.clicked.connect(self.openFileNameDialog10)
        tab.button10.move(660, 260)
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
        tab.RadioButtonInternet.move(210, 58)
        tab.RadioButtonIntranet.move(210, 90)


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
            for chuck in response.iter_content(chunk_size=128):
                f.write(chuck)
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
        self.DOC7path = str() # Customer effect
        self.DOC13Path = str() # Diversity mng

        # Optional Files Names
        self.DOC8Name = str() # CESARE
        self.DOC9Name = str() # TSD Config
        self.DOC7Name = str() # Customer effect
        self.DOC13Name = str() # Diversity mng
        self.DOC14Name = str()

        # Optional Files Content
        self.DOC9Dict = dict()

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

        os.system("taskkill /f /im EXCEL.EXE")
        self.checkLevel = str(self.tab1.combo.currentText()).strip().casefold()
        if self.excelApp is None:
            self.excelApp = win32.gencache.EnsureDispatch('Excel.Application')
        self.excelApp.Visible = True


        self.tab1.colorTextBox1.setStyleSheet('background-color: grey')
        self.tab1.colorTextBox2.setStyleSheet('background-color: grey')
        self.tab1.colorTextBox3.setStyleSheet('background-color: grey')

        self.tab1.textbox.setText("")
        self.tab1.pbar.setValue(0)

        if not self.tab2.myTextBox7.toPlainText():
            self.DOC8Path = self.download_file(self.DOC8Link)

        if self.DOC8Path == "Error":
            self.tab1.textbox.setText(
                "ERROR: No network available\nto continue, please select files for field in the Options tab ")
            return
        if self.DOC8Path == "False":
            return

        if not self.tab2.myTextBox8.toPlainText():
            self.DOC9Path = self.download_file(self.DOC9Link)

        if not self.tab2.myTextBox9.toPlainText():
            self.DOC7path = self.download_file(self.DOC7Link)

        if not self.tab2.myTextBox10.toPlainText():
            self.DOC13Path = self.download_file(self.DOC13Link)

        self.DOC9Dict = OptionalFilesParser.DOC9Parser(self.excelApp, self.DOC9Path)

        self.DOC3Name = self.download_file(self.DOC3Link)

        self.DOC4Name = self.download_file(self.DOC4Link)

        self.DOC5Name = self.download_file(self.DOC5Link)

        self.DOC8Name = self.download_file(self.DOC8Link)

        self.DOC14Name = self.download_file(self.DOC14Link)

        self.DOC7Name = self.download_file(self.DOC7Link)



        if self.tab1.myTextBox1.toPlainText():
            self.DOC3Path = self.tab1.myTextBox1.toPlainText()
            self.DOC3Workbook = self.excelApp.Workbooks.Open(self.DOC3Path)
            ExcelEdit.AddTestReportSheets(self.DOC3Workbook)
            ExcelEdit.AddTestReportSheetHeader(self.DOC3Workbook)

        #GeneralStructure

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0000(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0005(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0010(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0011(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0020(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0025(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0030(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0035(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0040(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0051(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0052(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0053(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0054(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0055(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0056(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0057(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0058(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0059(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0060(self.DOC3Workbook, self)

            # DOC3

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0100(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0110(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0120(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0130(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0140(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0150(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0160(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0170(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0180(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0190(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0200(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0210(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0220(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0230(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0240(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0250(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0260(self.DOC3Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0270(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)

        # Wholeness

            WholenessTester.Test_02043_18_04939_WHOLENESS_1000(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1001(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1010(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1011(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1020(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1021(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1030(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1031(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1040(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1041(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1050(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1055(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1060(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1061(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1062(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1070(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1080(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1090(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1100(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1110(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1120(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1130(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1140(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1150(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1160(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1170(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1180(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1190(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1200(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1210(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1220(self.DOC3Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1230(self.DOC3Workbook, self)


            #Coherence checks

            Coherence_checksTester.Test_02043_18_04939_COH_2000(self.DOC3Workbook, self)

            Coherence_checksTester.Test_02043_18_04939_COH_2005(self.DOC3Workbook, self)

            Coherence_checksTester.Test_02043_18_04939_COH_2006(self.excelApp, self.DOC3Workbook, self, self.DOC8Name)

            Coherence_checksTester.Test_02043_18_04939_COH_2007(self.excelApp, self.DOC3Workbook, self, self.DOC14Name)

            Coherence_checksTester.Test_02043_18_04939_COH_2010(self.DOC3Workbook, self)

            Coherence_checksTester.Test_02043_18_04939_COH_2020(self.DOC3Workbook, self)

            Coherence_checksTester.Test_02043_18_04939_COH_2030(self.DOC3Workbook, self)

            Coherence_checksTester.Test_02043_18_04939_COH_2040(self.DOC3Workbook, self)

            Coherence_checksTester.Test_02043_18_04939_COH_2050(self.DOC3Workbook, self)

           # Coherence_checksTester.Test_02043_18_04939_COH_2060(self.excelApp, self.DOC3Workbook, self, self.DOC7Name)

            ExcelEdit.WriteReportInformationSheet(self.DOC3Workbook, self)
            self.DOC3Workbook.Save()


       # del self.WorkbookStats

        if self.tab1.myTextBox2.toPlainText():
            self.DOC4Path = self.tab1.myTextBox2.toPlainText()
            self.DOC4Workbook = self.excelApp.Workbooks.Open(self.DOC4Path)
            ExcelEdit.AddTestReportSheets(self.DOC4Workbook)
            ExcelEdit.AddTestReportSheetHeader(self.DOC4Workbook)

            # GeneralStructure

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0000(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0005(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0010(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0011(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0020(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0025(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0030(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0035(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0040(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0051(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0052(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0053(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0054(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0055(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0056(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0057(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0058(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0059(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0060(self.DOC4Workbook, self)

        # DOC4
            GeneralStructureTester.Test_02043_18_04939_STRUCT_0400(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0410(self.excelApp, self.DOC4Workbook, self, self.DOC4Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0420(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0430(self.excelApp, self.DOC4Workbook, self, self.DOC4Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0440(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0450(self.excelApp, self.DOC4Workbook, self, self.DOC4Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0460(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0470(self.excelApp, self.DOC4Workbook, self, self.DOC4Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0480(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0490(self.excelApp, self.DOC4Workbook, self, self.DOC4Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0500(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0510(self.excelApp, self.DOC4Workbook, self, self.DOC4Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0520(self.DOC4Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0530(self.excelApp, self.DOC4Workbook, self, self.DOC4Name)

            # Wholeness

            WholenessTester.Test_02043_18_04939_WHOLENESS_1000(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1001(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1010(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1011(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1020(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1021(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1030(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1031(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1040(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1041(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1050(self.DOC4Workbook, self)

            #WholenessTester.Test_02043_18_04939_WHOLENESS_1055(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1060(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1061(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1062(self.DOC4Workbook, self)

            #WholenessTester.Test_02043_18_04939_WHOLENESS_1070(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1080(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1090(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1100(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1110(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1120(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1130(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1140(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1150(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1160(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1170(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1180(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1190(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1200(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1210(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1220(self.DOC4Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1230(self.DOC4Workbook, self)

            # Coherence checks

            '''Coherence_checksTester.Test_02043_18_04939_COH_2000(self.DOC4Workbook, self)

            Coherence_checksTester.Test_02043_18_04939_COH_2005(self.DOC4Workbook, self)

            Coherence_checksTester.Test_02043_18_04939_COH_2006(self.excelApp, self.DOC4Workbook, self, self.DOC8Name)

            Coherence_checksTester.Test_02043_18_04939_COH_2007(self.excelApp, self.DOC4Workbook, self, self.DOC14Name)

            Coherence_checksTester.Test_02043_18_04939_COH_2010(self.DOC4Workbook, self)

            Coherence_checksTester.Test_02043_18_04939_COH_2020(self.DOC4Workbook, self)

            Coherence_checksTester.Test_02043_18_04939_COH_2030(self.DOC4Workbook, self)

            Coherence_checksTester.Test_02043_18_04939_COH_2040(self.DOC34Workbook, self)

            Coherence_checksTester.Test_02043_18_04939_COH_2050(self.DOC4Workbook, self)'''


            ExcelEdit.WriteReportInformationSheet(self.DOC4Workbook, self)
            self.DOC4Workbook.Save()


        if self.tab1.myTextBox3.toPlainText():
            self.DOC5Path = self.tab1.myTextBox3.toPlainText()
            self.DOC5Workbook = self.excelApp.Workbooks.Open(self.DOC5Path)
            ExcelEdit.AddTestReportSheets(self.DOC5Workbook)
            ExcelEdit.AddTestReportSheetHeader(self.DOC5Workbook)

            # GeneralStructure

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0000(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0005(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0010(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0011(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0020(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0025(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0030(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0035(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0040(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0051(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0052(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0053(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0054(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0055(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0056(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0057(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0058(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0059(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0060(self.DOC5Workbook, self)

            # DOC5

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0700(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0710(self.excelApp, self.DOC5Workbook, self,
                                                                   self.DOC5Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0720(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0730(self.excelApp, self.DOC5Workbook, self,
                                                                   self.DOC5Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0740(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0750(self.excelApp, self.DOC5Workbook, self,
                                                                   self.DOC5Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0760(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0770(self.excelApp, self.DOC5Workbook, self,
                                                                   self.DOC5Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0780(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0790(self.excelApp, self.DOC5Workbook, self,
                                                                   self.DOC5Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0800(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0810(self.excelApp, self.DOC5Workbook, self,
                                                                   self.DOC5Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0820(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0830(self.excelApp, self.DOC5Workbook, self,
                                                                   self.DOC5Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0840(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0850(self.excelApp, self.DOC5Workbook, self,
                                                                   self.DOC5Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0860(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0870(self.excelApp, self.DOC5Workbook, self,
                                                                   self.DOC5Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0880(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0890(self.excelApp, self.DOC5Workbook, self,
                                                                   self.DOC5Name)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0900(self.DOC5Workbook, self)

            GeneralStructureTester.Test_02043_18_04939_STRUCT_0910(self.excelApp, self.DOC5Workbook, self,
                                                                   self.DOC5Name)

            # Wholeness

            WholenessTester.Test_02043_18_04939_WHOLENESS_1000(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1001(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1010(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1011(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1020(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1021(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1030(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1031(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1040(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1041(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1050(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1055(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1060(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1061(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1062(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1070(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1080(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1090(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1100(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1110(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1120(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1130(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1140(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1150(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1160(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1170(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1180(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1190(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1200(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1210(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1220(self.DOC5Workbook, self)

            WholenessTester.Test_02043_18_04939_WHOLENESS_1230(self.DOC5Workbook, self)

            # Coherence checks

            '''Coherence_checksTester.Test_02043_18_04939_COH_2000(self.DOC5Workbook, self)

            Coherence_checksTester.Test_02043_18_04939_COH_2005(self.DOC5Workbook, self)

            Coherence_checksTester.Test_02043_18_04939_COH_2006(self.excelApp, self.DOC5Workbook, self, self.DOC8Name)

            Coherence_checksTester.Test_02043_18_04939_COH_2007(self.excelApp, self.DOC5Workbook, self, self.DOC14Name)

            Coherence_checksTester.Test_02043_18_04939_COH_2010(self.DOC5Workbook, self)

            Coherence_checksTester.Test_02043_18_04939_COH_2020(self.DOC5Workbook, self)

            Coherence_checksTester.Test_02043_18_04939_COH_2030(self.DOC5Workbook, self)

            Coherence_checksTester.Test_02043_18_04939_COH_2040(self.DOC5Workbook, self)

            Coherence_checksTester.Test_02043_18_04939_COH_2050(self.DOC5Workbook, self)'''

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
