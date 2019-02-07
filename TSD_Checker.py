import sys
from PyQt5.QtWidgets import QWidget, QPushButton, QApplication, QComboBox, QLabel, QLineEdit,  QTabWidget, QVBoxLayout, QProgressBar, QRadioButton
from PyQt5 import QtCore, QtWidgets
import openpyxl
import xlrd
import win32com.client as win32
import requests
import os
from win32ui import FindWindow
from ctypes import windll
from openpyxl.styles import Color

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
        self.CesareLink = '''https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02043_18_05471/v.vc/pj'''
        self.TSDConfigLink = '''https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02043_18_05472/v.vc/pj'''
        self.CustomerEffectLink = '''https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02043_18_05499/v.vc/pj'''
        self.DiversityLink = '''https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02016_11_04964/v.vc/pj'''
        self.tabs.addTab(self.tab1, "TSD Checker")
        self.tabs.addTab(self.tab2, "Options")
        self.initUI(self.tab1)
        self.initUIOptions(self.tab2)
        self.layout = QVBoxLayout()
        self.layout.addWidget(self.tabs)
        self.setLayout(self.layout)
        self.setWindowTitle("TSD Checker  V0.4")

    def ToggleLink(self):
        if self.tab2.RadioButtonInternet.isChecked() == True:
            self.CesareLink = '''https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02043_18_05471/v.vc/pj'''
            self.TSDConfigLink = '''https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02043_18_05472/v.vc/pj'''
            self.CustomerEffectLink = '''https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02043_18_05499/v.vc/pj'''
            self.DiversityLink = '''https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.02016_11_04964/v.vc/pj'''
            self.tab2.link2.setText('''<a href=''' + self.CesareLink + '''>DocInfo Reference: 02043_18_05471</a>''')
            self.tab2.link1.setText('''<a href=''' + self.TSDConfigLink + '''>DocInfo Reference: 02043_18_05472</a>''')
            self.tab2.link3.setText('''<a href=''' + self.CustomerEffectLink + '''>DocInfo Reference: 02043_18_05499</a>''')
            self.tab2.link4.setText('''<a href=''' + self.DiversityLink + '''>DocInfo Reference: 02016_11_04964</a>''')
        elif self.tab2.RadioButtonIntranet.isChecked() == True:
            self.CesareLink = "http://docinfogroupe.inetpsa.com/ead/doc/ref.02043_18_05471/v.vc/pj"
            self.TSDConfigLink = "http://docinfogroupe.inetpsa.com/ead/doc/ref.02043_18_05472/v.vc/pj"
            self.CustomerEffectLink = "http://docinfogroupe.inetpsa.com/ead/doc/ref.02043_18_05499/v.vc/pj"
            self.DiversityLink = "http://docinfogroupe.inetpsa.com/ead/doc/ref.02016_11_04964/v.vc/pj"
            self.tab2.link2.setText('''<a href=''' + self.CesareLink + '''>DocInfo Reference: 02043_18_05471</a>''')
            self.tab2.link1.setText('''<a href=''' + self.TSDConfigLink + '''>DocInfo Reference: 02043_18_05472</a>''')
            self.tab2.link3.setText('''<a href=''' + self.CustomerEffectLink + '''>DocInfo Reference: 02043_18_05499</a>''')
            self.tab2.link4.setText('''<a href=''' + self.DiversityLink + '''>DocInfo Reference: 02016_11_04964</a>''')

        else:
            self.tab1.setText("ERROR: Incorrect network type")

    def openFileNameDialog1(self):
        fileName1, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab1, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab1.myTextBox1.setText(fileName1)
        self.tab1.textbox.setText("next file")

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
        fileName7, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab1, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab1.myTextBox7.setText(fileName7)

    def openFileNameDialog8(self):
        fileName8, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab1, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab1.myTextBox8.setText(fileName8)

    def openFileNameDialog9(self):
        fileName9, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab1, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab1.myTextBox9.setText(fileName9)

    def openFileNameDialog10(self):
        fileName10, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab1, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab1.myTextBox10.setText(fileName10)

    def ButtonReportClick(self):
        fileName = "C:\\Users\\admacesanu\\Desktop\\Python_Project\\reportDOC3_XLS.xlsx"
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.excel.Visible = True
        self.excel.Workbooks.Open(fileName)

    def buttonClicked(self):
        return

    def download_file(self, url):
        user = self.tab2.TextBoxUser.text()
        user = str(user)
        password = self.tab2.TextBoxPass.text()
        password = str(password)
        username = os.environ['USERNAME']
        out_path = "C:/Users/" + username + "/AppData/Local/Temp/TSD_Checker/"
        user_path = "C:/Users/" + username
        if not user or not password:
            self.tab1.textbox.setText("Missing Username or Password")
            return
        ''' try:
            os.stat(user_path)
          except:
            self.errorPopUp = QWidget()
            self.errorPopUp.setWindowTitle("ERRROR")
            self.errorPopUp.Label = QLabel(self.errorPopUp)
            self.errorPopUp.Label.setText("Username or Password Incorrect") 
            self.errorPopUp.Label.setAlignment(QtCore.Qt.AlignCenter)
            self.errorPopUp.setGeometry(550, 550, 200, 50)
            self.errorPopUp.show()
            return '''
        try:
            os.stat(out_path)
        except:
            os.mkdir(out_path)

        response = requests.get(url, stream=True, auth=(user, password))
        status = response.status_code
        if status == 401:
            self.tab1.textbox.setText("Username or Password Incorrect")
            return

        FileName = response.headers['Content-Disposition'].split('"')[1]
        FilePath = out_path + "/" + FileName
        success_download = self.tab1.textbox.toPlainText()
        success_download = success_download + "\nfile " + FileName + " has been successfully downloaded\n=======================\n"
        print("Saving file to location:" + FilePath)
        self.tab1.textbox.setText(success_download)
        with open(FilePath, 'wb') as f:
            for chuck in response.iter_content(chunk_size=128):
                f.write(chuck)
        return FilePath

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

    #create a progress bar
        tab.pbar = QProgressBar(self.tab1)
        tab.pbar.setGeometry(10, 310, 700, 20)
        tab.pbar.setAlignment(QtCore.Qt.AlignCenter)
        tab.pbar.setValue(0)
        tab.pbar.move(10, 410)

    #Create a color textbox1
        tab.colorTextBox1 = QtWidgets.QTextEdit(self.tab1)
        tab.colorTextBox1.setStyleSheet(  " background-color: grey ")
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

    #Create a drop down list
        tab.lbl = QLabel("Check level", tab)

        tab.combo = QComboBox(tab)
        tab.combo.addItem("   Previsional   ")
        tab.combo.addItem("   Consolidated   ")
        tab.combo.addItem("   Validated   ")
        tab.combo.resize(508, 20.4)  #rezise the drop down list
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


        #File Selectiom Dialog1
        tab.lbl2 = QLabel("TSD File:", tab)
        tab.lbl2.move(5,15)
        tab.myTextBox1 = QtWidgets.QTextEdit(tab)
        tab.myTextBox1.resize(460, 25)
        tab.myTextBox1.move(200, 10)
        tab.myTextBox1.setReadOnly(True)
        tab.myTextBox1.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        tab.button1 = QPushButton('...',tab)
        tab.button1.clicked.connect(self.openFileNameDialog1)
        tab.button1.move(660, 10)
        tab.button1.resize(45,22)

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



    # File Selectiom Dialog7
        tab.lbl8 = QLabel("AMDEC:", tab)
        tab.lbl8.move(5, 105)
        tab.myTextBox7 = QtWidgets.QTextEdit(tab)
        tab.myTextBox7.resize(460, 25)
        tab.myTextBox7.move(200, 100)
        tab.myTextBox7.setReadOnly(True)
        tab.myTextBox7.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        tab.button7 = QPushButton('...', tab)
        tab.button7.clicked.connect(self.openFileNameDialog7)
        tab.button7.move(660, 100)
        tab.button7.resize(45, 22)

    # File Selectiom Dialog8
        tab.lbl9 = QLabel("export MedialecMatrice:", tab)
        tab.lbl9.move(5, 135)
        tab.myTextBox8 = QtWidgets.QTextEdit(tab)
        tab.myTextBox8.resize(460, 25)
        tab.myTextBox8.move(200, 130)
        tab.myTextBox8.setReadOnly(True)
        tab.button8 = QPushButton('...', tab)
        tab.button8.clicked.connect(self.openFileNameDialog8)
        tab.button8.move(660, 130)
        tab.button8.resize(45, 22)
        tab.myTextBox8.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)



    # File Selectiom Dialog10
        tab.lbl11 = QLabel("Diagnostic matrix file:", tab)
        tab.lbl11.move(5, 165)
        tab.myTextBox10 = QtWidgets.QTextEdit(tab)
        tab.myTextBox10.resize(460, 25)
        tab.myTextBox10.move(200, 160)
        tab.myTextBox10.setReadOnly(True)
        tab.button10 = QPushButton('...', tab)
        tab.button10.clicked.connect(self.openFileNameDialog10)
        tab.button10.move(660, 160)
        tab.button10.resize(45, 22)
        tab.myTextBox10.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)



    # Check button
        button = QPushButton('Check', tab)
        button.move(310, 470)
        button.resize(90,25)
        button.clicked.connect(self.buttonClicked)
        button.setStyleSheet('QPushButton {background-color: white; color: black;}')
        buttonNew = QPushButton("Open \nReport", tab)
        buttonNew.resize(90, 60)
        buttonNew.move(710, 310)
        buttonNew.clicked.connect(self.ButtonReportClick)


        self.show()

    def initUIOptions(self, tab):

        tab.lblUser = QLabel("USER:", tab)
        tab.lblUser.move(165,25)
        tab.TextBoxUser = QtWidgets.QLineEdit(tab)
        tab.TextBoxUser.resize(200,25)
        tab.TextBoxUser.move(200, 20)


        tab.lblPass = QLabel("PASSWORD:", tab)
        tab.lblPass.move(450,25)
        tab.TextBoxPass = QtWidgets.QLineEdit(tab)
        tab.TextBoxPass.resize(180,25)
        tab.TextBoxPass.move(520, 20)
        tab.TextBoxPass.setEchoMode((QLineEdit.Password))


        # File Selectiom Dialog5
        tab.lbl6 = QLabel("Famille/Sous-Famille list export(CESARE):", tab)
        tab.lbl6.move(5, 145)
        tab.myTextBox5 = QtWidgets.QTextEdit(tab)
        tab.myTextBox5.resize(460, 25)
        tab.myTextBox5.move(200, 140)
        tab.myTextBox5.setReadOnly(True)

        tab.link2 = QLabel('''<a href=''' + self.CesareLink + '''>DocInfo Reference: 02043_18_05471</a>''', tab)
        tab.link2.setOpenExternalLinks(True)
        tab.link2.move(720, 145)


        tab.button5 = QPushButton('...', tab)
        tab.button5.move(660, 140)
        tab.button5.resize(45, 22)



        # File Selectiom Dialog4
        tab.lbl5 = QLabel("TSD configuration file:", tab)
        tab.lbl5.move(5,185)
        tab.myTextBox4 = QtWidgets.QTextEdit(tab)
        tab.myTextBox4.resize(460, 25)
        tab.myTextBox4.move(200, 180)
        tab.myTextBox4.setReadOnly(True)

        tab.link1 = QLabel('''<a href='''+self.TSDConfigLink+'''>DocInfo Reference: 02043_18_05472</a>''', tab)
        tab.link1.setOpenExternalLinks(True)
        tab.link1.move(720, 185)

        tab.button4 = QPushButton('...', tab)
        tab.button4.clicked.connect(self.openFileNameDialog4)
        tab.button4.move(660, 180)
        tab.button4.resize(45, 22)



        # File Selectiom Dialog6
        tab.lbl7 = QLabel("Customer effect file:", tab)
        tab.lbl7.move(5, 225)
        tab.myTextBox6 = QtWidgets.QTextEdit(tab)
        tab.myTextBox6.resize(460, 25)
        tab.myTextBox6.move(200, 220)
        tab.myTextBox6.setReadOnly(True)

        tab.link3 = QLabel('''<a href=''' + self.CustomerEffectLink + '''>DocInfo Reference: 02043_18_05499</a>''', tab)
        tab.link3.setOpenExternalLinks(True)
        tab.link3.move(720, 225)



        tab.button6 = QPushButton('...', tab)
        tab.button6.clicked.connect(self.openFileNameDialog6)
        tab.button6.move(660, 220)
        tab.button6.resize(45, 22)

        # File Selectiom Dialog9
        tab.lbl10 = QLabel("Diversity management file:", tab)
        tab.lbl10.move(5, 265)
        tab.myTextBox9 = QtWidgets.QTextEdit(tab)
        tab.myTextBox9.resize(460, 25)
        tab.myTextBox9.move(200,260)
        tab.myTextBox9.setReadOnly(True)

        tab.link4 = QLabel('''<a href=''' + self.DiversityLink + '''>DocInfo Reference: 02016_11_04964</a>''', tab)
        tab.link4.setOpenExternalLinks(True)
        tab.link4.move(720, 265)


        tab.button9 = QPushButton('...', tab)
        tab.button9.clicked.connect(self.openFileNameDialog9)
        tab.button9.move(660, 260)
        tab.button9.resize(45, 22)

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

    def onActivated(self):
        return

class Test(Application):

    def __init__(self):
        super().__init__()
        self.tsdFileExtension = str()
        self.tsdVehicleFunctionFileExtension = str()
        self.tsdSystemFileExtension = str()
        self.amdecFileExtension = str()
        self.exportMedialecMatriceFileExtension = str()
        self.diagnosticMatrixFileExtension = str()
        self.pbvalue = 0
        username = os.environ['USERNAME']
        out_path = "C:/Users/" + username + "/AppData/Local/Temp/TSD_Checker/"
        try:
            os.stat(out_path)
            filelist = [file for file in os.listdir(out_path)]
            for filename in filelist:
                os.remove(out_path + filename)
        except:
            os.mkdir(out_path)

    def TestGeneralStructureXLS(self, workBook, fileName):

        flag = True
        if self.Test_02043_18_04939_STRUCT_0000_XLS(workBook) == 1:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0000 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0000: The sheet “Informations Générales” (or “General information”) is not present or not written correctly."
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0005_XLS(fileName) == 1:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0005 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0005: The field “REFERENCE” of the sheet “Informations Générales” (or “General information”)  in the line 52 shall be indicated."
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0010_XLS(workBook) == 1:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0010 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0010: The information “Ref plan type” is missing in the sheet “Informations Générales” (or “General information”)."
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0011_XLS(workBook) == 1:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0011 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0011: The document does not specify the template or the template reference is not indicated in the sheet “Informations Générales” (or “General information”). \nAs indicated in to one of the 3 references AEEV_IAEE07_0033 or 02043_12_01665 or 02043_12_01666"
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0020_XLS(workBook) == 1:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0020 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0020: The sheet “Suppression ” (or “suppression ”) is not present or not written correctly."
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0025_XLS(workBook) == 1:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0025 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0025: The document does not follow the template, the column “Onglet” (or “sheet”) of the sheet “Suppression” (or “suppression”) is not present or not written correctly."
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)


        if self.Test_02043_18_04939_STRUCT_0030_XLS(workBook) == 1:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0030 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0030: The document does not follow the template, the column “Référence de la ligne” (or “Line number”) of the sheet “Suppression” (or “suppression”) is not present or not written correctly."
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0035_XLS(workBook) == 1:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0035 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0035: The document does not follow the template, the column “Version du TSD” (or “Version of the document”) of the sheet “Suppression” (or “suppression”) is not present or not written correctly."
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)


        if self.Test_02043_18_04939_STRUCT_0040_XLS(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0040 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0040: The document does not follow the template, the column “Justification de la modification” (or “Change reason”) of the sheet “Suppression” (or “suppression”) is not present or not written correctly. "
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)


        if self.Test_02043_18_04939_STRUCT_0051_XLS(workBook) == 1:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0051 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0051: The “Vehicle Architecture Schematic” document is not referenced. \nAs indicated in to one of the 3 references AEEV_IAEE07_0033 or 02043_12_01665 or 02043_12_01666 "
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)


        if self.Test_02043_18_04939_STRUCT_0052_XLS(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0052 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0052: The “Diagnostic Matrix” document is not referenced. As indicated in to one of the 3 references AEEV_IAEE07_0033 or 02043_12_01665 or 02043_12_01666"
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0053_XLS(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0053 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0053: The “Fault Tree” document is not referenced. As indicated in to one of the 3 references AEEV_IAEE07_0033 or 02043_12_01665 or 02043_12_01666 "
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0054_XLS(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0054 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0054: The “ECU schematic” document is not referenced. As indicated in to one of the 3 references AEEV_IAEE07_0033 or 02043_12_01665 or 02043_12_01666 "
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0055_XLS(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0055 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0055: The “STD” document is not referenced. As indicated in to one of the 3 references AEEV_IAEE07_0033 or 02043_12_01665 or 02043_12_01666 "
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)


        if self.Test_02043_18_04939_STRUCT_0056_XLS(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0056 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0056: The “Complexity Matrix (Decli EE)” document is not referenced. As indicated in to one of the 3 references AEEV_IAEE07_0033 or 02043_12_01665 or 02043_12_01666"
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0057_XLS(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0057 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0057: The “Décli” document is not referenced. As indicated in to one of the 3 references AEEV_IAEE07_0033 or 02043_12_01665 or 02043_12_01666"
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0058_XLS(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0058 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0058:  The “DCEE” document is not referenced. As indicated in to one of the 3 references AEEV_IAEE07_0033 or 02043_12_01665 or 02043_12_01666"
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0059_XLS(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0059 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0059: The “EEAD” document is not referenced. As indicated in to one of the 3 references AEEV_IAEE07_0033 or 02043_12_01665 or 02043_12_01666"
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0060_XLS(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0060 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0060: The “TFD” document is not referenced. As indicated in to one of the 3 references AEEV_IAEE07_0033 or 02043_12_01665 or 02043_12_01666 "
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)
        return flag

    def TestGeneralStructureXLSX_XLSM(self, workBook):

        flag = True

        if self.Test_02043_18_04939_STRUCT_0000_XLSX_XLSM(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0000 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0000: The sheet “Informations Générales” (or “General information”) is not present or not written correctly."
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0005_XLSX_XLSM(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0005 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0005: The field “REFERENCE” of the sheet “Informations Générales” (or “General information”)  in the line 52 shall be indicated."
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0010_XLSX_XLSM(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0010 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0010: The information “Ref plan type” is missing in the sheet “Informations Générales” (or “General information”)."
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0011_XLSX_XLSM(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0011 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0011: The document does not specify the template or the template reference is not indicated in the sheet “Informations Générales” (or “General information”). \nAs indicated in to one of the 3 references AEEV_IAEE07_0033 or 02043_12_01665 or 02043_12_01666"
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0020_XLSX_XLSM(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0020 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0020: The sheet “Suppression ” (or “suppression ”) is not present or not written correctly."
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0025_XLSX_XLSM(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0025 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0025: The document does not follow the template, the column “Onglet” (or “sheet”) of the sheet “Suppression” (or “suppression”) is not present or not written correctly."
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0030_XLSX_XLSM(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0030 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0030: The document does not follow the template, the column “Référence de la ligne” (or “Line number”) of the sheet “Suppression” (or “suppression”) is not present or not written correctly."
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)


        if self.Test_02043_18_04939_STRUCT_0035_XLSX_XLSM(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0035 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0035: The document does not follow the template, the column “Version du TSD” (or “Version of the document”) of the sheet “Suppression” (or “suppression”) is not present or not written correctly."
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)


        if self.Test_02043_18_04939_STRUCT_0040_XLSX_XLSM(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0040 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0040:  The document does not follow the template, the column “Justification de la modification” (or “Change reason”) of the sheet “Suppression” (or “suppression”) is not present or not written correctly. "
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)


        if self.Test_02043_18_04939_STRUCT_0051_XLSX_XLSM(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0051 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0051: The “Vehicle Architecture Schematic” document is not referenced. \nAs indicated in to one of the 3 references AEEV_IAEE07_0033 or 02043_12_01665 or 02043_12_01666 "
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)



        if self.Test_02043_18_04939_STRUCT_0052_XLSX_XLSM(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0052 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0052: The “Diagnostic Matrix” document is not referenced. As indicated in to one of the 3 references AEEV_IAEE07_0033 or 02043_12_01665 or 02043_12_01666"
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)



        if self.Test_02043_18_04939_STRUCT_0053_XLSX_XLSM(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0053 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0053: The “Fault Tree” document is not referenced. As indicated in to one of the 3 references AEEV_IAEE07_0033 or 02043_12_01665 or 02043_12_01666 "
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)



        if self.Test_02043_18_04939_STRUCT_0054_XLSX_XLSM(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0054 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0054: The “ECU schematic” document is not referenced. As indicated in to one of the 3 references AEEV_IAEE07_0033 or 02043_12_01665 or 02043_12_01666 "
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)



        if self.Test_02043_18_04939_STRUCT_0055_XLSX_XLSM(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0055 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0055: The “STD” document is not referenced. As indicated in to one of the 3 references AEEV_IAEE07_0033 or 02043_12_01665 or 02043_12_01666 "
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0056_XLSX_XLSM(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0056 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0056: The “Complexity Matrix (Decli EE)” document is not referenced. As indicated in to one of the 3 references AEEV_IAEE07_0033 or 02043_12_01665 or 02043_12_01666"
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0057_XLSX_XLSM(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0057 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0057: The “Décli” document is not referenced. As indicated in to one of the 3 references AEEV_IAEE07_0033 or 02043_12_01665 or 02043_12_01666"
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0058_XLSX_XLSM(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0058 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0058:  The “DCEE” document is not referenced. As indicated in to one of the 3 references AEEV_IAEE07_0033 or 02043_12_01665 or 02043_12_01666"
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0059_XLSX_XLSM(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0059 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0059: The “EEAD” document is not referenced. As indicated in to one of the 3 references AEEV_IAEE07_0033 or 02043_12_01665 or 02043_12_01666"
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)

        if self.Test_02043_18_04939_STRUCT_0060_XLSX_XLSM(workBook):
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0060 OK"
            self.tab1.textbox.setText(text)
        else:
            text = self.tab1.textbox.toPlainText()
            text = text + "\nTest_02043_18_04939_STRUCT_0060: The “TFD” document is not referenced. As indicated in to one of the 3 references AEEV_IAEE07_0033 or 02043_12_01665 or 02043_12_01666 "
            self.tab1.textbox.setText(text)
            flag = False
        self.pbvalue = self.pbvalue + 0.8772
        self.tab1.pbar.setValue(self.pbvalue)
        return flag

    def TestGeneralStructure_DOC5_XLS(self, workBook):

        flag = 1
        filename = "C:\\Users\\admacesanu\\Desktop\\Python_Project\\report.xlsx"
        reportWorkBook = openpyxl.Workbook()
        reportWorkSheet1 = reportWorkBook.active
        reportWorkSheet1.title = "Report Information"
        reportWorkSheet2 = reportWorkBook.create_sheet("Test Report")
        reportInformationCol1StringList = ["Tool version:", "Criticity configuration file:", "", "Extract CESARE file:",
                                           "Customer effects file:", "check level:", "", "Date of the test:", "Time of the test:",
                                           "", "TSD file checked:", "TSD function file checked:", "TSD system file checked:",
                                           "", "AMDEC:", "export MedialecMatrice:", "", "Status:"]
        for rowIndex in range(1, len(reportInformationCol1StringList)+1):
            reportWorkSheet1.cell(row = rowIndex, column = 1, value = reportInformationCol1StringList[rowIndex-1])

        testReportRow1StringList = ["Criticity", "Requirements", "Message", "Localisation"]
        for colIndex in range(1, len(testReportRow1StringList)+1):
            reportWorkSheet2.cell(row = 1, column = colIndex, value = testReportRow1StringList[colIndex-1])


        testResult = self.Test_02043_18_04939_STRUCT_0700_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0700","",""])
            reportWorkSheet2.cell(row=2, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00", fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0700", testResult, ""])
            reportWorkSheet2.cell(row=2, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")



        testResult = self.Test_02043_18_04939_STRUCT_0710_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0710","",""])
            reportWorkSheet2.cell(row=3, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00", fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0710", testResult, ""])
            reportWorkSheet2.cell(row=3, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")



        testResult = self.Test_02043_18_04939_STRUCT_0720_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0720", "", ""])
            reportWorkSheet2.cell(row=4, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0720", testResult, ""])
            reportWorkSheet2.cell(row=4, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")



        testResult = self.Test_02043_18_04939_STRUCT_0730_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0730", "", ""])
            reportWorkSheet2.cell(row=5, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00", fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0730", testResult, ""])
            reportWorkSheet2.cell(row=5, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")



        testResult = self.Test_02043_18_04939_STRUCT_0740_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0740", "", ""])
            reportWorkSheet2.cell(row=6, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0740", testResult, ""])
            reportWorkSheet2.cell(row=6, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")



        testResult = self.Test_02043_18_04939_STRUCT_0750_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0750", "", ""])
            reportWorkSheet2.cell(row=7, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0750", testResult, ""])
            reportWorkSheet2.cell(row=7, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")



        testResult = self.Test_02043_18_04939_STRUCT_0760_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0760", "", ""])
            reportWorkSheet2.cell(row=8, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0760", testResult, ""])
            reportWorkSheet2.cell(row=8, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")



        testResult = self.Test_02043_18_04939_STRUCT_0770_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0770", "", ""])
            reportWorkSheet2.cell(row=9, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0770", testResult, ""])
            reportWorkSheet2.cell(row=9, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")



        testResult = self.Test_02043_18_04939_STRUCT_0780_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0780", "", ""])
            reportWorkSheet2.cell(row=10, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0780", testResult, ""])
            reportWorkSheet2.cell(row=10, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")



        testResult = self.Test_02043_18_04939_STRUCT_0790_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0790", "", ""])
            reportWorkSheet2.cell(row=11, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0790", testResult, ""])
            reportWorkSheet2.cell(row=11, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")



        testResult = self.Test_02043_18_04939_STRUCT_0800_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0800", "", ""])
            reportWorkSheet2.cell(row=12, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0800", testResult, ""])
            reportWorkSheet2.cell(row=12, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")



        testResult = self.Test_02043_18_04939_STRUCT_0810_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0810", "", ""])
            reportWorkSheet2.cell(row=13, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0810", testResult, ""])
            reportWorkSheet2.cell(row=13, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")



        testResult = self.Test_02043_18_04939_STRUCT_0820_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0820", "", ""])
            reportWorkSheet2.cell(row=14, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0820", testResult, ""])
            reportWorkSheet2.cell(row=14, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")



        testResult = self.Test_02043_18_04939_STRUCT_0830_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0830", "", ""])
            reportWorkSheet2.cell(row=15, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0830", testResult, ""])
            reportWorkSheet2.cell(row=15, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")



        testResult = self.Test_02043_18_04939_STRUCT_0840_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0840", "", ""])
            reportWorkSheet2.cell(row=16, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0840", testResult, ""])
            reportWorkSheet2.cell(row=16, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")



        testResult = self.Test_02043_18_04939_STRUCT_0850_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0850", "", ""])
            reportWorkSheet2.cell(row=17, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0850", testResult, ""])
            reportWorkSheet2.cell(row=17, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")



        testResult = self.Test_02043_18_04939_STRUCT_0860_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0860", "", ""])
            reportWorkSheet2.cell(row=18, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0860", testResult, ""])
            reportWorkSheet2.cell(row=18, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")



        testResult = self.Test_02043_18_04939_STRUCT_0870_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0870", "", ""])
            reportWorkSheet2.cell(row=19, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0870", testResult, ""])
            reportWorkSheet2.cell(row=19, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")



        testResult = self.Test_02043_18_04939_STRUCT_0880_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0880", "", ""])
            reportWorkSheet2.cell(row=20, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0880", testResult, ""])
            reportWorkSheet2.cell(row=20, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")



        testResult = self.Test_02043_18_04939_STRUCT_0890_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0890", "", ""])
            reportWorkSheet2.cell(row=21, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0890", testResult, ""])
            reportWorkSheet2.cell(row=21, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")



        testResult = self.Test_02043_18_04939_STRUCT_0900_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0900", "", ""])
            reportWorkSheet2.cell(row=22, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0900", testResult, ""])
            reportWorkSheet2.cell(row=22, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")



        testResult = self.Test_02043_18_04939_STRUCT_0910_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0910", "", ""])
            reportWorkSheet2.cell(row=23, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0910", testResult, ""])
            reportWorkSheet2.cell(row=23, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        reportWorkBook.save(filename)

    def TestGeneralStructure_DOC5_XLSX_XLSM(self, workBook):

        flag = 1
        filename = "C:\\Users\\admacesanu\\Desktop\\Python_Project\\reportXLSX.xlsx"
        reportWorkBook = openpyxl.Workbook()
        reportWorkSheet1 = reportWorkBook.active
        reportWorkSheet1.title = "Report Information"
        reportWorkSheet2 = reportWorkBook.create_sheet("Test Report")
        reportInformationCol1StringList = ["Tool version:", "Criticity configuration file:", "", "Extract CESARE file:",
                                           "Customer effects file:", "check level:", "", "Date of the test:",
                                           "Time of the test:",
                                           "", "TSD file checked:", "TSD function file checked:",
                                           "TSD system file checked:",
                                           "", "AMDEC:", "export MedialecMatrice:", "", "Status:"]
        for rowIndex in range(1, len(reportInformationCol1StringList) + 1):
            reportWorkSheet1.cell(row=rowIndex, column=1, value=reportInformationCol1StringList[rowIndex - 1])

        testReportRow1StringList = ["Criticity", "Requirements", "Message", "Localisation"]
        for colIndex in range(1, len(testReportRow1StringList) + 1):
            reportWorkSheet2.cell(row=1, column=colIndex, value=testReportRow1StringList[colIndex - 1])

        testResult = self.Test_02043_18_04939_STRUCT_0700_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0700", "", ""])
            reportWorkSheet2.cell(row=2, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0700", testResult, ""])
            reportWorkSheet2.cell(row=2, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0710_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0710", "", ""])
            reportWorkSheet2.cell(row=3, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0710", testResult, ""])
            reportWorkSheet2.cell(row=3, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0720_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0720", "", ""])
            reportWorkSheet2.cell(row=4, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0720", testResult, ""])
            reportWorkSheet2.cell(row=4, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0730_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0730", "", ""])
            reportWorkSheet2.cell(row=5, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0730", testResult, ""])
            reportWorkSheet2.cell(row=5, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0740_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0740", "", ""])
            reportWorkSheet2.cell(row=6, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0740", testResult, ""])
            reportWorkSheet2.cell(row=6, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0750_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0750", "", ""])
            reportWorkSheet2.cell(row=7, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0750", testResult, ""])
            reportWorkSheet2.cell(row=7, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0760_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0760", "", ""])
            reportWorkSheet2.cell(row=8, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0760", testResult, ""])
            reportWorkSheet2.cell(row=8, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0770_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0770", "", ""])
            reportWorkSheet2.cell(row=9, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0770", testResult, ""])
            reportWorkSheet2.cell(row=9, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0780_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0780", "", ""])
            reportWorkSheet2.cell(row=10, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0780", testResult, ""])
            reportWorkSheet2.cell(row=10, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0790_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0790", "", ""])
            reportWorkSheet2.cell(row=11, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0790", testResult, ""])
            reportWorkSheet2.cell(row=11, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0800_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0800", "", ""])
            reportWorkSheet2.cell(row=12, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0800", testResult, ""])
            reportWorkSheet2.cell(row=12, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0810_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0810", "", ""])
            reportWorkSheet2.cell(row=13, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0810", testResult, ""])
            reportWorkSheet2.cell(row=13, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0820_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0820", "", ""])
            reportWorkSheet2.cell(row=14, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0820", testResult, ""])
            reportWorkSheet2.cell(row=14, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0830_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0830", "", ""])
            reportWorkSheet2.cell(row=15, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0830", testResult, ""])
            reportWorkSheet2.cell(row=15, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0840_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0840", "", ""])
            reportWorkSheet2.cell(row=16, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0840", testResult, ""])
            reportWorkSheet2.cell(row=16, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0850_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0850", "", ""])
            reportWorkSheet2.cell(row=17, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0850", testResult, ""])
            reportWorkSheet2.cell(row=17, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0860_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0860", "", ""])
            reportWorkSheet2.cell(row=18, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0860", testResult, ""])
            reportWorkSheet2.cell(row=18, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0870_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0870", "", ""])
            reportWorkSheet2.cell(row=19, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0870", testResult, ""])
            reportWorkSheet2.cell(row=19, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0880_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0880", "", ""])
            reportWorkSheet2.cell(row=20, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0880", testResult, ""])
            reportWorkSheet2.cell(row=20, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0890_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0890", "", ""])
            reportWorkSheet2.cell(row=21, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0890", testResult, ""])
            reportWorkSheet2.cell(row=21, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0900_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0900", "", ""])
            reportWorkSheet2.cell(row=22, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0900", testResult, ""])
            reportWorkSheet2.cell(row=22, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0910_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0910", "", ""])
            reportWorkSheet2.cell(row=23, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")

        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0910", testResult, ""])
            reportWorkSheet2.cell(row=23, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        reportWorkBook.save(filename)

    def TestGeneralStructure_DOC3_XLS(self,workBook):

        flag = 1
        filename = "C:\\Users\\admacesanu\\Desktop\\Python_Project\\reportDOC3_XLS.xlsx"
        reportWorkBook = openpyxl.Workbook()
        reportWorkSheet1 = reportWorkBook.active
        reportWorkSheet1.title = "Report Information"
        reportWorkSheet2 = reportWorkBook.create_sheet("Test Report")
        reportInformationCol1StringList = ["Tool version:", "Criticity configuration file:", "", "Extract CESARE file:",
                                           "Customer effects file:", "check level:", "", "Date of the test:",
                                           "Time of the test:",
                                           "", "TSD file checked:", "TSD function file checked:",
                                           "TSD system file checked:",
                                           "", "AMDEC:", "export MedialecMatrice:", "", "Status:"]
        for rowIndex in range(1, len(reportInformationCol1StringList) + 1):
            reportWorkSheet1.cell(row=rowIndex, column=1, value=reportInformationCol1StringList[rowIndex - 1])

        testReportRow1StringList = ["Criticity", "Requirements", "Message", "Localisation"]
        for colIndex in range(1, len(testReportRow1StringList) + 1):
            reportWorkSheet2.cell(row=1, column=colIndex, value=testReportRow1StringList[colIndex - 1])


        testResult = self.Test_02043_18_04939_STRUCT_0100_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0100", "", ""])
            reportWorkSheet2.cell(row=2, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0100", testResult, ""])
            reportWorkSheet2.cell(row=2, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0120_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0120", "", ""])
            reportWorkSheet2.cell(row=3, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0120", testResult, ""])
            reportWorkSheet2.cell(row=3, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0130_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0130", "", ""])
            reportWorkSheet2.cell(row=4, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00", fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0130", testResult, ""])
            reportWorkSheet2.cell(row=4, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0140_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0140", "", ""])
            reportWorkSheet2.cell(row=5, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0140", testResult, ""])
            reportWorkSheet2.cell(row=5, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0150_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0150", "", ""])
            reportWorkSheet2.cell(row=6, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0150", testResult, ""])
            reportWorkSheet2.cell(row=6, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0160_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0160", "", ""])
            reportWorkSheet2.cell(row=7, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0160", testResult, ""])
            reportWorkSheet2.cell(row=7, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0170_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0170", "", ""])
            reportWorkSheet2.cell(row=8, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0170", testResult, ""])
            reportWorkSheet2.cell(row=8, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0180_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0180", "", ""])
            reportWorkSheet2.cell(row=9, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0180", testResult, ""])
            reportWorkSheet2.cell(row=9, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0190_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0190", "", ""])
            reportWorkSheet2.cell(row=10, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0190", testResult, ""])
            reportWorkSheet2.cell(row=10, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0200_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0200", "", ""])
            reportWorkSheet2.cell(row=11, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0200", testResult, ""])
            reportWorkSheet2.cell(row=11, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0210_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0200", "", ""])
            reportWorkSheet2.cell(row=12, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00", fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0200", testResult, ""])
            reportWorkSheet2.cell(row=12, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0220_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0220", "", ""])
            reportWorkSheet2.cell(row=13, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0220", testResult, ""])
            reportWorkSheet2.cell(row=13, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0230_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0230", "", ""])
            reportWorkSheet2.cell(row=14, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00", fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0230", testResult, ""])
            reportWorkSheet2.cell(row=14, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0240_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0240", "", ""])
            reportWorkSheet2.cell(row=15, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0240", testResult, ""])
            reportWorkSheet2.cell(row=15, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0250_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0250", "", ""])
            reportWorkSheet2.cell(row=16, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0250", testResult, ""])
            reportWorkSheet2.cell(row=16, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0260_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0260", "", ""])
            reportWorkSheet2.cell(row=17, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0260", testResult, ""])
            reportWorkSheet2.cell(row=17, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")


        testResult = self.Test_02043_18_04939_STRUCT_0270_XLS(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0270", "", ""])
            reportWorkSheet2.cell(row=18, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0270", testResult, ""])
            reportWorkSheet2.cell(row=18, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        reportWorkBook.save(filename)

    def TestGeneralStructure_DOC3_XLSX_XLSM(self,workBook):

        flag = 1
        filename = "C:\\Users\\admacesanu\\Desktop\\Python_Project\\reportDOC3_XLSX_XLSM.xlsx"
        reportWorkBook = openpyxl.Workbook()
        reportWorkSheet1 = reportWorkBook.active
        reportWorkSheet1.title = "Report Information"
        reportWorkSheet2 = reportWorkBook.create_sheet("Test Report")
        reportInformationCol1StringList = ["Tool version:", "Criticity configuration file:", "", "Extract CESARE file:",
                                           "Customer effects file:", "check level:", "", "Date of the test:",
                                           "Time of the test:",
                                           "", "TSD file checked:", "TSD function file checked:",
                                           "TSD system file checked:",
                                           "", "AMDEC:", "export MedialecMatrice:", "", "Status:"]
        for rowIndex in range(1, len(reportInformationCol1StringList) + 1):
            reportWorkSheet1.cell(row=rowIndex, column=1, value=reportInformationCol1StringList[rowIndex - 1])

        testReportRow1StringList = ["Criticity", "Requirements", "Message", "Localisation"]
        for colIndex in range(1, len(testReportRow1StringList) + 1):
            reportWorkSheet2.cell(row=1, column=colIndex, value=testReportRow1StringList[colIndex - 1])

        #reportWorkSheet2.cell(row=1, column=1).font = openpyxl.styles.Font(bold=True, italic=True)
       # reportWorkSheet2.cell(row=1, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FFFF00", fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0100_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0100", "", ""])
            reportWorkSheet2.cell(row=2, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0100", testResult, ""])
            reportWorkSheet2.cell(row=2, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000", fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0120_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0120", "", ""])
            reportWorkSheet2.cell(row=3, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0120", testResult, ""])
            reportWorkSheet2.cell(row=3, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0130_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0130", "", ""])
            reportWorkSheet2.cell(row=4, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0130", testResult, ""])
            reportWorkSheet2.cell(row=4, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0140_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0140", "", ""])
            reportWorkSheet2.cell(row=5, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0140", testResult, ""])
            reportWorkSheet2.cell(row=5, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000", fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0150_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0150", "", ""])
            reportWorkSheet2.cell(row=6, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0150", testResult, ""])
            reportWorkSheet2.cell(row=6, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0160_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0160", "", ""])
            reportWorkSheet2.cell(row=7, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0160", testResult, ""])
            reportWorkSheet2.cell(row=7, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0170_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0170", "", ""])
            reportWorkSheet2.cell(row=8, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0170", testResult, ""])
            reportWorkSheet2.cell(row=8, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000", fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0180_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0180", "", ""])
            reportWorkSheet2.cell(row=9, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00", fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0180", testResult, ""])
            reportWorkSheet2.cell(row=9, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0190_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0190", "", ""])
            reportWorkSheet2.cell(row=10, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0190", testResult, ""])
            reportWorkSheet2.cell(row=10, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0200_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0200", "", ""])
            reportWorkSheet2.cell(row=11, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0200", testResult, ""])
            reportWorkSheet2.cell(row=11, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0210_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0200", "", ""])
            reportWorkSheet2.cell(row=12, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0200", testResult, ""])
            reportWorkSheet2.cell(row=12, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0220_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0220", "", ""])
            reportWorkSheet2.cell(row=13, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0220", testResult, ""])
            reportWorkSheet2.cell(row=13, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0230_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0230", "", ""])
            reportWorkSheet2.cell(row=14, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0230", testResult, ""])
            reportWorkSheet2.cell(row=14, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0240_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0240", "", ""])
            reportWorkSheet2.cell(row=15, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0240", testResult, ""])
            reportWorkSheet2.cell(row=15, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0250_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0250", "", ""])
            reportWorkSheet2.cell(row=16, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0250", testResult, ""])
            reportWorkSheet2.cell(row=16, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0260_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0260", "", ""])
            reportWorkSheet2.cell(row=17, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0260", testResult, ""])
            reportWorkSheet2.cell(row=17, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        testResult = self.Test_02043_18_04939_STRUCT_0270_XLSX_XLSM(workBook)
        if testResult == 1:
            reportWorkSheet2.append(["Good", "02043_18_04939_STRUCT_0270", "", ""])
            reportWorkSheet2.cell(row=18, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="00FF00",fill_type="solid")
        else:
            flag = 0
            reportWorkSheet2.append(["Blocking", "02043_18_04939_STRUCT_0270", testResult, ""])
            reportWorkSheet2.cell(row=18, column=1).fill = openpyxl.styles.fills.PatternFill(fgColor="FF0000",fill_type="solid")

        reportWorkBook.save(filename)

    def GetTsdFileExtension(self):

        fileName = self.tab1.myTextBox1.toPlainText()
        tokens = fileName.split(".")
        self.tsdFileExtension = tokens[-1]

    def GetTsdFileWorkbook(self):

        fileName = self.tab1.myTextBox1.toPlainText()
        fileName = fileName.replace("\\", "/")
        if self.tsdFileExtension == "xls":
            return xlrd.open_workbook(fileName,  formatting_info=True)
        elif self.tsdFileExtension == "xlsx":
            return openpyxl.load_workbook(fileName, read_only=True)
        elif self.tsdFileExtension == "xlsm":
            return openpyxl.load_workbook(fileName, keep_vba=True, read_only=True)

    def TestTsdFile(self):

        self.GetTsdFileExtension()
        flag = True
        if self.tsdFileExtension:
            text = self.tab1.textbox.toPlainText()
            text = text + "\n\tTesting TSD FILE ----------------------------------\n"
            self.tab1.textbox.setText(text)
            workBook = self.GetTsdFileWorkbook()
            if self.tsdFileExtension == "xls":
                flag = self.TestGeneralStructureXLS(workBook, self.tab1.myTextBox1.toPlainText())
                self.TestGeneralStructure_DOC3_XLS(workBook)
            else:
                flag = self.TestGeneralStructureXLSX_XLSM(workBook)
                self.TestGeneralStructure_DOC3_XLSX_XLSM(workBook)
            if flag == True:
                self.tab1.colorTextBox1.setStyleSheet('background-color: green')
            else:
                self.tab1.colorTextBox1.setStyleSheet('background-color: red')
        else:
            self.pbvalue = self.pbvalue + 0.8772*19
            self.tab1.pbar.setValue(self.pbvalue)

    def GetTsdVehicleFunctionFileExtension(self):
        fileName = self.tab1.myTextBox2.toPlainText()
        tokens = fileName.split(".")
        self.tsdVehicleFunctionFileExtension = tokens[-1]

    def GetTsdVehicleFunctionFileWorkbook(self):

        fileName = self.tab1.myTextBox2.toPlainText()
        fileName = fileName.replace("\\", "/")
        if self.tsdVehicleFunctionFileExtension == "xls":
            return xlrd.open_workbook(fileName,  formatting_info=True)
        elif self.tsdVehicleFunctionFileExtension == "xlsx":
            return openpyxl.load_workbook(fileName, read_only=True)
        elif self.tsdVehicleFunctionFileExtension == "xlsm":
            return openpyxl.load_workbook(fileName, keep_vba=True, read_only=True)

    def TestTsdVehicleFunctionFile(self):

        self.GetTsdVehicleFunctionFileExtension()
        flag = True
        if self.tsdVehicleFunctionFileExtension:
            text = self.tab1.textbox.toPlainText()
            text = text + "\n\tTesting TSD Vehicle Function FILE --------------------------\n"
            self.tab1.textbox.setText(text)
            workBook = self.GetTsdVehicleFunctionFileWorkbook()
            if self.tsdVehicleFunctionFileExtension == "xls":
                flag = self.TestGeneralStructureXLS(workBook, self.tab1.myTextBox2.toPlainText())
            else:
                flag = self.TestGeneralStructureXLSX_XLSM(workBook)
            if flag == True:
                self.tab1.colorTextBox2.setStyleSheet('background-color: green')
            else:
                self.tab1.colorTextBox2.setStyleSheet('background-color: red')
        else:
            self.pbvalue = self.pbvalue + 0.8772*19
            self.tab1.pbar.setValue(self.pbvalue)

    def GetTsdSystemFileExtension(self):
        fileName = self.tab1.myTextBox3.toPlainText()
        tokens = fileName.split(".")
        self.tsdSystemFileExtension = tokens[-1]

    def GetTsdSystemFileWorkbook(self):

        fileName = self.tab1.myTextBox3.toPlainText()
        if self.tsdSystemFileExtension == "xls":
            return xlrd.open_workbook(fileName,  formatting_info=True)
        elif self.tsdSystemFileExtension == "xlsx":
            return openpyxl.load_workbook(fileName, read_only=True)
        elif self.tsdSystemFileExtension == "xlsm":
            return openpyxl.load_workbook(fileName, keep_vba=True, read_only=True)

    def TestTsdSystemFile(self):

        self.GetTsdSystemFileExtension()
        if self.tsdSystemFileExtension:
            text = self.tab1.textbox.toPlainText()
            text = text + "\n\tTesting TSD System FILE ---------------------------------\n"
            self.tab1.textbox.setText(text)
            workBook = self.GetTsdSystemFileWorkbook()
            if self.tsdSystemFileExtension == "xls":
                flag = self.TestGeneralStructureXLS(workBook, self.tab1.myTextBox3.toPlainText())
                flag = flag and self.TestGeneralStructure_DOC5_XLS(workBook)
            else:
                flag = self.TestGeneralStructureXLSX_XLSM(workBook)
                flag = flag and self.TestGeneralStructure_DOC5_XLSX_XLSM(workBook)
            if flag == True:
                self.tab1.colorTextBox3.setStyleSheet('background-color: green')
            else:
                self.tab1.colorTextBox3.setStyleSheet('background-color: red')
        else:
            self.pbvalue = self.pbvalue + 0.8772*19
            self.tab1.pbar.setValue(self.pbvalue)

    def GetAmdecFileExtension(self):
        fileName = self.tab1.myTextBox7.toPlainText()
        tokens = fileName.split(".")
        self.amdecFileExtension = tokens[-1]

    def GetAmdecFileWorkbook(self):

        fileName = self.tab1.myTextBox7.toPlainText()
        if self.amdecFileExtension == "xls":
            return xlrd.open_workbook(fileName,  formatting_info=True)
        elif self.amdecFileExtension == "xlsx":
            return openpyxl.load_workbook(fileName, read_only=True)
        elif self.amdecFileExtension == "xlsm":
            return openpyxl.load_workbook(fileName, keep_vba=True, read_only=True)

    def TestAmdecFile(self):

        self.GetAmdecFileExtension()
        if self.amdecFileExtension:
            text = self.tab1.textbox.toPlainText()
            text = text + "\n\tTesting AMDEC FILE -------------------------------------\n"
            self.tab1.textbox.setText(text)
            workBook = self.GetAmdecFileWorkbook()
            if self.amdecFileExtension == "xls":
                flag = self.TestGeneralStructureXLS(workBook, self.tab1.myTextBox7.toPlainText())
            else:
                flag = self.TestGeneralStructureXLSX_XLSM(workBook)
            if flag == True:
                self.tab1.colorTextBox4.setStyleSheet('background-color: green')
            else:
                self.tab1.colorTextBox4.setStyleSheet('background-color: red')
        else:
            self.pbvalue = self.pbvalue + 0.8772*19
            self.tab1.pbar.setValue(self.pbvalue)

    def GetExportMedialecMatriceFileExtension(self):
        fileName = self.tab1.myTextBox8.toPlainText()
        tokens = fileName.split(".")
        self.exportMedialecMatriceFileExtension = tokens[-1]

    def GetExportMedialecMatriceFileWorkbook(self):

        fileName = self.tab1.myTextBox8.toPlainText()
        if self.exportMedialecMatriceFileExtension == "xls":
            return xlrd.open_workbook(fileName,  formatting_info=True)
        elif self.exportMedialecMatriceFileExtension == "xlsx":
            return openpyxl.load_workbook(fileName, read_only=True)
        elif self.exportMedialecMatriceFileExtension == "xlsm":
            return openpyxl.load_workbook(fileName, keep_vba=True, read_only=True)

    def TestExportMedialecMatriceFile(self):

        self.GetExportMedialecMatriceFileExtension()
        if self.exportMedialecMatriceFileExtension:
            text = self.tab1.textbox.toPlainText()
            text = text + "\n\tTesting Export Medialec Matrice FILE --------------------------\n"
            self.tab1.textbox.setText(text)
            workBook = self.GetExportMedialecMatriceFileWorkbook()
            if self.exportMedialecMatriceFileExtension == "xls":
                flag = self.TestGeneralStructureXLS(workBook, self.tab1.myTextBox8.toPlainText())
            else:
                flag = self.TestGeneralStructureXLSX_XLSM(workBook)
            if flag == True:
                self.tab1.colorTextBox5.setStyleSheet('background-color: green')
            else:
                self.tab1.colorTextBox5.setStyleSheet('background-color: red')
        else:
            self.pbvalue = self.pbvalue + 0.8772*19
            self.tab1.pbar.setValue(self.pbvalue)

    def GetDiagnosticMatrixFileExtension(self):
        fileName = self.tab1.myTextBox10.toPlainText()
        tokens = fileName.split(".")
        self.diagnosticMatrixFileExtension = tokens[-1]

    def GetDiagnosticMatrixFileWorkbook(self):

        fileName = self.tab1.myTextBox10.toPlainText()
        if self.diagnosticMatrixFileExtension == "xls":
            return xlrd.open_workbook(fileName,  formatting_info=True)
        elif self.diagnosticMatrixFileExtension == "xlsx":
            return openpyxl.load_workbook(fileName, read_only=True)
        elif self.diagnosticMatrixFileExtension == "xlsm":
            return openpyxl.load_workbook(fileName, keep_vba=True, read_only=True)

    def TestDiagnosticMatrixFile(self):

        self.GetDiagnosticMatrixFileExtension()
        if self.diagnosticMatrixFileExtension:
            text = self.tab1.textbox.toPlainText()
            text = text + "\n\tTesting Diagnostic Matrix FILE ---------------------------------\n"
            self.tab1.textbox.setText(text)
            workBook = self.GetDiagnosticMatrixFileWorkbook()
            if self.diagnosticMatrixFileExtension == "xls":
                flag = self.TestGeneralStructureXLS(workBook, self.tab1.myTextBox10.toPlainText())
            else:
                flag = self.TestGeneralStructureXLSX_XLSM(workBook)
            if flag == True:
                self.tab1.colorTextBox6.setStyleSheet('background-color: green')
            else:
                self.tab1.colorTextBox6.setStyleSheet('background-color: red')
        else:
            self.pbvalue = self.pbvalue + 0.8772*19
            self.tab1.pbar.setValue(self.pbvalue)

#Requirements for General structure


    def Test_02043_18_04939_STRUCT_0000_XLS(self, workBook):

        sheetNames = workBook.sheet_names()
        if sheetNames[0].casefold() in ["informations générales", "general information"]:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0000_XLSX_XLSM(self, workBook):

        sheetNames = workBook.sheetnames
        if sheetNames[0].casefold() in ["informations générales", "general information"]:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0005_XLS(self, workBook):

        os.system("taskkill /f /im EXCEL.EXE")
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(workBook)
        ws = wb.Worksheets(1)
        if ws.Cells(52,2).HasFormula is False:
            return 1
            excel.Application.Quit()
        else:
            return 0
            excel.Application.Quit()

    def Test_02043_18_04939_STRUCT_0005_XLSX_XLSM(self, workBook):

        workSheet = workBook.worksheets[0]
        if "=" in str(workSheet.cell(52,2).value):
            return 0
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0010_XLS(self, workBook):
        workSheet = workBook.sheet_by_index(0)
        try:
            value = workSheet.cell_value(51, 1)
        except:
            return 0
        if isinstance(value, str) and value.strip():
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0010_XLSX_XLSM(self, workBook):

        workSheet = workBook.worksheets[0]
        if isinstance(workSheet.cell(52, 2).value, str) and workSheet.cell(52, 2).value and not workSheet.cell(52,2).value.isspace():
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0011_XLS(self, workBook):
        workSheet = workBook.sheet_by_index(0)
        try:
            value = workSheet.cell_value(51, 1)
        except:
            return 0
        if value in {"AEEV_IAEE07_0033", "02043_12_01665", "02043_12_01666"}:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0011_XLSX_XLSM(self, workBook):

        workSheet = workBook.worksheets[0]
        if workSheet.cell(52, 2).value in {"AEEV_IAEE07_0033", "02043_12_01665", "02043_12_01666"}:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0020_XLS(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "suppression" in sheetNames:
            return 1
        else:
            return 0

    def  Test_02043_18_04939_STRUCT_0020_XLSX_XLSM(self, workBook):

         sheetNames = [x.casefold() for x in workBook.sheetnames]
         if "suppression" in sheetNames:
             return 1
         else:
             return 0

    def Test_02043_18_04939_STRUCT_0025_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        try:
            index = sheetNames.index("suppression")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)
        try:
            row = workSheet.row(0)
        except:
            return 0
        for cell in row:
            if cell.value.strip().casefold() in {"sheet", "onglet"}:
                return 1
        return 0

    def Test_02043_18_04939_STRUCT_0025_XLSX_XLSM(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheetnames]
        try:
            index = sheetNames.index("suppression")
        except:
            return 0
        workSheet = workBook.worksheets[index]

        row = workSheet.iter_rows(min_col=1, min_row=1, max_row=1, max_col=1)
        for cellTuple in row:
            for cell in cellTuple:
                if str(cell.value).casefold() in {"sheet", "onglet"}:
                    return 1
        return 0

    def Test_02043_18_04939_STRUCT_0030_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        try:
            index = sheetNames.index("suppression")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)

        try:
            row = workSheet.row(0)
        except:
            return 0
        for cell in row:
            if cell.value.casefold() in {"référence de la ligne", "line number"}:
                return 1
        return 0

    def Test_02043_18_04939_STRUCT_0030_XLSX_XLSM(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheetnames]
        try:
            index = sheetNames.index("suppression")
        except:
            return 0
        workSheet = workBook.worksheets[index]

        row = workSheet.iter_rows(min_col=1, min_row=1, max_row=1, max_col=1)
        for cellTuple in row:
            for cell in cellTuple:
               if str(cell.value).casefold() in {"référence de la ligne", "line number"}:
                   return 1
        return 0

    def Test_02043_18_04939_STRUCT_0035_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        try:
            index = sheetNames.index("suppression")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)

        try:
            row = workSheet.row(0)
        except:
            return 0
        for cell in row:
            if cell.value.casefold() in {"version du tsd", "version of the document"}:
                return 1
        return 0

    def Test_02043_18_04939_STRUCT_0035_XLSX_XLSM(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheetnames]
        try:
            index = sheetNames.index("suppression")
        except:
            return 0
        workSheet = workBook.worksheets[index]

        row = workSheet.iter_rows(min_col=1, min_row=1, max_row=1, max_col=1)
        for cellTuple in row:
            for cell in cellTuple:
                if str(cell.value).casefold() in {"version du tsd", "version of the document"}:
                   return 1
        return 0

    def Test_02043_18_04939_STRUCT_0040_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        try:
            index = sheetNames.index("suppression")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)

        try:
            row = workSheet.row(0)
        except:
            return 0
        for cell in row:
            if str(cell.value).casefold() in {"justification de la modification", "change reason"}:
                return 1
        return 0

    def Test_02043_18_04939_STRUCT_0040_XLSX_XLSM(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheetnames]
        try:
            index = sheetNames.index("suppression")
        except:
            return 0
        workSheet = workBook.worksheets[index]

        row = workSheet.iter_rows(min_col=1, min_row=1, max_row=1, max_col=1)
        for cellTuple in row:
            for cell in cellTuple:
                 if str(cell.value).casefold() in {"justification de la modification", "change reason"}:
                    return 1
        return 0

    def Test_02043_18_04939_STRUCT_0051_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        try:
            index = sheetNames.index("reference docs")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)
        indexCol = workSheet.ncols
        for index in range(0,indexCol):
            column = workSheet.col(index)
            for cell in column:
                if str(cell.value).strip().casefold() == "name":
                    colNameIndex = index
                if str(cell.value).strip().casefold() == "reference":
                    colReferenceIndex = index
        colName = workSheet.col(colNameIndex)
        for index in range(0,len(colName)):
            if str(colName[index].value).strip().casefold() in [x.casefold().strip() for x in ["Vehicle Architecture Schematic", "Planche d'architecture véhicule"]]:
                if isinstance(workSheet.cell_value(index, colReferenceIndex), str) and workSheet.cell_value(index, colReferenceIndex).strip():
                    return 1
                else:
                    return 0
        return 0

    def Test_02043_18_04939_STRUCT_0051_XLSX_XLSM(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheetnames]
        try:
            index = sheetNames.index("reference docs")
        except:
            return 0
        workSheet = workBook.worksheets[index]
        rows = workSheet.iter_rows(min_row=1, max_row=5)
        for cellRowTuple in rows:
            for cell in cellRowTuple:
                if str(cell.value).strip().casefold() == "name":
                    colNameIndex = cell.column
                if str(cell.value).strip().casefold() == "reference":
                    colReferenceIndex = cell.column
        for rowIndex in range(1, workSheet.max_row):
            if str(workSheet.cell(row = rowIndex, column = colNameIndex).value).casefold().strip() in [x.casefold().strip() for x in ["Vehicle Architecture Schematic", "Planche d'architecture véhicule"]]:
                if isinstance(workSheet.cell(row = rowIndex, column = colReferenceIndex).value, str) and workSheet.cell(row = rowIndex, column = colReferenceIndex).value.strip():
                    return 1
                else:
                    return 0
        return 0

        """      colName = workSheet.iter_rows(min_col=colNameIndex, max_col=colNameIndex)
            for cellObjectTuple in colName:
                for cellObject in cellObjectTuple:
                    if str(cellObject.value).strip().casefold() == "vehicle architecture schematic" or str(cellObject.value).strip().casefold() == "planche d'architecture véhicule":
                        if isinstance(workSheet.cell(cellObject.row, colReferenceIndex).value, str) and workSheet.cell(cellObject.row, colReferenceIndex).value and not workSheet.cell(cellObject.row, colReferenceIndex).value.isspace():
                            return 1
                        else:
                            return 0
        """

    def Test_02043_18_04939_STRUCT_0052_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        try:
            index = sheetNames.index("reference docs")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)
        indexCol = workSheet.ncols
        for index in range(0, indexCol):
            column = workSheet.col(index)
            for cell in column:
                if str(cell.value).strip().casefold() == "name":
                    colNameIndex = index
                if str(cell.value).strip().casefold() == "reference":
                    colReferenceIndex = index
        colName = workSheet.col(colNameIndex)
        for index in range(0, len(colName)):
            if str(colName[index].value).strip().casefold() in [x.casefold().strip() for x in["Diagnostic Matrix", "Matrice Diag"]]:
                if isinstance(workSheet.cell_value(index, colReferenceIndex), str) and workSheet.cell_value(index, colReferenceIndex).strip():
                    return 1
                else:
                    return 0
        return 0

    def Test_02043_18_04939_STRUCT_0052_XLSX_XLSM(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheetnames]
        try:
            index = sheetNames.index("reference docs")
        except:
            return 0
        workSheet = workBook.worksheets[index]
        rows = workSheet.iter_rows(min_row=1, max_row=5)
        for cellRowTuple in rows:
            for cell in cellRowTuple:
                if str(cell.value).strip().casefold() == "name":
                    colNameIndex = cell.column
                if str(cell.value).strip().casefold() == "reference":
                    colReferenceIndex = cell.column
        for rowIndex in range(1, workSheet.max_row):
            if str(workSheet.cell(row = rowIndex, column = colNameIndex).value).casefold().strip() in [x.casefold().strip() for x in ["Diagnostic Matrix", "Matrice Diag"]]:
                if isinstance(workSheet.cell(row = rowIndex, column = colReferenceIndex).value, str) and workSheet.cell(row = rowIndex, column = colReferenceIndex).value.strip():
                    return 1
                else:
                    return 0
        return 0

    def Test_02043_18_04939_STRUCT_0053_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        try:
            index = sheetNames.index("reference docs")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)
        indexCol = workSheet.ncols
        for index in range(0, indexCol):
            column = workSheet.col(index)
            for cell in column:
                if str(cell.value).strip().casefold() == "name":
                    colNameIndex = index
                if str(cell.value).strip().casefold() == "reference":
                    colReferenceIndex = index
        colName = workSheet.col(colNameIndex)
        for index in range(0, len(colName)):
            if str(colName[index].value).strip().casefold() in [x.casefold().strip() for x in["Fault Tree", "AMDEC"]]:
                if isinstance(workSheet.cell_value(index, colReferenceIndex), str) and workSheet.cell_value(index, colReferenceIndex).strip():
                    return 1
                else:
                    return 0
        return 0

    def Test_02043_18_04939_STRUCT_0053_XLSX_XLSM(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheetnames]
        try:
            index = sheetNames.index("reference docs")
        except:
            return 0
        workSheet = workBook.worksheets[index]
        rows = workSheet.iter_rows(min_row=1, max_row=5)
        for cellRowTuple in rows:
            for cell in cellRowTuple:
                if str(cell.value).strip().casefold() == "name":
                    colNameIndex = cell.column
                if str(cell.value).strip().casefold() == "reference":
                    colReferenceIndex = cell.column
        for rowIndex in range(1, workSheet.max_row):
            if str(workSheet.cell(row = rowIndex, column = colNameIndex).value).casefold().strip() in [x.casefold().strip() for x in ["Fault Tree", "AMDEC"]]:
                if isinstance(workSheet.cell(row = rowIndex, column = colReferenceIndex).value, str) and workSheet.cell(row = rowIndex, column = colReferenceIndex).value.strip():
                    return 1
                else:
                    return 0
        return 0

    def Test_02043_18_04939_STRUCT_0054_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        try:
            index = sheetNames.index("reference docs")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)
        indexCol = workSheet.ncols
        for index in range(0, indexCol):
            column = workSheet.col(index)
            for cell in column:
                if str(cell.value).strip().casefold() == "name":
                    colNameIndex = index
                if str(cell.value).strip().casefold() == "reference":
                    colReferenceIndex = index
        colName = workSheet.col(colNameIndex)
        for index in range(0, len(colName)):
            if str(colName[index].value).strip().casefold() in [x.casefold().strip() for x in["ECU schematic", "Synoptique ECU"]]:
                if isinstance(workSheet.cell_value(index, colReferenceIndex), str) and workSheet.cell_value(index, colReferenceIndex).strip():
                    return 1
                else:
                    return 0
        return 0

    def Test_02043_18_04939_STRUCT_0054_XLSX_XLSM(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheetnames]
        try:
            index = sheetNames.index("reference docs")
        except:
            return 0
        workSheet = workBook.worksheets[index]
        rows = workSheet.iter_rows(min_row=1, max_row=5)
        for cellRowTuple in rows:
            for cell in cellRowTuple:
                if str(cell.value).strip().casefold() == "name":
                    colNameIndex = cell.column
                if str(cell.value).strip().casefold() == "reference":
                    colReferenceIndex = cell.column
        for rowIndex in range(1, workSheet.max_row):
            if str(workSheet.cell(row = rowIndex, column = colNameIndex).value).casefold().strip() in [x.casefold().strip() for x in ["ECU schematic", "Synoptique ECU"]]:
                if isinstance(workSheet.cell(row = rowIndex, column = colReferenceIndex).value, str) and workSheet.cell(row = rowIndex, column = colReferenceIndex).value.strip():
                    return 1
                else:
                    return 0
        return 0

    def Test_02043_18_04939_STRUCT_0055_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        try:
            index = sheetNames.index("reference docs")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)
        indexCol = workSheet.ncols
        for index in range(0, indexCol):
            column = workSheet.col(index)
            for cell in column:
                if str(cell.value).strip().casefold() == "name":
                    colNameIndex = index
                if str(cell.value).strip().casefold() == "reference":
                    colReferenceIndex = index
        colName = workSheet.col(colNameIndex)
        for index in range(0, len(colName)):
            if str(colName[index].value).strip().casefold() == "STD".casefold().strip():
                if isinstance(workSheet.cell_value(index, colReferenceIndex), str) and workSheet.cell_value(index, colReferenceIndex).strip():
                    return 1
                else:
                    return 0
        return 0

    def Test_02043_18_04939_STRUCT_0055_XLSX_XLSM(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheetnames]
        try:
            index = sheetNames.index("reference docs")
        except:
            return 0
        workSheet = workBook.worksheets[index]
        rows = workSheet.iter_rows(min_row=1, max_row=5)
        for cellRowTuple in rows:
            for cell in cellRowTuple:
                if str(cell.value).strip().casefold() == "name":
                    colNameIndex = cell.column
                if str(cell.value).strip().casefold() == "reference":
                    colReferenceIndex = cell.column
        for rowIndex in range(1, workSheet.max_row):
            if str(workSheet.cell(row = rowIndex, column = colNameIndex).value).casefold().strip() == "STD".casefold().strip():
                if isinstance(workSheet.cell(row = rowIndex, column = colReferenceIndex).value, str) and workSheet.cell(row = rowIndex, column = colReferenceIndex).value.strip():
                    return 1
                else:
                    return 0
        return 0

    def Test_02043_18_04939_STRUCT_0056_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        try:
            index = sheetNames.index("reference docs")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)
        indexCol = workSheet.ncols
        for index in range(0, indexCol):
            column = workSheet.col(index)
            for cell in column:
                if str(cell.value).strip().casefold() == "name":
                    colNameIndex = index
                if str(cell.value).strip().casefold() == "reference":
                    colReferenceIndex = index
        colName = workSheet.col(colNameIndex)
        for index in range(0, len(colName)):
            if str(colName[index].value).strip().casefold() == "Complexity Matrix (Decli EE)".casefold().strip():
                if isinstance(workSheet.cell_value(index, colReferenceIndex), str) and workSheet.cell_value(index, colReferenceIndex).strip():
                    return 1
                else:
                    return 0
        return 0

    def Test_02043_18_04939_STRUCT_0056_XLSX_XLSM(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheetnames]
        try:
            index = sheetNames.index("reference docs")
        except:
            return 0
        workSheet = workBook.worksheets[index]
        rows = workSheet.iter_rows(min_row=1, max_row=5)
        for cellRowTuple in rows:
            for cell in cellRowTuple:
                if str(cell.value).strip().casefold() == "name":
                    colNameIndex = cell.column
                if str(cell.value).strip().casefold() == "reference":
                    colReferenceIndex = cell.column
        for rowIndex in range(1, workSheet.max_row):
            if str(workSheet.cell(row = rowIndex, column = colNameIndex).value).casefold().strip() == "Complexity Matrix (Decli EE)".casefold().strip():
                if isinstance(workSheet.cell(row = rowIndex, column = colReferenceIndex).value, str) and workSheet.cell(row = rowIndex, column = colReferenceIndex).value.strip():
                    return 1
                else:
                    return 0
        return 0

    def Test_02043_18_04939_STRUCT_0057_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        try:
            index = sheetNames.index("reference docs")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)
        indexCol = workSheet.ncols
        for index in range(0, indexCol):
            column = workSheet.col(index)
            for cell in column:
                if str(cell.value).strip().casefold() == "name":
                    colNameIndex = index
                if str(cell.value).strip().casefold() == "reference":
                    colReferenceIndex = index
        colName = workSheet.col(colNameIndex)
        for index in range(0, len(colName)):
            if str(colName[index].value).strip().casefold() == "Décli".casefold().strip():
                if isinstance(workSheet.cell_value(index, colReferenceIndex), str) and workSheet.cell_value(index, colReferenceIndex).strip():
                    return 1
                else:
                    return 0
        return 0

    def Test_02043_18_04939_STRUCT_0057_XLSX_XLSM(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheetnames]
        try:
            index = sheetNames.index("reference docs")
        except:
            return 0
        workSheet = workBook.worksheets[index]
        rows = workSheet.iter_rows(min_row=1, max_row=5)
        for cellRowTuple in rows:
            for cell in cellRowTuple:
                if str(cell.value).strip().casefold() == "name":
                    colNameIndex = cell.column
                if str(cell.value).strip().casefold() == "reference":
                    colReferenceIndex = cell.column
        for rowIndex in range(1, workSheet.max_row):
            if str(workSheet.cell(row = rowIndex, column = colNameIndex).value).casefold().strip() == "Décli".casefold().strip():
                if isinstance(workSheet.cell(row = rowIndex, column = colReferenceIndex).value, str) and workSheet.cell(row = rowIndex, column = colReferenceIndex).value.strip():
                    return 1
                else:
                    return 0
        return 0

    def Test_02043_18_04939_STRUCT_0058_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        try:
            index = sheetNames.index("reference docs")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)
        indexCol = workSheet.ncols
        for index in range(0, indexCol):
            column = workSheet.col(index)
            for cell in column:
                if str(cell.value).strip().casefold() == "name":
                    colNameIndex = index
                if str(cell.value).strip().casefold() == "reference":
                    colReferenceIndex = index
        colName = workSheet.col(colNameIndex)
        for index in range(0, len(colName)):
            if str(colName[index].value).strip().casefold() == "DCEE".casefold().strip():
                if isinstance(workSheet.cell_value(index, colReferenceIndex), str) and workSheet.cell_value(index, colReferenceIndex).strip():
                    return 1
                else:
                    return 0
        return 0

    def Test_02043_18_04939_STRUCT_0058_XLSX_XLSM(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheetnames]
        try:
            index = sheetNames.index("reference docs")
        except:
            return 0
        workSheet = workBook.worksheets[index]
        rows = workSheet.iter_rows(min_row=1, max_row=5)
        for cellRowTuple in rows:
            for cell in cellRowTuple:
                if str(cell.value).strip().casefold() == "name":
                    colNameIndex = cell.column
                if str(cell.value).strip().casefold() == "reference":
                    colReferenceIndex = cell.column
        for rowIndex in range(1, workSheet.max_row):
            if str(workSheet.cell(row = rowIndex, column = colNameIndex).value).casefold().strip() == "DCEE".casefold().strip():
                if isinstance(workSheet.cell(row = rowIndex, column = colReferenceIndex).value, str) and workSheet.cell(row = rowIndex, column = colReferenceIndex).value.strip():
                    return 1
                else:
                    return 0
        return 0

    def Test_02043_18_04939_STRUCT_0059_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        try:
            index = sheetNames.index("reference docs")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)
        indexCol = workSheet.ncols
        for index in range(0, indexCol):
            column = workSheet.col(index)
            for cell in column:
                if str(cell.value).strip().casefold() == "name":
                    colNameIndex = index
                if str(cell.value).strip().casefold() == "reference":
                    colReferenceIndex = index
        colName = workSheet.col(colNameIndex)
        for index in range(0, len(colName)):
            if str(colName[index].value).strip().casefold() == "EEAD".casefold().strip():
                if isinstance(workSheet.cell_value(index, colReferenceIndex), str) and workSheet.cell_value(index, colReferenceIndex).strip():
                    return 1
                else:
                    return 0
        return 0

    def Test_02043_18_04939_STRUCT_0059_XLSX_XLSM(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheetnames]
        try:
            index = sheetNames.index("reference docs")
        except:
            return 0
        workSheet = workBook.worksheets[index]
        rows = workSheet.iter_rows(min_row=1, max_row=5)
        for cellRowTuple in rows:
            for cell in cellRowTuple:
                if str(cell.value).strip().casefold() == "name":
                    colNameIndex = cell.column
                if str(cell.value).strip().casefold() == "reference":
                    colReferenceIndex = cell.column
        for rowIndex in range(1, workSheet.max_row):
            if str(workSheet.cell(row = rowIndex, column = colNameIndex).value).casefold().strip() == "EEAD".casefold().strip():
                if isinstance(workSheet.cell(row = rowIndex, column = colReferenceIndex).value, str) and workSheet.cell(row = rowIndex, column = colReferenceIndex).value.strip():
                    return 1
                else:
                    return 0
        return 0

    def Test_02043_18_04939_STRUCT_0060_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        try:
            index = sheetNames.index("reference docs")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)
        indexCol = workSheet.ncols
        for index in range(0, indexCol):
            column = workSheet.col(index)
            for cell in column:
                if str(cell.value).strip().casefold() == "name":
                    colNameIndex = index
                if str(cell.value).strip().casefold() == "reference":
                    colReferenceIndex = index
        colName = workSheet.col(colNameIndex)
        for index in range(0, len(colName)):
            if str(colName[index].value).strip().casefold() == "TFD".casefold().strip():
                if isinstance(workSheet.cell_value(index, colReferenceIndex), str) and workSheet.cell_value(index, colReferenceIndex).strip():
                    return 1
                else:
                    return 0
        return 0

    def Test_02043_18_04939_STRUCT_0060_XLSX_XLSM(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheetnames]
        try:
            index = sheetNames.index("reference docs")
        except:
            return 0
        workSheet = workBook.worksheets[index]
        rows = workSheet.iter_rows(min_row=1, max_row=5)
        for cellRowTuple in rows:
            for cell in cellRowTuple:
                if str(cell.value).strip().casefold() == "name":
                    colNameIndex = cell.column
                if str(cell.value).strip().casefold() == "reference":
                    colReferenceIndex = cell.column
        for rowIndex in range(1, workSheet.max_row):
            if str(workSheet.cell(row = rowIndex, column = colNameIndex).value).casefold().strip() == "TFD".casefold().strip():
                if isinstance(workSheet.cell(row = rowIndex, column = colReferenceIndex).value, str) and workSheet.cell(row = rowIndex, column = colReferenceIndex).value.strip():
                    return 1
                else:
                    return 0
        return 0

#Requirements for [DOC4]

    def Test_02043_18_04939_STRUCT_0400_XLS(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "table" in sheetNames or "tableau" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0400_XLSX_XLSM(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheetnames]
        if "table" in sheetNames or "tableau" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0410_XLS(self, workBook):

        cellNamesRow3 = ["Version", "To diagnose", "Supplier system", "Logical flow", "Physical flow", "Client system",
                         "Type of connection", "Type", "Logical failure mode", "Physical failure mode", "Wiring harness cause",
                         "Other cause", "Operation situation / Scenario", "system effect", "Customer effect", "Comment",
                         "Feared event", "Severity", "Level", "target","Safety measure (G4) / Functional diagnostic(G3,G2,G1)",
                         "Type of failure", "Degraded mode /Safe state", "lead time", "Efficiency", "recovering mode",
                         "Requirement N° to the Design Document", "Requirement N° from Design document",
                         "research time allocated to the system (in minutes)", "HMI\n(Indicators/messages)","High level test",
                         "Diagnosis needs", "Comments"]

        cellHeaderFlow = ["Supplier system", "Logical flow", "Physical flow", "Client system", "Type of connection"]
        cellHeaderFlowPosition = []
        cellHeaderFailureModes = ["Logical failure mode", "Physical failure mode", "Wiring harness cause",
                                  "Other cause"]
        cellHeaderFailureModesPosition = []
        cellHeaderEffects = ["Operation situation / Scenario", "system effect", "Customer effect", "Comment"]
        cellHeaderEffectsPosition = []
        cellHeaderFearedEvents = ["Feared event", "Severity", "Level", "target"]
        cellHeaderFearedEventsPosition = ["Safety measure (G4) / Functional diagnostic(G3,G2,G1)",]
        cellHeaderSafetyMeasures = ["Type of failure", "Degraded mode /Safe state", "lead time", "Efficiency", "recovering mode",
                                    "Requirement N° to the Design Document", "Requirement N° from Design document"]
        cellHeaderSafetyMeasuresPosition = []
        cellHeaderDiagnosis = ["research time allocated to the system (in minutes)", "HMI\n(Indicators/messages)", "High level test",
                               "Diagnosis needs", "Comments"]
        cellHeaderDiagnosisPosition = []

# check if row 3 is OK

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("table")
        except:
            index = sheetNames.index("tableau")
        workSheet = workBook.sheet_by_index(index)
        rowsIterator = workSheet.rows(2)
        row3CellValues = list()
        for cell in rowsIterator:
            row3CellValues.append(cell.value.casefold())
        row3NumbersOfValues = len(cellNamesRow3)
        trueCases = 0
        for value in cellNamesRow3:
            if value.casefold() in row3CellValues:
                if row3CellValues.count(value.casefold()) is 1:
                    trueCases = trueCases + 1
            if "reference" in row3CellValues:
                if row3CellValues.count("reference") is 2:
                    trueCases = trueCases + 1
            if not trueCases is row3NumbersOfValues + 1:
                return 0

        for value in cellHeaderFlow:
            cellHeaderFlowPosition.append(row3CellValues.index(value.casefold()))
        for value in cellHeaderFailureModes:
            cellHeaderFailureModesPosition.append(row3CellValues.index(value.casefold()))
        for value in cellHeaderEffects:
            cellHeaderEffectsPosition.append(row3CellValues.index(value.casefold()))
        for value in cellHeaderFearedEvents:
            cellHeaderFearedEventsPosition.append(row3CellValues.index(value.casefold()))
        for value in cellHeaderSafetyMeasures:
            cellHeaderSafetyMeasuresPosition.append(row3CellValues.index(value.casefold()))
        for value in cellHeaderDiagnosis:
            cellHeaderDiagnosisPosition.append(row3CellValues.index(value.casefold()))

# sort index

        cellHeaderFlowPosition.sort()
        cellHeaderFailureModesPosition.sort()
        cellHeaderEffectsPosition.sort()
        cellHeaderFearedEventsPosition.sort()
        cellHeaderSafetyMeasuresPosition.sort()
        cellHeaderDiagnosisPosition.sort()
        tempList = []
        tempList.append(row3CellValues.index("type"))

# see if subcells of headers are together

        for listIndex in range(1, len(cellHeaderFlowPosition)):
            if cellHeaderFlowPosition[listIndex] - cellHeaderFlowPosition[listIndex - 1] > 1:
                return 0
        for listIndex in range(1, len(cellHeaderFailureModesPosition)):
            if cellHeaderFailureModesPosition[listIndex] - cellHeaderFailureModesPosition[listIndex - 1] > 1:
                return 0
        for listIndex in range(1, len(cellHeaderEffectsPosition)):
            if cellHeaderEffectsPosition[listIndex] - cellHeaderEffectsPosition[listIndex - 1] > 1:
                return 0
        for listIndex in range(1, len(cellHeaderFearedEventsPosition)):
            if cellHeaderFearedEventsPosition[listIndex] - cellHeaderFearedEventsPosition[listIndex - 1] > 1:
                return 0
        for listIndex in range(1, len(cellHeaderSafetyMeasuresPosition)):
            if cellHeaderSafetyMeasuresPosition[listIndex] - cellHeaderSafetyMeasuresPosition[listIndex - 1] > 1:
                return 0
        for listIndex in range(1, len(cellHeaderDiagnosisPosition)):
            if cellHeaderDiagnosisPosition[listIndex] - cellHeaderDiagnosisPosition[listIndex - 1] > 1:
                return 0

# get second row

        rowsIterator = workSheet.rows(1)
        row3CellValues = list()
        for row in rowsIterator:
            for cell in row:
                row3CellValues.append(cell.value.casefold())

# see if headers are OK

        for value in cellHeaderFlowPosition:
            if row3CellValues[value] != "flow":
                return 0
        for value in cellHeaderFailureModesPosition:
            if row3CellValues[value] != "failures modes":
                return 0
        for value in cellHeaderEffectsPosition:
            if row3CellValues[value] != "effects":
                return 0
        for value in cellHeaderFearedEventsPosition:
            if row3CellValues[value] != "feared events":
                return 0
        for value in cellHeaderSafetyMeasuresPosition:
            if row3CellValues[value] != "safety measure (G4) / functional diagnostic(G3,G2,G1)":
                return 0

# get first row

        rowsIterator = workSheet.rows(0)
        row3CellValues = list()
        for row in rowsIterator:
            for cell in row:
                row3CellValues.append(cell.value.casefold())

# see if main headers are OK

        cellHeaderFMEAPosition = cellHeaderFlowPosition + cellHeaderFailureModesPosition + cellHeaderEffectsPosition \
                                 + cellHeaderFearedEventsPosition + cellHeaderSafetyMeasuresPosition + tempList
        for value in cellHeaderFMEAPosition:
            if row3CellValues[value] != "fmea":
                return 0
        for value in cellHeaderDiagnosisPosition:
            if row3CellValues[value] != "diagnosis":
                return 0
        return 1

    def Test_02043_18_04939_STRUCT_0410_XLSX_XLSM(self, workBook):

        cellNamesRow3 = ["Version", "To diagnose", "Supplier system", "Logical flow", "Physical flow", "Client system",
                         "Type of connection", "Type", "Logical failure mode", "Physical failure mode", "Wiring harness cause",
                         "Other cause", "Operation situation / Scenario", "system effect", "Customer effect", "Comment",
                         "Feared event", "Severity", "Level", "target", "Safety measure (G4) / Functional diagnostic(G3,G2,G1)",
                         "Type of failure", "Degraded mode /Safe state", "lead time", "Efficiency", "recovering mode",
                         "Requirement N° to the Design Document", "Requirement N° from Design document",
                         "research time allocated to the system (in minutes)", "HMI\n(Indicators/messages)", "High level test",
                         "Diagnosis needs", "Comments"]

        cellHeaderFlow = ["Supplier system", "Logical flow", "Physical flow", "Client system", "Type of connection"]
        cellHeaderFlowPosition = []
        cellHeaderFailureModes = ["Logical failure mode", "Physical failure mode", "Wiring harness cause", "Other cause"]
        cellHeaderFailureModesPosition = []
        cellHeaderEffects = ["Operation situation / Scenario", "system effect", "Customer effect", "Comment"]
        cellHeaderEffectsPosition = []
        cellHeaderFearedEvents = ["Feared event", "Severity", "Level", "target"]
        cellHeaderFearedEventsPosition = []
        cellHeaderSafetyMeasures = ["Type of failure", "Degraded mode /Safe state", "lead time", "Efficiency", "recovering mode",
                                    "Requirement N° to the Design Document", "Requirement N° from Design document"]
        cellHeaderSafetyMeasuresPosition = []
        cellHeaderDiagnosis = ["research time allocated to the system (in minutes)", "HMI\n(Indicators/messages)", "High level test",
                               "Diagnosis needs", "Comments"]
        cellHeaderDiagnosisPosition = []

# check if row 3 is OK

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("table")
        except:
            index = sheetNames.index("tableau")
        workSheet = workBook.worksheets[index]
        rowsIterator = workSheet.iter_rows(min_row=3, max_row=3)
        row3CellValues = list()
        for row in rowsIterator:
            for cell in row:
                row3CellValues.append(cell.value.casefold())
        row3NumberOfValues = len(cellNamesRow3)
        trueCases = 0
        for value in cellNamesRow3:
            if value.casefold() in row3CellValues:
                if row3CellValues.count(value.casefold()) is 1:
                    trueCases = trueCases + 1
        if "reference" in row3CellValues:
            if row3CellValues.count("reference") is 2:
                trueCases = trueCases +1
        if not trueCases is row3NumberOfValues + 1:
            return 0

# get index of different headers in sheet

        for value in cellHeaderFlow:
            cellHeaderFlowPosition.append(row3CellValues.index(value.casefold()))
        for value in cellHeaderFailureModes:
            cellHeaderFailureModesPosition.append(row3CellValues.index(value.casefold()))
        for value in cellHeaderEffects:
            cellHeaderEffectsPosition.append(row3CellValues.index(value.casefold()))
        for value in cellHeaderFearedEvents:
            cellHeaderFearedEventsPosition.append(row3CellValues.index(value.casefold()))
        for value in cellHeaderSafetyMeasures:
            cellHeaderSafetyMeasuresPosition.append(row3CellValues.index(value.casefold()))
        for value in cellHeaderDiagnosis:
            cellHeaderDiagnosisPosition.append(row3CellValues.index(value.casefold()))

# sort index

        cellHeaderFlowPosition.sort()
        cellHeaderFailureModesPosition.sort()
        cellHeaderEffectsPosition.sort()
        cellHeaderFearedEventsPosition.sort()
        cellHeaderSafetyMeasuresPosition.sort()
        cellHeaderDiagnosisPosition.sort()
        tempList = []
        tempList.append(row3CellValues.index("type"))

# see if subcells of headers are together

        for listIndex in range(1,len(cellHeaderFlowPosition)):
            if cellHeaderFlowPosition[listIndex] - cellHeaderFlowPosition[listIndex-1] > 1:
                return 0
        for listIndex in range(1,len(cellHeaderFailureModesPosition)):
            if cellHeaderFailureModesPosition[listIndex] - cellHeaderFailureModesPosition[listIndex-1] > 1:
                return 0
        for listIndex in range(1,len(cellHeaderEffectsPosition)):
            if cellHeaderEffectsPosition[listIndex] - cellHeaderEffectsPosition[listIndex-1] > 1:
                return 0
        for listIndex in range(1,len(cellHeaderFearedEventsPosition)):
            if cellHeaderFearedEventsPosition[listIndex] - cellHeaderFearedEventsPosition[listIndex-1] > 1:
                return 0
        for listIndex in range(1,len(cellHeaderSafetyMeasuresPosition)):
            if cellHeaderSafetyMeasuresPosition[listIndex] - cellHeaderSafetyMeasuresPosition[listIndex-1] > 1:
                return 0
        for listIndex in range(1,len(cellHeaderDiagnosisPosition)):
            if cellHeaderDiagnosisPosition[listIndex] - cellHeaderDiagnosisPosition[listIndex-1] > 1:
                return 0

# get second row

        rowsIterator = workSheet.iter_rows(min_row=2, max_row=2)
        row3CellValues = list()
        for row in rowsIterator:
            for cell in row:
                row3CellValues.append(cell.value.casefold())

# see if headers are OK

        for value in cellHeaderFlowPosition:
            if row3CellValues[value] != "flow":
                return 0
        for value in cellHeaderFailureModesPosition:
            if row3CellValues[value] != "failures modes":
                return 0
        for value in cellHeaderEffectsPosition:
            if row3CellValues[value] != "effects":
                return 0
        for value in cellHeaderFearedEventsPosition:
            if row3CellValues[value] != "feared events":
                return 0
        for value in cellHeaderSafetyMeasuresPosition:
            if row3CellValues[value] != "safety measure (G4) / functional diagnostic(G3,G2,G1)":
                return 0

# get first row

        rowsIterator = workSheet.iter_rows(min_row=1, max_row=1)
        row3CellValues = list()
        for row in rowsIterator:
            for cell in row:
                row3CellValues.append(cell.value.casefold())

# see if main headers are OK

        cellHeaderFMEAPosition = cellHeaderFlowPosition + cellHeaderFailureModesPosition + cellHeaderEffectsPosition \
                + cellHeaderFearedEventsPosition + cellHeaderSafetyMeasuresPosition + tempList
        for value in cellHeaderFMEAPosition:
            if row3CellValues[value] != "fmea":
                return 0
        for value in cellHeaderDiagnosisPosition:
            if row3CellValues[value] != "diagnosis":
                return 0

        return 1

    def Test_02043_18_04939_STRUCT_0420_XLS(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "diagnostic needs" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0420_XLSX_XLSM(self, workBook):

        if "diagnostic needs" in workBook.sheetnames.casefold():
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0430_XLS(self, workBook):

        cellNamesRow1 = ["Reference", "Version", "Label", "Description", "Situation during which the diagnosis is active",
                         "Technical Effect covers by the need", "Diversity", "Allocated to the system", "Upstream requirements",
                         "Taken into account", "comment"]

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        index = sheetNames.index("diagnostic needs")
        workSheet = workBook.sheet_by_index(index)
        rowsIterator = workSheet.rows(0)
        row1CellValues = list()
        for cell in rowsIterator:
            row1CellValues.append(cell.value.casefold())
        trueCases = 0
        for value in cellNamesRow1:
            if value.casefold() in row3CellValues:
                if row1CellValues.count(value.casefold()) is 1:
                    trueCases = trueCases + 1
        if trueCases == len(cellNamesRow1):
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0430_XLSX_XLSM(self, workBook):

        cellNamesRow1 = ["Reference", "Version", "Label", "Description", "Situation during which the diagnosis is active",
                         "Technical Effect covers by the need", "Diversity", "Allocated to the system", "Upstream requirements",
                         "Taken into account", "comment"]

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        index = sheetNames.index("diagnostic needs")
        workSheet = workBook.worksheets[index]
        rowsIterator = workSheet.iter_rows(min_row=1, max_row=1)
        row1CellValues = list()
        for cell in rowsIterator:
            row1CellValues.append(cell.value.casefold())
        trueCases = 0
        for value in cellNamesRow1:
            if value.casefold() in row3CellValues:
                if row1CellValues.count(value.casefold()) is 1:
                    trueCases = trueCases + 1
        if trueCases == len(cellNamesRow1):
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0440_XLS(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "customer effects" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0440_XLSX_XLSM(self, workBook):

        if "customer effects" in workBook.sheetnames.casefold():
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0450_XLS(self, workBook):

        cellNamesRow1 = ["Taken into account", "Diagnosticability synthesis", "Comments"]

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        index = sheetNames.index("customer effects")
        workSheet = workBook.sheet_by_index(index)
        rowsIterator = workSheet.rows(0)
        row1CellValues = list()
        for cell in rowsIterator:
            row1CellValues.append(cell.value.casefold())
        trueCases = 0
        for value in cellNamesRow1:
            if value.casefold() in row3CellValues:
                if row1CellValues.count(value.casefold()) is 1:
                    trueCases = trueCases + 1
        if trueCases == len(cellNamesRow1):
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0450_XLSX_XLSM(self, workBook):

        cellNamesRow1 = ["Taken into account", "Diagnosticability synthesis", "Comments"]

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        index = sheetNames.index("customer effects")
        workSheet = workBook.worksheets[index]
        rowsIterator = workSheet.iter_rows(min_row=1, max_row=1)
        row1CellValues = list()
        for cell in rowsIterator:
            row1CellValues.append(cell.value.casefold())
        trueCases = 0
        for value in cellNamesRow1:
            if value.casefold() in row3CellValues:
                if row1CellValues.count(value.casefold()) is 1:
                    trueCases = trueCases + 1
        if trueCases == len(cellNamesRow1):
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0460_XLS(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "feared events" in sheetNames or "er" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0460_XLSX_XLSM(self, workBook):

        if "feared events" in workBook.sheetnames.casefold() or "er" in workBook.sheetnames.casefold():
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0470_XLS(self, workBook):

        cellNamesRow1 = ["Reference", "Severity", "Level", "Taken into account",
                         "Justification for not taking into account the dread Event", "Commentaire"]

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("feared events")
        except:
            index = sheetNames.index("er")
        workSheet = workBook.sheet_by_index(index)
        rowsIterator = workSheet.rows(0)
        row1CellValues = list()
        for cell in rowsIterator:
            row1CellValues.append(cell.value.casefold())
        trueCases = 0
        for value in cellNamesRow1:
            if value.casefold() in row3CellValues:
                if row1CellValues.count(value.casefold()) is 1:
                    trueCases = trueCases + 1
        if trueCases == len(cellNamesRow1):
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0470_XLSX_XLSM(self, workBook):

        cellNamesRow1 = ["Reference", "Severity", "Level", "Taken into account",
                         "Justification for not taking into account the dread Event", "Commentaire"]

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("feared events")
        except:
            index = sheetNames.index("er")
        workSheet = workBook.worksheets[index]
        rowsIterator = workSheet.iter_rows(min_row=1, max_row=1)
        row1CellValues = list()
        for cell in rowsIterator:
            row1CellValues.append(cell.value.casefold())
        trueCases = 0
        for value in cellNamesRow1:
            if value.casefold() in row3CellValues:
                if row1CellValues.count(value.casefold()) is 1:
                    trueCases = trueCases + 1
        if trueCases == len(cellNamesRow1):
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0480_XLS(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "system" in sheetNames or "système" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0480_XLSX_XLSM(self, workBook):

        if "system" in workBook.sheetnames.casefold() or "système" in workBook.sheetnames.casefold():
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0490_XLS(self, workBook):

        cellNamesRow1 = ["Description", "Taken into account"]

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("system")
        except:
            index = sheetNames.index("système")
        workSheet = workBook.sheet_by_index(index)
        rowsIterator = workSheet.rows(0)
        row1CellValues = list()
        for cell in rowsIterator:
            row1CellValues.append(cell.value.casefold())
        trueCases = 0
        for value in cellNamesRow1:
            if value.casefold() in row3CellValues:
                if row1CellValues.count(value.casefold()) is 1:
                    trueCases = trueCases + 1
        if trueCases == len(cellNamesRow1):
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0490_XLSX_XLSM(self, workBook):

        cellNamesRow1 = ["Description", "Taken into account"]

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("system")
        except:
            index = sheetNames.index("système")
        workSheet = workBook.worksheets[index]
        rowsIterator = workSheet.iter_rows(min_row=1, max_row=1)
        row1CellValues = list()
        for cell in rowsIterator:
            row1CellValues.append(cell.value.casefold())
        trueCases = 0
        for value in cellNamesRow1:
            if value.casefold() in row3CellValues:
                if row1CellValues.count(value.casefold()) is 1:
                    trueCases = trueCases + 1
        if trueCases == len(cellNamesRow1):
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0500_XLS(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "operation situation" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0500_XLSX_XLSM(self, workBook):

        if "operation situation" in workBook.sheetnames.casefold():
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0510_XLS(self, workBook):

        cellNamesRow1 = ["Description", "Taken into account", "Comments"]

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        index = sheetNames.index("operation situation")
        workSheet = workBook.sheet_by_index(index)
        rowsIterator = workSheet.rows(0)
        row1CellValues = list()
        for cell in rowsIterator:
            row1CellValues.append(cell.value.casefold())
        trueCases = 0
        for value in cellNamesRow1:
            if value.casefold() in row3CellValues:
                if row1CellValues.count(value.casefold()) is 1:
                    trueCases = trueCases + 1
        if trueCases == len(cellNamesRow1):
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0510_XLSX_XLSM(self, workBook):

        cellNamesRow1 = ["Description", "Taken into account", "Comments"]

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        index = sheetNames.index("operation situation")
        workSheet = workBook.worksheets[index]
        rowsIterator = workSheet.iter_rows(min_row=1, max_row=1)
        row1CellValues = list()
        for cell in rowsIterator:
            row1CellValues.append(cell.value.casefold())
        trueCases = 0
        for value in cellNamesRow1:
            if value.casefold() in row3CellValues:
                if row1CellValues.count(value.casefold()) is 1:
                    trueCases = trueCases + 1
        if trueCases == len(cellNamesRow1):
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0520_XLS(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "req. of tech. effects" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0520_XLSX_XLSM(self, workBook):

        if "req. of tech. effects" in workBook.sheetnames.casefold():
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0530_XLS(self, workBook):
        cellNamesRow1 = ["Reference", "version", "Description", "technical effect", "Allocated to", "Tracability with the TSD"]
        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        index = sheetNames.index("req. of tech. effects")
        workSheet = workBook.sheet_by_index(index)
        rowsIterator = workSheet.rows(0)
        row1CellValues = list()
        for cell in rowsIterator:
            row1CellValues.append(cell.value.casefold())
        trueCases = 0
        for value in cellNamesRow1:
            if value.casefold() in row3CellValues:
                if row1CellValues.count(value.casefold()) is 1:
                    trueCases = trueCases + 1
        if trueCases == len(cellNamesRow1):
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0530_XLSX_XLSM(self, workBook):
        cellNamesRow1 = ["Reference", "version", "Description", "technical effect", "Allocated to", "Tracability with the TSD"]
        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        index = sheetNames.index("req. of tech. effects")
        workSheet = workBook.worksheets[index]
        rowsIterator = workSheet.iter_rows(min_row=1, max_row=1)
        row1CellValues = list()
        for cell in rowsIterator:
            row1CellValues.append(cell.value.casefold())
        trueCases = 0
        for value in cellNamesRow1:
            if value.casefold() in row3CellValues:
                if row1CellValues.count(value.casefold()) is 1:
                    trueCases = trueCases + 1
        if trueCases == len(cellNamesRow1):
            return 1
        else:
            return 0

    # Requirements for [DOC3]

    def Test_02043_18_04939_STRUCT_0100_XLS(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "table" in sheetNames or "tableau" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0100_XLSX_XLSM(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheetnames]
        if "table" in sheetNames or "tableau" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0110_XLS(self, workBook):
        return

    def Test_02043_18_04939_STRUCT_0110_XLSX_XLSM(self, workBook):

        cellNamesRow3 = ["Référence", "Version", "Réf doc", "Variante/\noption", "Version de soft (MOTEUR / BSI,...)",
                         "sous Fonction de conception incriminée", "Groupe de constituant", "Constituant défaillant détecté",
                         "Flux fonctionnel", "Défaillance logique", "Défaillance constituant", "PPM réparties",
                         "poids", "Situation de vie client", "lien vers autre TSD", "Effet(s) client(s)",
                         "Evenement(s) redouté(s) (ER)", "Voyant(s) ou \nmessage(s)", "Code défaut",
                         "Code défauts induits", "mesures et commandes (Mesure Parametre et Test Actionneur) / Tests de cohérence",
                         "Critère de décision", "DIAGNOSTIC DEBARQUE", "Critère de decision", "Action sur constituant incriminé",
                         "Statut réunion DSP-DRD", "Action a réaliser / Commentaires", "Référence AMDEC",
                         "Justification de la modification", "Validation", "Controle Usine", "Pris en compte dans logigramme ",
                         ]

    def Test_02043_18_04939_STRUCT_0120_XLS(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "codes défauts" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0120_XLSX_XLSM(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheetnames]
        if "codes défauts" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0130_XLS(self, workBook):

        cellNamesRow2 = ["Référence", "Version", "Code défaut", "libellé (signification)", "Flux Fonctionnel", "Description de la strategie pour détecter le défaut",
                           "Seuil de détection  /  valeur  du défaut ", "Temps de confirmation du défaut",
                           "Description de la strategie de disparition du défaut / Procedure à effectuer pour vérifier la disparition du défaut",
                           "Situation de vie véhicule pour faire remonter le code défaut", "Mode dégradé", "Taux de remonté du code défaut",
                           "Voyant", "Accès scantool", "Groupe de contextes associés", "Diversité", "Applicabilité usine",
                           "condition d'applicabilité en usine", "supporté par constituant (s)", "se référer au document spécifiant DRD : (réf & version)",
                           "Référence amont", "Version de la référence amont", "Pris en compte", "Justification de la modification",
                           "Validation"]

        cellNamesRow1 = ["Liste des codes défauts", "Applicabilité projet"]

        # check if row 2 is OK

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("codes défauts")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)
        errorColValueList = list()

        # check row 2

        cellNamesRow2 = [x.strip().casefold() for x in cellNamesRow2]

        rowValueList = list()
        rowCellsList = workSheet.row(1)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())

        for value in cellNamesRow2:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        # check row 1

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsList = workSheet.row(0)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “codes défauts”, the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0130_XLSX_XLSM(self, workBook):

        cellNamesRow2 = ["Référence", "Version", "Code défaut", "libellé (signification)", "Flux Fonctionnel", "Description de la strategie pour détecter le défaut",
                         "Seuil de détection  /  valeur  du défaut ", "Temps de confirmation du défaut",
                         "Description de la strategie de disparition du défaut / Procedure à effectuer pour vérifier la disparition du défaut",
                         "Situation de vie véhicule pour faire remonter le code défaut", "Mode dégradé", "Taux de remonté du code défaut",
                         "Voyant", "Accès scantool", "Groupe de contextes associés", "Diversité", "Applicabilité usine",
                         "condition d'applicabilité en usine", "supporté par constituant (s)", "se référer au document spécifiant DRD : (réf & version)",
                          "Référence amont", "Version de la référence amont", "Pris en compte", "Justification de la modification",
                          "Validation"]

        cellNamesRow1 = ["Liste des codes défauts", "Applicabilité projet"]



        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("codes défauts")
        except:
            return 0
        workSheet = workBook.worksheets[index]

        # check if row 2 is OK

        cellNamesRow2 = [x.strip().casefold() for x in cellNamesRow2]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=2, max_row=2)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow2:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")


        #check if row1 is OK

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=1, max_row=1)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")


        if errorColValueList:
            errorColValueString = "In the sheet “codes défauts”, the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0140_XLS(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "mesures et commandes" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0140_XLSX_XLSM(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheetnames]
        if "mesures et commandes" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0150_XLS(self, workBook):

        cellNamesRow2 = ["Référence", "Version", "Type (choix par menu)", "libellé (signification)", "Description",
                         "Situation pendant laquelle la mesure ou commande est utilisable", "Statut",
                         "Taux de fiabilité du test (50%, 100%)", "Flux fonctionnel", "Uniquement /npour O Control/nlecture /nsortie effective //commande",
                         "Diversité", "Applicabilité usine", "condition d'applicabilité en usine",
                         "supporté par constituant (s)", "se référer au document spécifiant DRD : (réf & version)",
                         "Référence amont", "Version de la référence amont", "Pris en compte", "Justification de la modification", "Validation"]

        cellNamesRow1 = ["mesures et commandes", "Applicabilité projet"]

        # check if row 2 is OK

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("mesures et commandes")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)
        errorColValueList = list()

        #check row 2

        cellNamesRow2 = [x.strip().casefold() for x in cellNamesRow2]

        rowValueList = list()
        rowCellsList = workSheet.row(1)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())

        for value in cellNamesRow2:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        #check row 1

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsList = workSheet.row(0)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “mesures et commandes”, the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0150_XLSX_XLSM(self, workBook):

        cellNamesRow2 = ["Référence", "Version", "Type (choix par menu)", "libellé (signification)", "Description",
                          "Situation pendant laquelle la mesure ou commande est utilisable", "Statut",
                          "Taux de fiabilité du test (50%, 100%)", "Flux fonctionnel", "Uniquement /npour O Control/nlecture /nsortie effective //commande",
                          "Diversité", "Applicabilité usine", "condition d'applicabilité en usine",
                          "supporté par constituant (s)", "se référer au document spécifiant DRD : (réf & version)",
                          "Référence amont", "Version de la référence amont", "Pris en compte",
                          "Justification de la modification", "Validation"]

        cellNamesRow1 = ["mesures et commandes", "Applicabilité projet"]

        # check if row 2 is OK


        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("mesures et commandes")
        except:
            return 0
        workSheet = workBook.worksheets[index]

        # check if row 2 is OK

        cellNamesRow2 = [x.strip().casefold() for x in cellNamesRow2]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=2, max_row=2)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow2:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")


        #check if row1 is OK

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=1, max_row=1)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")


        if errorColValueList:
            errorColValueString = "In the sheet “mesures et commandes”, the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0160_XLS(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "diagnostic débarqués" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0160_XLSX_XLSM(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheetnames]
        if "diagnostic débarqués" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0170_XLS(self, workBook):

        cellNamesRow2 = ["Référence", "Version", "libellé (signification)", "Description", "Taux de fiabilité du test (50%, 100%)",
                          "Applicabilité Usine", "se référer au document spécifiant : (réf & version)",
                          "Pris en compte", "Justification de la modification", "Validation"]

        # check if row 2 is OK

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("diagnostic débarqués")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)

        cellNamesRow2 = [x.strip().casefold() for x in cellNamesRow2]

        rowValueList = list()
        rowCellsList = workSheet.row(1)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow2:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “diagnostic débarqués” the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0170_XLSX_XLSM(self, workBook):

        cellNamesRow1 = ["Référence", "Version", "libellé (signification)", "Description", "Taux de fiabilité du test (50%, 100%)",
                          "Applicabilité Usine", "se référer au document spécifiant : (réf & version)",
                          "Pris en compte", "Justification de la modification", "Validation"]


        # check if row 1 is OK

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("diagnostic débarqués")
        except:
            return 0
        workSheet = workBook.worksheets[index]

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=2, max_row=2)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “diagnostic débarqués”, the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0180_XLS(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "effets clients" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0180_XLSX_XLSM(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheetnames]
        if "effets clients" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0190_XLS(self, workBook):

        cellNamesRow2 = ["Noms", "Pris en compte", "Synthèse de la diagnosticabilité", "Justification de la modification"]

        cellNamesRow1= ["effets clients"]

        # check if row 2 is OK

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("effets clients")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)
        errorColValueList = list()

        # check row 2

        cellNamesRow2 = [x.strip().casefold() for x in cellNamesRow2]

        rowValueList = list()
        rowCellsList = workSheet.row(1)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())

        for value in cellNamesRow2:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        # check row 1

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsList = workSheet.row(0)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “effets clients” the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0190_XLSX_XLSM(self, workBook):

        cellNamesRow2 = ["Noms", "Pris en compte", "Synthèse de la diagnosticabilité", "Justification de la modification"]

        cellNamesRow1 = ["effets clients"]

        # check if row 2 is OK

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("effets clients")
        except:
            return 0
        workSheet = workBook.worksheets[index]

        # check if row 2 is OK

        cellNamesRow2 = [x.strip().casefold() for x in cellNamesRow2]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=2, max_row=2)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow2:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        # check if row1 is OK

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=1, max_row=1)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “effets clients”, the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0200_XLS(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "er" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0200_XLSX_XLSM(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheetnames]
        if "er" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0210_XLS(self, workBook):

        cellNamesRow2 = ["nom", "désignation", "Gravité", "Pris en compte", "Justification de non prise en compte de l'ER",
                           "Justification de la modification"]

        cellNamesRow1 = ["liste des er"]


        # check if row 2 is OK


        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("er")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)
        errorColValueList = list()

        # check row 2

        cellNamesRow2 = [x.strip().casefold() for x in cellNamesRow2]

        rowValueList = list()
        rowCellsList = workSheet.row(1)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())

        for value in cellNamesRow2:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        # check row 1

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsList = workSheet.row(0)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “er” the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0210_XLSX_XLSM(self, workBook):

        cellNamesRow2 = ["nom", "désignation", "Gravité", "Pris en compte", "Justification de non prise en compte de l'ER",
                           "Justification de la modification"]

        cellNamesRow1 = ["liste des er"]

        # check if row 2 is OK

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("er")
        except:
            return 0
        workSheet = workBook.worksheets[index]

        # check if row 2 is OK

        cellNamesRow2 = [x.strip().casefold() for x in cellNamesRow2]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=2, max_row=2)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow2:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        # check if row1 is OK

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=1, max_row=1)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “er”, the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0220_XLS(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "constituants" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0220_XLSX_XLSM(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheetnames]
        if "constituants" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0230_XLS(self, workBook):

        cellNamesRow2 = ["Noms", "Description", "Taux de défaillance (en ppm)",
                          "Découpage PSA", "Pris en compte", "Justification de la modification"]

        cellNamesRow1 = ["constituants"]

        # check if row 2 is OK


        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("constituants")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)
        errorColValueList = list()

        # check row 2

        cellNamesRow2 = [x.strip().casefold() for x in cellNamesRow2]

        rowValueList = list()
        rowCellsList = workSheet.row(1)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())

        for value in cellNamesRow2:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        # check row 1

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsList = workSheet.row(0)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “constituants” the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0230_XLSX_XLSM(self, workBook):

        cellNamesRow2 = ["Noms", "Description", "Taux de défaillance (en ppm)",
                          "Découpage PSA", "Pris en compte", "Justification de la modification"]

        cellNamesRow1 = ["constituants"]

        # check if row 2 is OK

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("constituants")
        except:
            return 0
        workSheet = workBook.worksheets[index]

        # check if row 2 is OK

        cellNamesRow2 = [x.strip().casefold() for x in cellNamesRow2]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=2, max_row=2)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow2:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        # check if row1 is OK

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=1, max_row=1)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “constituants”, the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0240_XLS(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "situations de vie" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0240_XLSX_XLSM(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheetnames]
        if "situations de vie" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0250_XLS(self, workBook):

        cellNamesRow2 = ["Situations de vie", "Justification de la modification"]

        cellNamesRow1 = ["situations de vie de la fonction ou du système:"]

        # check if row 2 is OK

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("situations de vie")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)
        errorColValueList = list()

        # check row 2

        cellNamesRow2 = [x.strip().casefold() for x in cellNamesRow2]

        rowValueList = list()
        rowCellsList = workSheet.row(1)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())

        for value in cellNamesRow2:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        # check row 1

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsList = workSheet.row(0)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “situations de vie” the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0250_XLSX_XLSM(self, workBook):

        cellNamesRow2 = ["Situations de vie", "Justification de la modification"]

        cellNamesRow1 = ["situations de vie de la fonction ou du système:"]

        # check if row 2 is OK

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("situations de vie")
        except:
            return 0
        workSheet = workBook.worksheets[index]

        # check if row 2 is OK

        cellNamesRow2 = [x.strip().casefold() for x in cellNamesRow2]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=2, max_row=2)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow2:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        # check if row1 is OK

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=1, max_row=1)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “situations de vie”, the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0260_XLS(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "liste mdd" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0260_XLSX_XLSM(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheetnames]
        if "liste mdd" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0270_XLS(self, workBook):

        cellNamesRow2 = ["Modes dégradés:", "Justification de la modification"]

        cellNamesRow1 = ["modes dégradés:"]
        cellFirstRowPosition = []

        # check if row 2 is OK

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("liste mdd")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)
        errorColValueList = list()

        # check row 2

        cellNamesRow2 = [x.strip().casefold() for x in cellNamesRow2]

        rowValueList = list()
        rowCellsList = workSheet.row(1)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())

        for value in cellNamesRow2:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        # check row 1

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsList = workSheet.row(0)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “liste mdd” the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0270_XLSX_XLSM(self, workBook):

        cellNamesRow2 = ["Modes dégradés:", "Justification de la modification"]

        cellNamesRow1 = ["modes dégradés:"]

        # check if row 2 is OK

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("liste mdd")
        except:
            return 0
        workSheet = workBook.worksheets[index]

        # check if row 2 is OK

        cellNamesRow2 = [x.strip().casefold() for x in cellNamesRow2]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=2, max_row=2)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow2:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        # check if row1 is OK

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=1, max_row=1)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “liste mdd”, the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1


# Requirements for [DOC5]

    def Test_02043_18_04939_STRUCT_0700_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "table" in sheetNames or "tableau" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0700_XLSX_XLSM(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheetnames]
        if "table" in sheetNames or "tableau" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0710_XLS(self, workBook):

        cellNamesRow3 = ["Reference", "Version", "Document of reference", "Variant/\noption",
                         "Sub-function of the system incriminated",
                         "Module / Group of parts", "Defective part", "Contribution to fonctionnality",
                         "Logical failure mode",
                         "Physical failure mode", "Weight", "Situation", "Detailed situation",
                         "Link to another DST",
                         "Technical effect", "Customer effect", "Feared events", "Degraded mode",
                         "HMI\n(Indicator lights/messages)", "Data Trouble code", "Mislead Data trouble code",
                         "Read data or I/O control", "decision criterion", "Non-embedded diagnosis",
                         "decision criterion", "Action on the incriminated part", "to do list / Comments",
                         "FMEA reference"]

        cellHeaderFailureAnalysis = ["Document of reference", "Variant/\noption",
                                     "Sub-function of the system incriminated",
                                     "Module / Group of parts", "Defective part", "Contribution to fonctionnality",
                                     "Logical failure mode",
                                     "Physical failure mode", "Weight", "Situation", "Detailed situation",
                                     "Link to another DST"]
        cellHeaderFailureAnalysisPosition = []
        cellHeaderCustomerEffects = ["Technical effect", "Customer effect", "Feared events", "Degraded mode",
                                     "HMI\n(Indicator lights/messages)"]
        cellHeaderCustomerEffectsPosition = []
        cellHeaderElementLeading = ["Data Trouble code", "Mislead Data trouble code",
                                    "Read data or I/O control", "decision criterion", "Non-embedded diagnosis"]
        cellHeaderElementLeadingPosition = []
        cellHeaderFollowUp = ["Action on the incriminated part", "to do list / Comments",
                              "FMEA reference"]
        cellHeaderFollowUpPosition = []

        cellNamesRow3 = [x.strip().casefold() for x in cellNamesRow3]
        cellHeaderFailureAnalysis = [x.strip().casefold() for x in cellHeaderFailureAnalysis]
        cellHeaderCustomerEffects = [x.strip().casefold() for x in cellHeaderCustomerEffects]
        cellHeaderElementLeading = [x.strip().casefold() for x in cellHeaderElementLeading]
        cellHeaderFollowUp = [x.strip().casefold() for x in cellHeaderFollowUp]

        # get table sheet

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("table")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)

        # check if row 3 is ok

        rowValueList = list()
        rowCellsList = workSheet.row(1)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())
       # rowValueList = [str(x).casefold().strip() for x in rowCellsList.value]
        errorColValueList = list()

        for value in cellNamesRow3:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            elif value == "decision criterion" and value in rowValueList and rowValueList.count(value) == 2:
                pass
            else:
                errorColValueList.append(value + ", ")

        rowCellsList = workSheet.row(0)
        tempList = list()
        for cell in rowCellsList:
            tempList.append(str(cell.value).casefold().strip())
        if not "Project applicability".casefold().strip() in tempList:
            errorColValueList.append("Project applicability")

        if errorColValueList:
            errorColValueString = "In the sheet “tableau” (or “tableau”), the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString

        # get index of different headers in sheet

        for value in cellHeaderFailureAnalysis:
            cellHeaderFailureAnalysisPosition.append(rowValueList.index(value))
        for value in cellHeaderCustomerEffects:
            cellHeaderCustomerEffectsPosition.append(rowValueList.index(value))
        for value in cellHeaderElementLeading:
            if value == "decision criterion":
                for index, string in enumerate(rowValueList):
                    if string == "decision criterion":
                        cellHeaderElementLeadingPosition.append(index)
            else:
                cellHeaderElementLeadingPosition.append(rowValueList.index(value))
        for value in cellHeaderFollowUp:
            cellHeaderFollowUpPosition.append(rowValueList.index(value))

        # sort index

        cellHeaderFailureAnalysisPosition.sort()
        cellHeaderCustomerEffectsPosition.sort()
        cellHeaderElementLeadingPosition.sort()
        cellHeaderFollowUpPosition.sort()

        # see if subcells of headers are together

        for listIndex in range(1, len(cellHeaderFailureAnalysisPosition)):
            if cellHeaderFailureAnalysisPosition[listIndex] - cellHeaderFailureAnalysisPosition[listIndex - 1] > 1:
                errorColValueString = "In the sheet “tableau” (or “tableau”), the column(s) below Failure Analysis: "
                errorColValueString =[errorColValueString + x + ", " for x in cellHeaderFailureAnalysis]
                errorColValueString = errorColValueString + "are not adjacent"
                return errorColValueString

        for listIndex in range(1, len(cellHeaderCustomerEffectsPosition)):
            if cellHeaderCustomerEffectsPosition[listIndex] - cellHeaderCustomerEffectsPosition[listIndex - 1] > 1:
                errorColValueString = "In the sheet “tableau” (or “tableau”), the column(s) below Customer Effects: "
                errorColValueString =[errorColValueString + x + ", " for x in cellHeaderCustomerEffects]
                errorColValueString = errorColValueString + "are not adjacent"
                return errorColValueString

        for listIndex in range(1, len(cellHeaderElementLeadingPosition)):
            if cellHeaderElementLeadingPosition[listIndex] - cellHeaderElementLeadingPosition[listIndex - 1] > 1:
                errorColValueString = "In the sheet “tableau” (or “tableau”), the column(s) below ELEMENT LEADING TO THE DEFECTIVE PART: "
                errorColValueString = [errorColValueString + x + ", " for x in cellHeaderElementLeading]
                errorColValueString = errorColValueString + "are not adjacent"
                return errorColValueString

        for listIndex in range(1, len(cellHeaderFollowUpPosition)):
            if cellHeaderFollowUpPosition[listIndex] - cellHeaderFollowUpPosition[listIndex - 1] > 1:
                errorColValueString = "In the sheet “tableau” (or “tableau”), the column(s) below FOLLOW-UP: "
                errorColValueString = [errorColValueString + x + ", " for x in cellHeaderFollowUp]
                errorColValueString = errorColValueString + "are not adjacent"
                return errorColValueString


        #check headers


        rowCellsList = workSheet.row(0)

        tempindex = cellHeaderFailureAnalysisPosition[0]
        tempcell = rowCellsList[tempindex]
        tempstring = str(tempcell.value)

        if str(rowCellsList[cellHeaderFailureAnalysisPosition[0]].value).casefold().strip() != "FAILURE ANALYSIS".casefold().strip():
            errorColValueList.append("FAILURE ANALYSIS")

        if str(rowCellsList[cellHeaderCustomerEffectsPosition[0]].value).casefold().strip() != "CUSTOMER EFFECTS".casefold().strip():
            errorColValueList.append("CUSTOMER EFFECTS")


        if str(rowCellsList[cellHeaderElementLeadingPosition[0]].value).casefold().strip() != "ELEMENT LEADING TO THE DEFECTIVE PART".casefold().strip():
            errorColValueList.append("ELEMENT LEADING TO THE DEFECTIVE PART")


        if str(rowCellsList[cellHeaderFollowUpPosition[0]].value).casefold().strip() != "FOLLOW-UP".casefold().strip():
            errorColValueList.append("FOLLOW-UP")


        if errorColValueList:
            errorColValueString = "In the sheet “tableau” (or “tableau”), the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0710_XLSX_XLSM(self, workBook):


        cellNamesRow3 = ["Reference", "Version", "Document of reference", "Variant/\noption", "Sub-function of the system incriminated",
                         "Module / Group of parts", "Defective part", "Contribution to fonctionnality", "Logical failure mode",
                         "Physical failure mode", "Weight", "Situation", "Detailed situation", "Link to another DST", "Technical effect", "Customer effect", "Feared events", "Degraded mode",
                         "HMI\n(Indicator lights/messages)", "Data Trouble code", "Mislead Data trouble code",
                         "Read data or I/O control", "decision criterion", "Non-embedded diagnosis",
                         "decision criterion", "Action on the incriminated part", "to do list / Comments", "FMEA reference"]

        cellHeaderFailureAnalysis = ["Document of reference", "Variant/\noption",
                                     "Sub-function of the system incriminated",
                                     "Module / Group of parts", "Defective part", "Contribution to fonctionnality",
                                     "Logical failure mode",
                                     "Physical failure mode", "Weight", "Situation", "Detailed situation",
                                     "Link to another DST"]
        cellHeaderFailureAnalysisPosition = []
        cellHeaderCustomerEffects = ["Technical effect", "Customer effect", "Feared events", "Degraded mode",
                                     "HMI\n(Indicator lights/messages)"]
        cellHeaderCustomerEffectsPosition = []
        cellHeaderElementLeading = ["Data Trouble code", "Mislead Data trouble code",
                                    "Read data or I/O control", "decision criterion", "Non-embedded diagnosis"]
        cellHeaderElementLeadingPosition = []
        cellHeaderFollowUp = ["Action on the incriminated part", "to do list / Comments",
                              "FMEA reference"]
        cellHeaderFollowUpPosition = []

        cellNamesRow3 = [x.strip().casefold() for x in cellNamesRow3]
        cellHeaderFailureAnalysis = [x.strip().casefold() for x in cellHeaderFailureAnalysis]
        cellHeaderCustomerEffects = [x.strip().casefold() for x in cellHeaderCustomerEffects]
        cellHeaderElementLeading = [x.strip().casefold() for x in cellHeaderElementLeading]
        cellHeaderFollowUp = [x.strip().casefold() for x in cellHeaderFollowUp]

        # get table sheet

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("table")
        except:
            return 0
        workSheet = workBook.worksheets[index]

        # check if row 3 is ok

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=2, max_row=2)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow3:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            elif value == "decision criterion" and value in rowValueList and rowValueList.count(value) == 2:
                pass
            else:
                errorColValueList.append(value + ", ")

        rowCellsGenrator = workSheet.iter_rows(min_row=1, max_row=1)
        tempList = list()
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                tempList.append(str(cell.value).casefold().strip())
        if not "Project applicability".casefold().strip() in tempList:
            errorColValueList.append("Project applicability")

        if errorColValueList:
            errorColValueString = "In the sheet “tableau” (or “tableau”), the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString

        # get index of different headers in sheet

        for value in cellHeaderFailureAnalysis:
            cellHeaderFailureAnalysisPosition.append(rowValueList.index(value))
        for value in cellHeaderCustomerEffects:
            cellHeaderCustomerEffectsPosition.append(rowValueList.index(value))
        for value in cellHeaderElementLeading:
            if value == "decision criterion":
                for index, string in enumerate(rowValueList):
                    if string == "decision criterion":
                        cellHeaderElementLeadingPosition.append(index)
            else:
                cellHeaderElementLeadingPosition.append(rowValueList.index(value))
        for value in cellHeaderFollowUp:
            cellHeaderFollowUpPosition.append(rowValueList.index(value))

        # sort index

        cellHeaderFailureAnalysisPosition.sort()
        cellHeaderCustomerEffectsPosition.sort()
        cellHeaderElementLeadingPosition.sort()
        cellHeaderFollowUpPosition.sort()

        # see if subcells of headers are together

        for listIndex in range(1, len(cellHeaderFailureAnalysisPosition)):
            if cellHeaderFailureAnalysisPosition[listIndex] - cellHeaderFailureAnalysisPosition[listIndex - 1] > 1:
                errorColValueString = "In the sheet “tableau” (or “tableau”), the column(s) below Failure Analysis: "
                errorColValueString = [errorColValueString + x + ", " for x in cellHeaderFailureAnalysis]
                errorColValueString = errorColValueString + "are not adjacent"
                return errorColValueString

        for listIndex in range(1, len(cellHeaderCustomerEffectsPosition)):
            if cellHeaderCustomerEffectsPosition[listIndex] - cellHeaderCustomerEffectsPosition[listIndex - 1] > 1:
                errorColValueString = "In the sheet “tableau” (or “tableau”), the column(s) below Customer Effects: "
                errorColValueString = [errorColValueString + x + ", " for x in cellHeaderCustomerEffects]
                errorColValueString = errorColValueString + "are not adjacent"
                return errorColValueString

        for listIndex in range(1, len(cellHeaderElementLeadingPosition)):
            if cellHeaderElementLeadingPosition[listIndex] - cellHeaderElementLeadingPosition[listIndex - 1] > 1:
                errorColValueString = "In the sheet “tableau” (or “tableau”), the column(s) below ELEMENT LEADING TO THE DEFECTIVE PART: "
                errorColValueString = [errorColValueString + x + ", " for x in cellHeaderElementLeading]
                errorColValueString = errorColValueString + "are not adjacent"
                return errorColValueString

        for listIndex in range(1, len(cellHeaderFollowUpPosition)):
            if cellHeaderFollowUpPosition[listIndex] - cellHeaderFollowUpPosition[listIndex - 1] > 1:
                errorColValueString = "In the sheet “tableau” (or “tableau”), the column(s) below FOLLOW-UP: "
                errorColValueString = [errorColValueString + x + ", " for x in cellHeaderFollowUp]
                errorColValueString = errorColValueString + "are not adjacent"
                return errorColValueString

        # check headers


        rowCellsGenrator = workSheet.iter_rows(min_row=1, max_row=1)
        rowValueList = list()
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())

        if str(rowValueList[cellHeaderFailureAnalysisPosition[0]]).casefold().strip() != "FAILURE ANALYSIS".casefold().strip():
            errorColValueList.append("FAILURE ANALYSIS")

        if str(rowValueList[cellHeaderCustomerEffectsPosition[0]]).casefold().strip() != "CUSTOMER EFFECTS".casefold().strip():
            errorColValueList.append("CUSTOMER EFFECTS")

        if str(rowValueList[cellHeaderElementLeadingPosition[0]]).casefold().strip() != "ELEMENT LEADING TO THE DEFECTIVE PART".casefold().strip():
            errorColValueList.append("ELEMENT LEADING TO THE DEFECTIVE PART")

        if str(rowValueList[cellHeaderFollowUpPosition[0]]).casefold().strip() != "FOLLOW-UP".casefold().strip():
            errorColValueList.append("FOLLOW-UP")

        if errorColValueList:
            errorColValueString = "In the sheet “tableau” (or “tableau”), the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0720_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "data trouble codes" in sheetNames or "codes défauts" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0720_XLSX_XLSM(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheetnames]
        if "data trouble codes" in sheetNames or "codes défauts" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0730_XLS(self, workBook):

        cellNamesRow1 = ["Reference", "Version", "Data trouble code", "Label", "Description of the qualification conditions",
                           "Detection threshold", "Qualification time","Description of the dequalification conditions / Operation to do to check if the defect disappeared",
                           "Conditions of the diagnostic activation", "Degraded mode", "Failure detection rate",
                           "Indicateur light", "Visibility of the failure with the Scantool", "Freeze Frame Class",
                            "Diversity", "Stored by the ECU", "Upstream requirements", "Taken into account", "B78", "DV", "projet X", "Projet Y"]

        # check if row 1 is OK

        # get table sheet

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("data trouble codes")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsList = workSheet.row(0)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “data trouble codes” (or “codes défauts”), the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0730_XLSX_XLSM(self, workBook):

        cellNamesRow1 = ["Reference", "Version", "Data trouble code", "Label", "Description of the qualification conditions",
                           "Detection threshold", "Qualification time",
                           "Description of the dequalification conditions / Operation to do to check if the defect disappeared",
                           "Conditions of the diagnostic activation", "Degraded mode", "Failure detection rate",
                           "Indicateur light", "Visibility of the failure with the Scantool", "Freeze Frame Class",
                            "Diversity", "Stored by the ECU", "Upstream requirements", "Taken into account",
                           "B78", "DV", "projet X", "Projet Y"]

        # check if row 1 is OK

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("data trouble codes")
        except:
            return 0
        workSheet = workBook.worksheets[index]

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=1, max_row=1)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()


        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")


        if errorColValueList:
            errorColValueString = "In the sheet “data trouble codes” (or “codes défauts”), the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0740_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "read data and io control" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0740_XLSX_XLSM(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheetnames]
        if "read data and io control" in sheetNames :
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0750_XLS(self, workBook):

        cellNamesRow1 = ["Reference", "Version", "Type of diagnosis", "Label", "Description", "Conditions of the diagnostic activation",
                           "Status", "Diversity", "Stored by the ECU", "Upstream requirements", "Taken into account",
                           "B78", "DV", "projet X", "Projet Y"]

        # check if row 1 is OK

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("read data and io control")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsList = workSheet.row(0)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “Read data and IO control” (or “mesures et commandes”), the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0750_XLSX_XLSM(self, workBook):

        cellNamesRow1 = ["Reference", "Version", "Type of diagnosis", "Label", "Description", "Conditions of the diagnostic activation",
                           "Status", "Diversity", "Stored by the ECU", "Upstream requirements", "Taken into account",
                           "B78", "DV", "projet X", "Projet Y"]

        # check if row 1 is OK

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("read data and io control")
        except:
            return 0
        workSheet = workBook.worksheets[index]

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=1, max_row=1)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()


        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")


        if errorColValueList:
            errorColValueString = "In the sheet “Read data and IO control” (or “mesures et commandes”), the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0760_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "not embedded diagnosis" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0760_XLSX_XLSM(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheetnames]
        if "not embedded diagnosis" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0770_XLS(self, workBook):

        cellNamesRow1 = ["Reference", "Version", "Label", "Description", "Upstream requirements",
                           "Taken into account", "B78", "DV", "projet X", "Projet Y"]

        # check if row 1 is OK

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("not embedded diagnosis")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsList = workSheet.row(0)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “Not embedded diagnosis“, the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0770_XLSX_XLSM(self, workBook):

        cellNamesRow1 = ["Reference", "Version", "Label", "Description", "Upstream requirements",
                           "Taken into account", "B78", "DV", "projet X", "Projet Y"]

        # check if row 1 is OK

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("not embedded diagnosis")
        except:
            return 0
        workSheet = workBook.worksheets[index]

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=1, max_row=1)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “Not embedded diagnosis“, the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0780_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "not embedded diagnosis" in sheetNames or "effets clients" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0780_XLSX_XLSM(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheetnames]
        if "not embedded diagnosis" in sheetNames or "effets clients" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0790_XLS(self, workBook):

        cellNamesRow1 = ["Name", "Taken into account", "Diagnosticability synthesis"]

        # check if row 1 is OK

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("customer effect")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsList = workSheet.row(0)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “Customer effect” (or “Effets clients”), the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0790_XLSX_XLSM(self, workBook):

        cellNamesRow1 = ["Name", "Taken into account", "Diagnosticability synthesis"]

        # check if row 1 is OK

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("customer effect")
        except:
            return 0
        workSheet = workBook.worksheets[index]

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=1, max_row=1)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “Customer effect” (or “Effets clients”), the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0800_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "feared events" in sheetNames or "er" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0800_XLSX_XLSM(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheetnames]
        if "feared events" in sheetNames or "er" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0810_XLS(self, workBook):

        cellNamesRow1 = ["Description", "Description", "Severity", "Taken into account",
                           "Justification for not taking into account the dread Event"]

        # check if row 1 is OK

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("feared events")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsList = workSheet.row(0)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “Feared events” (or “ER”), the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0810_XLSX_XLSM(self, workBook):

        cellNamesRow1 = ["Description", "Description", "Severity", "Taken into account",
                           "Justification for not taking into account the dread Event"]

        # check if row 1 is OK

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("feared events")
        except:
            return 0
        workSheet = workBook.worksheets[index]
        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=1, max_row=1)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “Feared events” (or “ER”), the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0820_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "parts" in sheetNames or "constituants" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0820_XLSX_XLSM(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheetnames]
        if "parts" in sheetNames or "constituants" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0830_XLS(self, workBook):

        cellNamesRow1 = ["Name", "Description", "Taken into account"]

        # check if row 1 is OK

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("parts")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsList = workSheet.row(0)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “Parts” (or “Constituants”), the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0830_XLSX_XLSM(self, workBook):

        cellNamesRow1 = ["Name", "Description", "Taken into account"]

        # check if row 1 is OK

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("parts")
        except:
            return 0
        workSheet = workBook.worksheets[index]

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=1, max_row=1)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “Parts” (or “Constituants”), the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0840_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "situation" in sheetNames or "situations de vie" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0840_XLSX_XLSM(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheetnames]
        if "situation" in sheetNames or "situations de vie" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0850_XLS(self, workBook):

        cellNamesRow1 = ["Description", "Taken into account", "Comments"]

        # check if row 1 is OK

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("situation")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsList = workSheet.row(0)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “Situation” (or “situations de vie”), the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0850_XLSX_XLSM(self, workBook):

        cellNamesRow1 = ["Description", "Taken into account", "Comments"]

        # check if row 1 is OK

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("situation")
        except:
            return 0
        workSheet = workBook.worksheets[index]

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=1, max_row=1)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “Situation” (or “situations de vie”), the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0860_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "degraded mode" in sheetNames or "liste mdd" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0860_XLSX_XLSM(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheetnames]
        if "degraded mode" in sheetNames or "liste mdd" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0870_XLS(self, workBook):

        cellNamesRow1 = ["Modes dégradés:", "Taken into account"]

        # check if row 1 is OK

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("degraded mode")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsList = workSheet.row(0)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “Degraded mode” (or “Liste MDD”), the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0870_XLSX_XLSM(self, workBook):

        cellNamesRow1 = ["Modes dégradés:", "Taken into account"]

        # check if row 1 is OK

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("degraded mode")
        except:
            return 0
        workSheet = workBook.worksheets[index]

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=1, max_row=1)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “Degraded mode” (or “Liste MDD”), the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0880_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "technical effect" in sheetNames or "effets techniques" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0880_XLSX_XLSM(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheetnames]
        if "technical effect" in sheetNames or "effets techniques" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0890_XLS(self, workBook):

        cellNamesRow1 = ["Name", "Taken into account", "Upstream requirements"]

        # check if row 1 is OK

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("technical effect")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsList = workSheet.row(0)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “Technical effect “ (or “Effets techniques”), the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0890_XLSX_XLSM(self, workBook):

        cellNamesRow1 = ["Name", "Taken into account", "Upstream requirements"]

        # check if row 1 is OK

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("technical effect")
        except:
            return 0
        workSheet = workBook.worksheets[index]

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=1, max_row=1)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “Technical effect “ (or “Effets techniques”), the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0900_XLS(self, workBook):
        sheetNames = [x.casefold() for x in workBook.sheet_names()]
        if "variant" in sheetNames or "variantes" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0900_XLSX_XLSM(self, workBook):

        sheetNames = [x.casefold() for x in workBook.sheetnames]
        if "variant" in sheetNames or "variantes" in sheetNames:
            return 1
        else:
            return 0

    def Test_02043_18_04939_STRUCT_0910_XLS(self, workBook):

        cellNamesRow1 = ["Name", "Description", "Taken into account"]

        # check if row 1 is OK

        sheetNames = workBook.sheet_names()
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("variant")
        except:
            return 0
        workSheet = workBook.sheet_by_index(index)

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsList = workSheet.row(0)
        for cell in rowCellsList:
            rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “Variant “ (or “Variantes”), the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1

    def Test_02043_18_04939_STRUCT_0910_XLSX_XLSM(self, workBook):

        cellNamesRow1 = ["Name", "Description", "Taken into account"]

        # check if row 1 is OK

        sheetNames = workBook.sheetnames
        sheetNames = [x.casefold() for x in sheetNames]
        try:
            index = sheetNames.index("variant")
        except:
            return 0
        workSheet = workBook.worksheets[index]

        cellNamesRow1 = [x.strip().casefold() for x in cellNamesRow1]

        rowValueList = list()
        rowCellsGenrator = workSheet.iter_rows(min_row=1, max_row=1)
        for cellTuple in rowCellsGenrator:
            for cell in cellTuple:
                rowValueList.append(str(cell.value).casefold().strip())
        errorColValueList = list()

        for value in cellNamesRow1:
            if value in rowValueList and rowValueList.count(value) == 1:
                pass
            else:
                errorColValueList.append(value + ", ")

        if errorColValueList:
            errorColValueString = "In the sheet “Variant “ (or “Variantes”), the column(s) "
            errorColValueString = errorColValueString + "".join(errorColValueList)
            errorColValueString = errorColValueString + " is(are) not present or not written correctly as in the document [DOC3]."
            return errorColValueString
        else:
            return 1



    def buttonClicked(self):

        self.tab1.textbox.setText("")
        self.tab1.pbar.setValue(0)
        self.pbvalue = 0
        if not self.tab2.myTextBox5.toPlainText():
            self.download_file(self.CesareLink)
        if not self.tab2.myTextBox4.toPlainText():
            self.download_file(self.TSDConfigLink)
        if not self.tab2.myTextBox6.toPlainText():
            self.download_file(self.CustomerEffectLink)
        if not self.tab2.myTextBox9.toPlainText():
            self.download_file(self.DiversityLink)

        self.TestTsdFile()
        self.TestTsdVehicleFunctionFile()
        self.TestTsdSystemFile()
        self.TestAmdecFile()
        self.TestExportMedialecMatriceFile()
        self.TestDiagnosticMatrixFile()

if __name__ == '__main__':

    try:
        FindWindow(None, "TSD Checker  V0.2")
        windll.user32.MessageBoxW(0, "Application already running", "Warning", 0|48)

    except:
        app = QApplication(sys.argv)
        apel = Test()
        myQLabel = QLabel()
        sys.exit(app.exec_())



