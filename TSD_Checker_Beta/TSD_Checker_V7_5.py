from PyQt5.QtWidgets import QWidget, QPushButton, QApplication, QComboBox, QLabel, QLineEdit,  QTabWidget, QVBoxLayout, QProgressBar, QRadioButton
from PyQt5 import QtCore, QtWidgets
import win32com.client as win32
import os
import requests
from ctypes import windll
import OptionalFilesParser
import GeneralStructureTester
import FileMeasure
from lxml import etree, objectify
import ExcelEdit
import WholenessTester
import Coherence_checksTester
import IndicatorTester
import time
import xlrd
import json
import zipfile
import shutil
import win32api
import getpass
import io
import sys
import xlwt

appName = "TSD Checker V7.5"
pBarIncrement = 100/174

class Application(QWidget):

    def __init__(self):
        super().__init__()
        self.left = 200
        self.top = 200
        self.width = 1160
        self.height = 570
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
        self.criticity_blocking_passed = 0
        self.criticity_warning_passed = 0
        self.criticity_information_passed = 0
        self.start_time = 0
        self.end_time = 0
        self.opening_time = 0


        self.list_element = []

        self.flag_load_configuration = False
        self.DOC3Exists = False
        self.DOC4Exists = False
        self.DOC5Exists = False

        self.version_cesare_file = ""
        self.version_criticity_file = ""
        self.version_cutomer_effect = ""
        self.version_diversity_file = ""

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

    def setIntranet(self):
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

    def setInternet(self):
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

    def openFileNameDialog1(self):
        fileName1, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab1, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab1.myTextBox1.setText(fileName1)

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
        fileName7, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab2, 'Open File4', QtCore.QDir.rootPath(), '*.*')
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
        tab.lbl_coverage.move(750, 12)
        tab.message = ""
        tab.textbox_coverage = QtWidgets.QTextEdit(self.tab1)
        tab.textbox_coverage.setText(tab.message)
        tab.textbox_coverage.move(850, 10)
        tab.textbox_coverage.resize(70, 20)
        tab.textbox_coverage.setReadOnly(True)

        # Create convergence textbox
        tab.lbl_convergence = QLabel("Convergence Indicator:", tab)
        tab.lbl_convergence.move(930, 12)
        tab.message = ""
        tab.textbox_convergence = QtWidgets.QTextEdit(self.tab1)
        tab.textbox_convergence.setText(tab.message)
        tab.textbox_convergence.move(1050, 10)
        tab.textbox_convergence.resize(70, 20)
        tab.textbox_convergence.setReadOnly(True)
        # Create coverage textbox

        tab.lb2_coverage = QLabel("Coverage Indicator:", tab)
        tab.lb2_coverage.move(750, 42)
        tab.message = ""
        tab.textbox_coverage_1 = QtWidgets.QTextEdit(self.tab1)
        tab.textbox_coverage_1.setText(tab.message)
        tab.textbox_coverage_1.move(850, 40)
        tab.textbox_coverage_1.resize(70, 20)
        tab.textbox_coverage_1.setReadOnly(True)

        # Create convergence textbox
        tab.lb2_convergence = QLabel("Convergence Indicator:", tab)
        tab.lb2_convergence.move(930, 42)
        tab.message = ""
        tab.textbox_convergence_1 = QtWidgets.QTextEdit(self.tab1)
        tab.textbox_convergence_1.setText(tab.message)
        tab.textbox_convergence_1.move(1050, 40)
        tab.textbox_convergence_1.resize(70, 20)
        tab.textbox_convergence_1.setReadOnly(True)

        tab.lb2_coverage = QLabel("Coverage Indicator:", tab)
        tab.lb2_coverage.move(750, 72)
        tab.message = ""
        tab.textbox_coverage_2 = QtWidgets.QTextEdit(self.tab1)
        tab.textbox_coverage_2.setText(tab.message)
        tab.textbox_coverage_2.move(850, 72)
        tab.textbox_coverage_2.resize(70, 20)
        tab.textbox_coverage_2.setReadOnly(True)

        # Create convergence textbox
        tab.lb3_convergence = QLabel("Convergence Indicator:", tab)
        tab.lb3_convergence.move(930, 72)
        tab.message = ""
        tab.textbox_convergence_2 = QtWidgets.QTextEdit(self.tab1)
        tab.textbox_convergence_2.setText(tab.message)
        tab.textbox_convergence_2.move(1050, 70)
        tab.textbox_convergence_2.resize(70, 20)
        tab.textbox_convergence_2.setReadOnly(True)

        # Create a textbox
        tab.message = ""
        tab.textbox = QtWidgets.QTextEdit(self.tab1)
        tab.textbox.setText(tab.message)
        tab.textbox.move(10, 330)
        tab.textbox.resize(1110, 130)
        tab.textbox.setReadOnly(True)

        # create a progress bar
        tab.pbar = QProgressBar(self.tab1)
        tab.pbar.setGeometry(10, 310, 1110, 20)
        tab.pbar.setAlignment(QtCore.Qt.AlignCenter)
        tab.pbar.setValue(0)
        tab.pbar.move(10, 460)

        # Create a color textbox1
        tab.colorTextBox1 = QtWidgets.QTextEdit(self.tab1)
        tab.colorTextBox1.setStyleSheet(" background-color: grey ")
        tab.colorTextBox1.resize(20, 20)
        tab.colorTextBox1.move(725, 10)

        # Create a color textbox2
        tab.colorTextBox2 = QtWidgets.QTextEdit(self.tab1)
        tab.colorTextBox2.setStyleSheet(" background-color: grey ")
        tab.colorTextBox2.resize(20, 20)
        tab.colorTextBox2.move(725, 40)

        # Create a color textbox3
        tab.colorTextBox3 = QtWidgets.QTextEdit(self.tab1)
        tab.colorTextBox3.setStyleSheet(" background-color: grey ")
        tab.colorTextBox3.resize(20, 20)
        tab.colorTextBox3.move(725, 70)

        # Create a drop down list
        tab.lbl = QLabel("Check level", tab)
        tab.combo = QComboBox(tab)
        tab.combo.addItem("Final")
        tab.combo.addItem("Previsional")
        tab.combo.addItem("Consolidated")
        tab.combo.resize(508, 20.4)  # rezise the drop down list
        tab.combo.move(215, 200)
        tab.lbl.move(5, 205)
        tab.combo.activated[str].connect(self.onActivated)

        # Create a drop down list
        tab.lbl3 = QLabel("Diversity management", tab)
        tab.combo3 = QComboBox(tab)
        tab.combo3.addItem("Codes LCDV")
        tab.combo3.addItem("Codes EC")
        tab.combo3.resize(508, 20.4)  # rezise the drop down list
        tab.combo3.move(215, 260)
        tab.lbl3.move(5, 265)
        tab.combo3.activated[str].connect(self.onActivated)

        # Create a drop down list
        tab.lbl1 = QLabel("Project name", tab)
        tab.combo1 = QComboBox(tab)
        tab.combo1.addItem("Generic")
        tab.combo1.addItem("All")
        tab.combo1.resize(330, 20.4)  # rezise the drop down list
        tab.combo1.move(215, 290)
        tab.lbl1.move(5, 295)
        tab.combo1.activated[str].connect(self.onActivated)

        # Create a dropdown list
        tab.lbl2 = QLabel("Architecture type", tab)
        tab.combo2 = QComboBox(tab)
        tab.combo2.addItem("Archi 2010")
        tab.combo2.addItem("Archi NEA R1")
        tab.combo2.addItem("Archi NEA R2")
        tab.combo2.resize(508, 20.4)
        tab.combo2.move(215, 230)
        tab.lbl2.move(5, 235)
        tab.combo2.activated[str].connect(self.onActivated)

        self.setGeometry(self.left, self.top, self.width, self.height)
        self.setWindowTitle('TSD Checker')

        tab.importNames = QPushButton(tab)
        tab.importNames.setText("Import Project names")
        tab.importNames.resize(160, 20.4)
        tab.importNames.move(565, 290)
        tab.importNames.clicked.connect(self.ImportProjectNames)

        tab.save_config = QPushButton(tab)
        tab.save_config.setText("Save \nconfiguration")
        tab.save_config.resize(120,40)
        tab.save_config.move(870, 270)
        tab.save_config.setEnabled(False)
        tab.save_config.clicked.connect(self.ButtonSaveConfigClick)

        tab.load_config = QPushButton(tab)
        tab.load_config.setText("Load \nconfiguration")
        tab.load_config.resize(120, 40)
        tab.load_config.move(1000, 270)
        tab.load_config.clicked.connect(self.ButtonLoadConfigClick)

        # File Selection Dialog1
        tab.lbl2 = QLabel("FSE TSD File:", tab)
        tab.lbl2.move(5, 15)
        tab.myTextBox1 = QtWidgets.QTextEdit(tab)
        tab.myTextBox1.resize(460, 25)
        tab.myTextBox1.move(215, 10)
        tab.myTextBox1.setReadOnly(True)
        tab.myTextBox1.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        tab.button1 = QPushButton('...', tab)
        tab.button1.clicked.connect(self.openFileNameDialog1)
        tab.button1.move(675, 10)
        tab.button1.resize(45, 22)

        # File Selection Dialog2
        tab.lbl3 = QLabel("TSD vehicle Function file:", tab)
        tab.lbl3.move(5, 45)
        tab.myTextBox2 = QtWidgets.QTextEdit(tab)
        tab.myTextBox2.resize(460, 25)
        tab.myTextBox2.move(215, 40)
        tab.myTextBox2.setReadOnly(True)
        tab.myTextBox2.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        tab.button2 = QPushButton('...', tab)
        tab.button2.clicked.connect(self.openFileNameDialog2)
        tab.button2.move(675, 40)
        tab.button2.resize(45, 22)

        # File Selection Dialog3
        tab.lbl4 = QLabel("TSD system file:", tab)
        tab.lbl4.move(5, 75)
        tab.myTextBox3 = QtWidgets.QTextEdit(tab)
        tab.myTextBox3.resize(460, 25)
        tab.myTextBox3.move(215, 70)
        tab.myTextBox3.setReadOnly(True)
        tab.myTextBox3.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        tab.button3 = QPushButton('...', tab)
        tab.button3.clicked.connect(self.openFileNameDialog3)
        tab.button3.move(675, 70)
        tab.button3.resize(45, 22)

        # File Selection Dialog4
        tab.lbl8 = QLabel("AMDEC:", tab)
        tab.lbl8.move(5, 105)
        tab.myTextBox4 = QtWidgets.QTextEdit(tab)
        tab.myTextBox4.resize(460, 25)
        tab.myTextBox4.move(215, 100)
        tab.myTextBox4.setReadOnly(True)
        tab.myTextBox4.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        tab.button4 = QPushButton('...', tab)
        tab.button4.clicked.connect(self.openFileNameDialog7)
        tab.button4.move(675, 100)
        tab.button4.resize(45, 22)

        # File Selection Dialog5
        tab.lbl9 = QLabel("Diagnostic matrix (Export MedialecMatrice):", tab)
        tab.lbl9.move(5, 135)
        tab.myTextBox5 = QtWidgets.QTextEdit(tab)
        tab.myTextBox5.resize(460, 25)
        tab.myTextBox5.move(215, 130)
        tab.myTextBox5.setReadOnly(True)
        tab.button5 = QPushButton('...', tab)
        tab.button5.clicked.connect(self.openFileNameDialog8)
        tab.button5.move(675, 130)
        tab.button5.resize(45, 22)
        tab.myTextBox5.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)

        # File Selection Dialog6
        tab.lbl11 = QLabel("Diagnostic messagery (odx):", tab)
        tab.lbl11.move(5, 165)
        tab.myTextBox6 = QtWidgets.QTextEdit(tab)
        tab.myTextBox6.resize(460, 25)
        tab.myTextBox6.move(215, 160)
        tab.myTextBox6.setReadOnly(True)
        tab.myTextBox6.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        tab.button6 = QPushButton('...', tab)
        tab.button6.clicked.connect(self.openFileNameDialog20)
        tab.button6.move(675, 160)
        tab.button6.resize(45, 22)

        tab.lbl61 = QLabel("SubFamilly :", tab)
        tab.lbl61.move(725, 164)
        tab.myTextBox61 = QtWidgets.QTextEdit(tab)
        tab.myTextBox61.resize(90, 25)
        tab.myTextBox61.move(785, 160)
        # tab.myTextBox61.setReadOnly(True)
        tab.myTextBox61.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)

        # Check button
        tab.button = QPushButton('Check', tab)
        tab.button.move(512, 490)
        tab.button.resize(90, 25)
        tab.button.clicked.connect(self.buttonClicked)
        #button.setStyleSheet('QPushButton {background-color: white; color: black;}')
        tab.buttonNew = QPushButton("Open \nReport", tab)
        tab.buttonNew.resize(120, 40)
        tab.buttonNew.move(740, 270)
        tab.buttonNew.setEnabled(False)
        tab.buttonNew.clicked.connect(self.ButtonReportClick)

        self.show()

    def ButtonReportClick(self):

        self.excel = win32.gencache.EnsureDispatch('Excel.Application')

        if self.DOC3Exists:
           fileName = self.tab1.myTextBox1.toPlainText()
           self.excel.Visible = True
           self.excel.Workbooks.Open(fileName)

        if self.DOC4Exists:
           fileName = self.tab1.myTextBox2.toPlainText()
           self.excel.Visible = True
           self.excel.Workbooks.Open(fileName)

        if self.DOC5Exists:
           fileName = self.tab1.myTextBox3.toPlainText()
           self.excel.Visible = True
           self.excel.Workbooks.Open(fileName)

    def ImportProjectNames(self):
        if self.tab1.myTextBox1.toPlainText() != "":
            DOCPath = self.tab1.myTextBox1.toPlainText()
            projectsRow = 4
        elif self.tab1.myTextBox1.toPlainText() == "" and self.tab1.myTextBox2.toPlainText() != "":
            DOCPath = self.tab1.myTextBox2.toPlainText()
            projectsRow = 2
        elif self.tab1.myTextBox1.toPlainText() == "" and self.tab1.myTextBox2.toPlainText() == "" and self.tab1.myTextBox3.toPlainText() != "":
            DOCPath = self.tab1.myTextBox3.toPlainText()
            projectsRow = 2
        try:
            extension = DOCPath.split(".")[-1]
            if extension == "xls":
                DOCWorkbook = xlrd.open_workbook(DOCPath, formatting_info=True)
            else:
                DOCWorkbook = xlrd.open_workbook(DOCPath)

            try:
                workSheet = DOCWorkbook.sheet_by_name("tableau")
            except:
                workSheet = DOCWorkbook.sheet_by_name("Table")

            flag = False
            projects = []
            for index1 in range(0, workSheet.nrows):
                for index2 in range(0, workSheet.ncols):
                    if str(workSheet.cell(index1, index2).value).casefold().strip() == "Applicabilité projet".casefold() or str(workSheet.cell(index1, index2).value).casefold().strip() == "Project applicability".casefold():
                        row = index1
                        column = index2
                        flag = True
                        break
                if flag:
                    break

            for index in range(column, workSheet.ncols):
                if workSheet.cell(projectsRow, index).value != "":
                    projects.append(workSheet.cell(projectsRow, index).value)
            for project in projects:
                self.tab1.combo1.addItem(project)
        except:
            self.tab1.textbox.setText("ERROR: when trying to import project names" + DOCPath.split('/')[-1])


    def ButtonSaveConfigClick(self):

        if self.tab1.myTextBox1.toPlainText() or self.tab1.myTextBox2.toPlainText() or self.tab1.myTextBox3.toPlainText():
            data = {}
            list_elements = []

            data['name'] = 'FSE TSD File'
            if self.DOC3Path is None or self.DOC3Path == "":
                data['value'] = 'null'
            else:
                data['value'] = self.DOC3Path
            list_elements.append(data)


            data = {}
            data['name'] = 'TSD Vehicle Function File'
            if self.DOC4Path is None or self.DOC4Path == "":
                data['value'] = 'null'
            else:
                data['value'] = self.DOC4Path
            list_elements.append(data)


            data = {}
            data['name'] = 'TSD System File'
            if self.DOC5Path is None or self.DOC5Path == "":
                data['value'] = 'null'
            else:
                data['value'] = self.DOC5Path
            list_elements.append(data)


            data = {}
            data['name'] = 'AMDEC'
            if self.tab1.myTextBox4.toPlainText() is None or self.tab1.myTextBox4.toPlainText() == "":
                data['value'] = 'null'
            else:
                data['value'] = self.tab1.myTextBox4.toPlainText()
            list_elements.append(data)


            data = {}
            data['name'] = 'Diagnostic matrix'
            if self.tab1.myTextBox5.toPlainText() is None or self.tab1.myTextBox5.toPlainText() == "":
                data['value'] = 'null'
            else:
                data['value'] = self.tab1.myTextBox5.toPlainText()
            list_elements.append(data)


            data = {}
            data['name'] = 'Diagnostic messagery (odx)'
            if self.tab1.myTextBox6.toPlainText() is None or self.tab1.myTextBox6.toPlainText() == "":
                data['value'] = 'null'
            else:
                data['value'] = self.tab1.myTextBox6.toPlainText()
            list_elements.append(data)

            data = {}
            data['name'] = 'SubFamily'
            if self.tab1.myTextBox61.toPlainText() is None or self.tab1.myTextBox61.toPlainText() == "":
                data['value'] = 'null'
            else:
                data['value'] = self.tab1.myTextBox61.toPlainText()
            list_elements.append(data)

            data = {}
            data['name'] = 'Check level'
            if self.tab1.combo.currentText() == "Previsional":
                data['value'] = 'Previsional'
            elif self.tab1.combo.currentText() == "Consolidated":
                data['value'] = 'Consolidated'
            elif self.tab1.combo.currentText() == "Final":
                data['value'] = 'Final'
            list_elements.append(data)


            data = {}
            data['name'] = 'Architecture type'
            if self.tab1.combo2.currentText() == "Archi 2010":
                data['value'] = 'Archi 2010'
            elif self.tab1.combo2.currentText() == "Archi NEA R1":
                data['value'] = 'Archi NEA R1'
            elif self.tab1.combo2.currentText() == "Archi NEA R2":
                data['value'] = 'Archi NEA R2'
            list_elements.append(data)

            data = {}
            data['name'] = 'Diversity management'
            if self.tab1.combo3.currentText() == "Codes LCDV":
                data['value'] = 'Codes LCDV'
            elif self.tab1.combo3.currentText() == "Codes EC":
                data['value'] = 'Codes EC'
            list_elements.append(data)

            data = {}
            data['name'] = 'Project name'
            if self.tab1.combo1.currentText() == "Generic":
                data['value'] = 'Generic'
            elif self.tab1.combo1.currentText() == "All":
                data['value'] = 'All'
            list_elements.append(data)


            data = {}
            data['name'] = 'Network type'
            if self.tab2.RadioButtonInternet.isChecked() is True:
                data['value'] = 'Internet'
            else:
                data['value'] = 'Intranet'
            list_elements.append(data)


            data = {}
            data['name'] = 'CESARE Export'
            if self.tab2.myTextBox7.toPlainText() is None or self.tab2.myTextBox7.toPlainText() == "":
                data['value'] = 'null'
            else:
                data['value'] = self.tab2.myTextBox7.toPlainText()
            list_elements.append(data)


            data = {}
            data['name'] = 'Criticity'
            if self.tab2.myTextBox8.toPlainText() is None or self.tab2.myTextBox8.toPlainText() == "":
                data['value'] = 'null'
            else:
                data['value'] = self.tab2.myTextBox8.toPlainText()
            list_elements.append(data)


            data = {}
            data['name'] = 'Customer Effect File'
            if self.tab2.myTextBox9.toPlainText() is None or self.tab2.myTextBox9.toPlainText() == "":
                data['value'] = 'null'
            else:
                data['value'] = self.tab2.myTextBox9.toPlainText()
            list_elements.append(data)


            data = {}
            data['name'] = 'Diversity'
            if self.tab2.myTextBox10.toPlainText() is None or self.tab2.myTextBox10.toPlainText() == "":
                data['value'] = 'null'
            else:
                data['value'] = self.tab2.myTextBox10.toPlainText()
            list_elements.append(data)


            save_config_fileName, _filter = QtWidgets.QFileDialog.getSaveFileName(self.tab1, 'Save File',QtCore.QDir.rootPath(), 'JSON(*.json)')
            with open(save_config_fileName,'w') as outfile:
                json.dump(list_elements,outfile)

    def ButtonLoadConfigClick(self):

        load_config_fileName, _filter = QtWidgets.QFileDialog.getOpenFileName(self.tab1, 'Open File', QtCore.QDir.rootPath(),'*.*')
        try:
            with open(load_config_fileName) as json_file:
                data = json.load(json_file)
                self.list_element = {}
                for element in data:
                    dict = {}
                    dict['value'] = element['value']
                    self.list_element[element['name']] = dict

            if not self.list_element:
                self.list_element = None

            if self.list_element is not None:
                self.flag_load_configuration = True
                if self.list_element["FSE TSD File"]["value"] != "null":
                    self.tab1.myTextBox1.setText(self.list_element["FSE TSD File"]["value"])
                else:
                    self.tab1.myTextBox1.setText("")

                if self.list_element["TSD Vehicle Function File"]["value"] != "null":
                    self.tab1.myTextBox2.setText(self.list_element["TSD Vehicle Function File"]["value"])
                else:
                    self.tab1.myTextBox2.setText("")

                if self.list_element["TSD System File"]["value"] != "null":
                    self.tab1.myTextBox3.setText(self.list_element["TSD System File"]["value"])
                else:
                    self.tab1.myTextBox3.setText("")

                if self.list_element["AMDEC"]["value"] != "null":
                    self.tab1.myTextBox4.setText(self.list_element["AMDEC"]["value"])
                else:
                    self.tab1.myTextBox4.setText("")

                if self.list_element["Diagnostic matrix"]["value"] != "null":
                    self.tab1.myTextBox5.setText(self.list_element["Diagnostic matrix"]["value"])
                else:
                    self.tab1.myTextBox5.setText("")

                if self.list_element["Diagnostic messagery (odx)"]["value"] != "null":
                    self.tab1.myTextBox6.setText(self.list_element["Diagnostic messagery (odx)"]["value"])
                else:
                    self.tab1.myTextBox6.setText("")

                if self.list_element["SubFamily"]["value"] != "null":
                    self.tab1.myTextBox61.setText(self.list_element["SubFamily"]["value"])
                else:
                    self.tab1.myTextBox61.setText("")

                if self.list_element["Diagnostic messagery (odx)"]["value"] != "null":
                    self.tab1.myTextBox6.setText(self.list_element["Diagnostic messagery (odx)"]["value"])
                else:
                    self.tab1.myTextBox6.setText("")

                if self.list_element["Check level"]["value"] == "Previsional":
                    self.tab1.combo.setCurrentIndex(1)
                elif self.list_element["Check level"]["value"] == "Consolidated":
                    self.tab1.combo.setCurrentIndex(2)
                elif self.list_element["Check level"]["value"] == "Final":
                    self.tab1.combo.setCurrentIndex(0)

                if self.list_element["Architecture type"]["value"] == "Archi 2010":
                    self.tab1.combo2.setCurrentIndex(0)
                elif self.list_element["Architecture type"]["value"] == "Archi NEA R1":
                    self.tab1.combo2.setCurrentIndex(1)
                elif self.list_element["Architecture type"]["value"] == "Archi NEA R2":
                    self.tab1.combo2.setCurrentIndex(2)

                if self.list_element["Diversity management"]["value"] == "Codes LCDV":
                    self.tab1.combo3.setCurrentIndex(0)
                elif self.list_element["Diversity management"]["value"] == "Codes EC":
                    self.tab1.combo3.setCurrentIndex(1)

                if self.list_element["Project name"]["value"] == "Generic":
                    self.tab1.combo1.setCurrentIndex(0)
                elif self.list_element["Project name"]["value"] == "All":
                    self.tab1.combo1.setCurrentIndex(1)

                if self.list_element["Network type"]["value"] == "Internet":
                    self.tab2.RadioButtonInternet.setChecked(True)
                elif self.list_element["Network type"]["value"] == "Intranet":
                    self.tab2.RadioButtonIntranet.setChecked(True)

                if self.list_element["CESARE Export"]["value"] != "null":
                    self.tab2.myTextBox7.setText(self.list_element["CESARE Export"]["value"])
                else:
                    self.tab2.myTextBox7.setText("")

                if self.list_element["Criticity"]["value"] != "null":
                    self.tab2.myTextBox8.setText(self.list_element["Criticity"]["value"])
                else:
                    self.tab2.myTextBox8.setText("")

                if self.list_element["Customer Effect File"]["value"] != "null":
                    self.tab2.myTextBox9.setText(self.list_element["Customer Effect File"]["value"])
                else:
                    self.tab2.myTextBox9.setText("")

                if self.list_element["Diversity"]["value"] != "null":
                    self.tab2.myTextBox10.setText(self.list_element["Diversity"]["value"])
                else:
                    self.tab2.myTextBox10.setText("")


        except:
            pass



    def initUIOptions(self, tab):

        tab.lblUser = QLabel("USER:", tab)
        tab.lblUser.move(300, 25)
        tab.TextBoxUser = QtWidgets.QLineEdit(tab)
        tab.TextBoxUser.resize(200, 25)
        tab.TextBoxUser.move(350, 20)
        # tab.TextBoxUser.setText(getpass.getuser())
        tab.TextBoxUser.setText("E518720")


        tab.lblPass = QLabel("PASSWORD:", tab)
        tab.lblPass.move(570, 25)
        tab.TextBoxPass = QtWidgets.QLineEdit(tab)
        tab.TextBoxPass.resize(180, 25)
        tab.TextBoxPass.move(660, 20)
        tab.TextBoxPass.setEchoMode((QLineEdit.Password))
        tab.TextBoxPass.setText("Cst98988")


        # File Selection Dialog5
        tab.lbl6 = QLabel("Family list export(CESARE):", tab)
        tab.lbl6.move(205, 145)
        tab.myTextBox7 = QtWidgets.QTextEdit(tab)
        tab.myTextBox7.resize(460, 25)
        tab.myTextBox7.move(410, 140)
        tab.myTextBox7.setReadOnly(True)

        tab.link2 = QLabel('''<a href=''' + self.DOC8Link + '''>DocInfo Reference: 02043_18_05471</a>''', tab)
        tab.link2.setOpenExternalLinks(True)
        tab.link2.move(420, 167)


        tab.button7 = QPushButton('...', tab)
        tab.button7.move(870, 141)
        tab.button7.resize(45, 22)
        tab.button7.clicked.connect(self.openFileNameDialog30)


        # File Selection Dialog4
        tab.lbl5 = QLabel("Criticity configuration file:", tab)
        tab.lbl5.move(205,215)
        tab.myTextBox8 = QtWidgets.QTextEdit(tab)
        tab.myTextBox8.resize(460, 25)
        tab.myTextBox8.move(410, 210)
        tab.myTextBox8.setReadOnly(True)

        tab.link1 = QLabel('''<a href='''+self.DOC9Link+'''>DocInfo Reference: 02043_18_05474</a>''', tab)
        tab.link1.setOpenExternalLinks(True)
        tab.link1.move(420, 237)

        tab.button8 = QPushButton('...', tab)
        tab.button8.clicked.connect(self.openFileNameDialog40)
        tab.button8.move(870, 211)
        tab.button8.resize(45, 22)


        # File Selection Dialog6
        tab.lbl7 = QLabel("Customer effect file:", tab)
        tab.lbl7.move(205, 275)
        tab.myTextBox9 = QtWidgets.QTextEdit(tab)
        tab.myTextBox9.resize(460, 25)
        tab.myTextBox9.move(410, 270)
        tab.myTextBox9.setReadOnly(True)

        tab.link3 = QLabel('''<a href=''' + self.DOC7Link + '''>DocInfo Reference: 02043_18_05499</a>''', tab)
        tab.link3.setOpenExternalLinks(True)
        tab.link3.move(420, 297)

        tab.button9 = QPushButton('...', tab)
        tab.button9.clicked.connect(self.openFileNameDialog9)
        tab.button9.move(870, 271)
        tab.button9.resize(45, 22)

        # File Selection Dialog9
        tab.lbl10 = QLabel("Diversity management file:", tab)
        tab.lbl10.move(205, 335)
        tab.myTextBox10 = QtWidgets.QTextEdit(tab)
        tab.myTextBox10.resize(460, 25)
        tab.myTextBox10.move(410,330)
        tab.myTextBox10.setReadOnly(True)

        tab.link4 = QLabel('''<a href=''' + self.DOC13Link + '''>DocInfo Reference: 02016_11_04964</a>''', tab)
        tab.link4.setOpenExternalLinks(True)
        tab.link4.move(420, 357)


        tab.button10 = QPushButton('...', tab)
        tab.button10.clicked.connect(self.openFileNameDialog10)
        tab.button10.move(870, 331)
        tab.button10.resize(45, 22)

        tab.labelInternetAndIntranet = QLabel("Network type:", tab)
        tab.labelInternetAndIntranet.move(300, 60)
        tab.RadioButtonInternet = QRadioButton(self.tab2)
        tab.RadioButtonInternet.setText("Internet link")
        tab.RadioButtonInternet.setChecked(True)
        tab.RadioButtonIntranet = QRadioButton(self.tab2)
        tab.RadioButtonIntranet.setText("Intranet link")
        tab.RadioButtonInternet.toggled.connect(self.ToggleLink)
        tab.RadioButtonIntranet.toggled.connect(self.ToggleLink)
        tab.RadioButtonInternet.move(400, 58)
        tab.RadioButtonIntranet.move(400, 90)

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

        try:
            FileName = response.headers['Content-Disposition'].split('"')[1]
        except:
            error_message = "\nThe file's metadata cannot be properly identified. Please check the network connection!"
            text_box = self.tab1.textbox.toPlainText()
            self.tab1.textbox.setText(text_box + error_message)
            sys.exit(0)

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
        self.DOC13List_2 = []

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
        QApplication.processEvents()


    def buttonClicked(self):


        self.unique_items = []
        self.unique_list = []
        self.refSignature = -1
        self.tab1.textbox_coverage.setText("")
        self.tab1.textbox_convergence.setText("")
        self.tab1.textbox_coverage_1.setText("")
        self.tab1.textbox_convergence_1.setText("")
        self.tab1.textbox_coverage_2.setText("")
        self.tab1.textbox_convergence_2.setText("")

        self.tab1.buttonNew.setEnabled(False)
        self.tab1.pbar.setValue(0)
        self.tab1.textbox.setText("File analysis started...")
        QApplication.processEvents()

        self.start_time = time.time()
        self.checkLevel = str(self.tab1.combo.currentText()).strip().casefold()

        self.tab1.colorTextBox1.setStyleSheet(" background-color: grey ")
        self.tab1.colorTextBox2.setStyleSheet(" background-color: grey ")
        self.tab1.colorTextBox3.setStyleSheet(" background-color: grey ")
        QApplication.processEvents()

        xl = win32.gencache.EnsureDispatch('Excel.Application')
        names = []
        if self.tab1.myTextBox1.toPlainText() != "":
            names.append(self.tab1.myTextBox1.toPlainText().split("/")[-1])
        if self.tab1.myTextBox2.toPlainText() != "":
            names.append(self.tab1.myTextBox2.toPlainText().split("/")[-1])
        if self.tab1.myTextBox3.toPlainText() != "":
            names.append(self.tab1.myTextBox3.toPlainText().split("/")[-1])

        self.flag_opened_file = False

        if xl.Workbooks.Count > 0:
            for name in names:
                if any (i.Name == name for i in xl.Workbooks):
                    win32api.MessageBox(0, 'TSD File is already open. Please close before launching the test!', 'Information')
                    self.flag_opened_file = True

        self.flag_subfamily_odx = False
        if self.tab1.myTextBox6.toPlainText() != "" and self.tab1.myTextBox61.toPlainText() == "":
            win32api.MessageBox(0, 'The SubFamily name field is not completed!')
            self.flag_subfamily_odx = True
        elif self.tab1.myTextBox6.toPlainText() == "" and self.tab1.myTextBox61.toPlainText() != "":
            win32api.MessageBox(0, 'The  Diagnostic messagery (odx) file is not selected!')
            self.flag_subfamily_odx = True

        if not self.flag_opened_file and not self.flag_subfamily_odx:

            if self.tab1.myTextBox1.toPlainText() != "":
                self.DOC3Exists = True
            if self.tab1.myTextBox2.toPlainText() != "":
                self.DOC4Exists = True
            if self.tab1.myTextBox3.toPlainText() != "":
                self.DOC5Exists = True

            self.tab1.textbox.setText("")
            self.tab1.pbar.setValue(0)
            if self.tab1.myTextBox6.toPlainText() is not None and self.tab1.myTextBox6.toPlainText() != "":
                self.Doc15Path = self.tab1.myTextBox6.toPlainText()
            else:
                self.Doc15Path = None

            if not self.tab2.myTextBox7.toPlainText():
                self.DOC8Path = self.download_file(self.DOC8Link)

                extensions = ["xlsx", "xlsm"]
                if self.DOC8Path.split(".")[-1] in extensions:
                    ext = self.DOC8Path.split(".")[0]
                    with zipfile.ZipFile(self.DOC8Path, 'r') as zip_ref:
                        zip_ref.extractall(ext)

                    try:
                        if os.path.isfile(ext + "\docProps\custom.xml"):
                            path = ext + "\docProps\custom.xml"
                            parser = etree.XMLParser(remove_comments=True)
                            tree = objectify.parse(path, parser=parser)
                            root = tree.getroot()
                            self.version_cesare_file = root.find(".//{http://schemas.openxmlformats.org/officeDocument/2006/custom-properties}property[@name = 'psa_version']/{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}lpwstr").text
                            shutil.rmtree(ext, ignore_errors=True)
                    except:
                        shutil.rmtree(ext, ignore_errors=True)
            else:
                self.DOC8Path = self.tab2.myTextBox7.toPlainText()

            if self.DOC8Path == "Error":
                self.tab1.textbox.setText(
                    "ERROR: No network available\nTo continue, please select files for field in the Options tab ")
                return
            if self.DOC8Path == "False":
                return

            if not self.tab2.myTextBox8.toPlainText():
                self.DOC9Path = self.download_file(self.DOC9Link)

                extensions = ["xlsx", "xlsm"]
                if self.DOC9Path.split(".")[-1] in extensions:
                    ext = self.DOC9Path.split(".")[0]
                    with zipfile.ZipFile(self.DOC9Path, 'r') as zip_ref:
                        zip_ref.extractall(ext)

                    try:
                        if os.path.isfile(ext + "\docProps\custom.xml"):
                            path = ext + "\docProps\custom.xml"
                            parser = etree.XMLParser(remove_comments=True)
                            tree = objectify.parse(path, parser=parser)
                            root = tree.getroot()
                            self.version_criticity_file = root.find(
                                ".//{http://schemas.openxmlformats.org/officeDocument/2006/custom-properties}property[@name = 'psa_version']/{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}lpwstr").text
                            shutil.rmtree(ext, ignore_errors=True)
                    except:
                        shutil.rmtree(ext, ignore_errors=True)
            else:
                self.DOC9Path = self.tab2.myTextBox8.toPlainText()

            if not self.tab2.myTextBox9.toPlainText():
                self.DOC7Path = self.download_file(self.DOC7Link)

                extensions = ["xlsx", "xlsm"]
                if self.DOC7Path.split(".")[-1] in extensions:
                    ext = self.DOC7Path.split(".")[0]
                    with zipfile.ZipFile(self.DOC7Path, 'r') as zip_ref:
                        zip_ref.extractall(ext)

                    try:
                        if os.path.isfile(ext + "\docProps\custom.xml"):
                            path = ext + "\docProps\custom.xml"
                            parser = etree.XMLParser(remove_comments=True)
                            tree = objectify.parse(path, parser=parser)
                            root = tree.getroot()
                            self.version_cutomer_effect = root.find(
                                ".//{http://schemas.openxmlformats.org/officeDocument/2006/custom-properties}property[@name = 'psa_version']/{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}lpwstr").text
                            shutil.rmtree(ext, ignore_errors=True)
                    except:
                        shutil.rmtree(ext, ignore_errors=True)
            else:
                self.DOC7Path = self.tab2.myTextBox9.toPlainText()

            if not self.tab2.myTextBox10.toPlainText():
                self.DOC13Path = self.download_file(self.DOC13Link)

                extensions = ["xlsx", "xlsm"]
                if self.DOC13Path.split(".")[-1] in extensions:
                    ext = self.DOC13Path.split(".")[0]
                    with zipfile.ZipFile(self.DOC13Path, 'r') as zip_ref:
                        zip_ref.extractall(ext)

                    try:
                        if os.path.isfile(ext + "\docProps\custom.xml"):
                            path = ext + "\docProps\custom.xml"
                            parser = etree.XMLParser(remove_comments=True)
                            tree = objectify.parse(path, parser=parser)
                            root = tree.getroot()
                            self.version_diversity_file = root.find(
                                ".//{http://schemas.openxmlformats.org/officeDocument/2006/custom-properties}property[@name = 'psa_version']/{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}lpwstr").text
                            shutil.rmtree(ext, ignore_errors=True)
                    except:
                        shutil.rmtree(ext, ignore_errors=True)
            else:
                self.DOC13Path = self.tab2.myTextBox10.toPlainText()

            self.DOC9Dict = OptionalFilesParser.DOC9Parser(self, self.excelApp, self.DOC9Path)
            if self.DOC9Dict == None:
                return

            self.DOC13List, self.DOC13List_2 = OptionalFilesParser.DOC13Parser(self, self.excelApp, self.DOC13Path)
            if self.DOC13List == None or self.DOC13List_2 == None:
                return

            self.DOC8List = OptionalFilesParser.DOC8Parser(self, self.excelApp, self.DOC8Path)
            if self.DOC8List == None:
                return

            if self.Doc15Path is not None and self.Doc15Path != "":
                self.subfamily_name, self.Doc15List = OptionalFilesParser.DOC15Parser(self ,self.Doc15Path)
                if self.subfamily_name == None or self.Doc15List == None:
                    return
            else:
                self.Doc15List = None
                self.subfamily_name = None

            if self.tab1.myTextBox5.toPlainText() is not None and self.tab1.myTextBox5.toPlainText() != "":
                self.DOC14Name = self.tab1.myTextBox5.toPlainText()
            else:
                self.DOC14Name = None

            self.DOC7Name = self.download_file(self.DOC7Link)
            archi_type = self.tab1.combo2.currentText()
            diversity_management = self.tab1.combo3.currentText()

        # elif self.flag_load_configuration and not self.flag_opened_file and not self.flag_subfamily_odx:
        #
        #     if self.list_element["FSE TSD File"]["value"] != "null":
        #         self.DOC3Exists = True
        #     elif self.list_element["TSD Vehicle Function File"]["value"] != "null":
        #         self.DOC4Exists = True
        #     elif self.list_element["TSD System File"]["value"] != "null":
        #         self.DOC5Exists = True
        #
        #     if self.list_element["Network type"]["value"] == "Intranet":
        #         Application.setIntranet(self)
        #     else:
        #         Application.setInternet(self)
        #
        #     if self.list_element["Check level"]["value"] == "Previsional":
        #         self.tab1.combo.setCurrentIndex(1)
        #     elif self.list_element["Check level"]["value"] == "Consolidated":
        #         self.tab1.combo.setCurrentIndex(2)
        #     elif self.list_element["Check level"]["value"] == "Final":
        #         self.tab1.combo.setCurrentIndex(0)
        #
        #     if self.list_element["Project name"]["value"] == "Generic":
        #         self.tab1.combo1.setCurrentIndex(0)
        #     elif self.list_element["Project name"]["value"] == "All":
        #         self.tab1.combo1.setCurrentIndex(1)
        #
        #     if self.list_element["Architecture type"]["value"] == "Archi 2010":
        #         self.tab1.combo2.setCurrentIndex(0)
        #         archi_type = "Archi 2010"
        #     elif self.list_element["Architecture type"]["value"] == "Archi NEA R1":
        #         self.tab1.combo2.setCurrentIndex(1)
        #         archi_type = "Archi NEA R1"
        #     elif self.list_element["Architecture type"]["value"] == "Archi NEA R2":
        #         self.tab1.combo2.setCurrentIndex(2)
        #         archi_type = "Archi NEA R2"
        #
        #     if self.list_element["Diversity management"]["value"] == "Codes LCDV":
        #         self.tab1.combo3.setCurrentIndex(0)
        #         diversity_management = "Codes LCDV"
        #     elif self.list_element["Diversity management"]["value"] == "Codes EC":
        #         self.tab1.combo3.setCurrentIndex(1)
        #         diversity_management = "Codes EC"
        #
        #
        #     self.tab1.textbox.setText("")
        #     self.tab1.pbar.setValue(0)
        #
        #     if self.list_element["Diagnostic messagery (odx)"]['value'] != "null":
        #         self.Doc15Path = self.list_element["Diagnostic messagery (odx)"]['value']
        #     else:
        #          self.Doc15Path = None
        #
        #     if self.list_element["CESARE Export"]["value"] == "null":
        #         self.DOC8Path = self.download_file(self.DOC8Link)
        #
        #         extensions = ["xlsx", "xlsm"]
        #         if self.DOC8Path.split(".")[-1] in extensions:
        #             ext = self.DOC8Path.split(".")[0]
        #             with zipfile.ZipFile(self.DOC8Path, 'r') as zip_ref:
        #                 zip_ref.extractall(ext)
        #
        #             try:
        #                 if os.path.isfile(ext + "\docProps\custom.xml"):
        #                     path = ext + "\docProps\custom.xml"
        #                     parser = etree.XMLParser(remove_comments=True)
        #                     tree = objectify.parse(path, parser=parser)
        #                     root = tree.getroot()
        #                     self.version_cesare_file = root.find(
        #                         ".//{http://schemas.openxmlformats.org/officeDocument/2006/custom-properties}property[@name = 'psa_version']/{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}lpwstr").text
        #                     shutil.rmtree(ext, ignore_errors=True)
        #             except:
        #                 shutil.rmtree(ext, ignore_errors=True)
        #     else:
        #         self.DOC8Path = self.list_element["CESARE Export"]["value"]
        #
        #     if self.DOC8Path == "Error":
        #         self.tab1.textbox.setText(
        #             "ERROR: No network available\nTo continue, please select files for field in the Options tab ")
        #         return
        #     if self.DOC8Path == "False":
        #         return
        #
        #     if self.list_element["Criticity"]["value"] == "null":
        #         self.DOC9Path = self.download_file(self.DOC9Link)
        #
        #         extensions = ["xlsx", "xlsm"]
        #         if self.DOC9Path.split(".")[-1] in extensions:
        #             ext = self.DOC9Path.split(".")[0]
        #             with zipfile.ZipFile(self.DOC9Path, 'r') as zip_ref:
        #                 zip_ref.extractall(ext)
        #
        #             try:
        #                 if os.path.isfile(ext + "\docProps\custom.xml"):
        #                     path = ext + "\docProps\custom.xml"
        #                     parser = etree.XMLParser(remove_comments=True)
        #                     tree = objectify.parse(path, parser=parser)
        #                     root = tree.getroot()
        #                     self.version_criticity_file = root.find(
        #                         ".//{http://schemas.openxmlformats.org/officeDocument/2006/custom-properties}property[@name = 'psa_version']/{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}lpwstr").text
        #                     shutil.rmtree(ext, ignore_errors=True)
        #             except:
        #                 shutil.rmtree(ext, ignore_errors=True)
        #     else:
        #         self.DOC9Path = self.list_element["Criticity"]["value"]
        #
        #     if self.list_element["Customer Effect File"]["value"] == "null":
        #         self.DOC7Path = self.download_file(self.DOC7Link)
        #
        #         extensions = ["xlsx", "xlsm"]
        #         if self.DOC7Path.split(".")[-1] in extensions:
        #             ext = self.DOC7Path.split(".")[0]
        #             with zipfile.ZipFile(self.DOC7Path, 'r') as zip_ref:
        #                 zip_ref.extractall(ext)
        #
        #             try:
        #                 if os.path.isfile(ext + "\docProps\custom.xml"):
        #                     path = ext + "\docProps\custom.xml"
        #                     parser = etree.XMLParser(remove_comments=True)
        #                     tree = objectify.parse(path, parser=parser)
        #                     root = tree.getroot()
        #                     self.version_cutomer_effect = root.find(
        #                         ".//{http://schemas.openxmlformats.org/officeDocument/2006/custom-properties}property[@name = 'psa_version']/{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}lpwstr").text
        #                     shutil.rmtree(ext, ignore_errors=True)
        #             except:
        #                 shutil.rmtree(ext, ignore_errors=True)
        #     else:
        #         self.DOC7Path = self.list_element["Customer Effect File"]["value"]
        #
        #     if self.list_element["Diversity"]["value"] == "null":
        #         self.DOC13Path = self.download_file(self.DOC13Link)
        #
        #         extensions = ["xlsx", "xlsm"]
        #         if self.DOC13Path.split(".")[-1] in extensions:
        #             ext = self.DOC13Path.split(".")[0]
        #             with zipfile.ZipFile(self.DOC13Path, 'r') as zip_ref:
        #                 zip_ref.extractall(ext)
        #
        #             try:
        #                 if os.path.isfile(ext + "\docProps\custom.xml"):
        #                     path = ext + "\docProps\custom.xml"
        #                     parser = etree.XMLParser(remove_comments=True)
        #                     tree = objectify.parse(path, parser=parser)
        #                     root = tree.getroot()
        #                     self.version_diversity_file = root.find(
        #                         ".//{http://schemas.openxmlformats.org/officeDocument/2006/custom-properties}property[@name = 'psa_version']/{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}lpwstr").text
        #                     shutil.rmtree(ext, ignore_errors=True)
        #             except:
        #                 shutil.rmtree(ext, ignore_errors=True)
        #     else:
        #         self.DOC13Path = self.list_element["Diversity"]["value"]
        #
        #     self.DOC9Dict = OptionalFilesParser.DOC9Parser(self, self.excelApp, self.DOC9Path)
        #     if self.DOC9Dict == None:
        #         return
        #
        #     self.DOC13List = OptionalFilesParser.DOC13Parser(self, self.excelApp, self.DOC13Path)
        #     if self.DOC13List == None:
        #         return
        #
        #     self.DOC8List = OptionalFilesParser.DOC8Parser(self, self.excelApp, self.DOC8Path)
        #     if self.DOC8List == None:
        #         return
        #
        #     if self.Doc15Path is not None and self.Doc15Path != "":
        #         self.subfamily_name, self.Doc15List = OptionalFilesParser.DOC15Parser(self, self.Doc15Path)
        #         if self.subfamily_name == None or self.Doc15List == None:
        #             return
        #     else:
        #         self.Doc15List = None
        #         self.subfamily_name = None
        #
        #     # self.DOC8Name = self.download_file(self.DOC8Link)
        #
        #     if self.list_element["Diagnostic matrix"]["value"] != "null":
        #         self.DOC14Name = self.list_element["Diagnostic matrix"]["value"]
        #     else:
        #         self.DOC14Name = None
        #
        #     self.DOC7Name = self.download_file(self.DOC7Link)


        if self.DOC3Exists is True and not self.flag_opened_file and not self.flag_subfamily_odx:

            self.suppressionHeaderRow = 1
            self.tableHeaderRow = 3
            self.codeHeaderRow = 1
            self.measureHeaderRow = 1
            self.diagDebHeaderRow = 1
            self.effClientsHeaderRow = 1
            self.ERHeaderRow = 1
            self.constituantsHeaderRow = 1
            self.sitDeVieHeaderRow = 1
            self.listeMDDHeaderRow = 1

            self.suppressionFirstInfoRow = 2
            self.tableFirstInfoRow = 5
            self.codeFirstInfoRow = 2
            self.measureFirstInfoRow = 2
            self.diagDebFirstInfoRow = 2
            self.effClientsFirstInfoRow = 2
            self.ERFirstInfoRow = 2
            self.constituantsFirstInfoRow = 2
            self.sitDeVieFirstInfoRow = 2
            self.listeMDDFirstInfoRow = 2

            self.return_list = []
            self.DOC3Name = self.download_file(self.DOC3Link)

            self.DOC3Path = self.tab1.myTextBox1.toPlainText()

            try:
                extension = self.DOC3Path.split(".")[-1]
                if extension == "xls":
                    self.DOC3Workbook = xlrd.open_workbook(self.DOC3Path, formatting_info=True)
                else:
                    self.DOC3Workbook = xlrd.open_workbook(self.DOC3Path)
            except:
                self.tab1.textbox.setText("ERROR: when trying to parse the plan type Tableau de synthèse diagnosticabilité file " + self.DOC3Path.split('/')[-1])
                return
            if self.DOC3Workbook == None:
                return
            check = False
            check_indicator = False
            ok_indicator = 0
            ok = 0
            FileMeasure.resetFlags(self)
            FileMeasure.DOC3Info1(self.DOC3Workbook, self)

            self.opening_time = time.time()
            self.tab1.updatesEnabled()

        #GeneralStructure

            if "Test_02043_18_04939_STRUCT_0000" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0000"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0000(self.DOC3Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0005" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0005"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0005(self.DOC3Workbook, self)
                    QApplication.processEvents()
                    
            if "Test_02043_18_04939_STRUCT_0010" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0010"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0010(self.DOC3Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0011" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0011"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0011(self.DOC3Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0020" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0020"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0020(self.DOC3Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0025" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0025"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0025(self.DOC3Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0030" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0030"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0030(self.DOC3Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0035" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0035"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0035(self.DOC3Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0040" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0040"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0040(self.DOC3Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0046" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0046"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0046(self.DOC3Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0051" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0051"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0051(self.DOC3Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0052" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0052"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0052(self.DOC3Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0053" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0053"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0053(self.DOC3Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0055" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0055"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0055(self.DOC3Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0056" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0056"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0056(self.DOC3Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0057" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0057"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0057(self.DOC3Workbook, self)
                    QApplication.processEvents()

            if archi_type == "Archi 2010":
                if "Test_02043_18_04939_STRUCT_0058" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_STRUCT_0058"][self.checkLevel].casefold().strip() != "n/a":
                        GeneralStructureTester.Test_02043_18_04939_STRUCT_0058(self.DOC3Workbook, self)
                        QApplication.processEvents()

                if "Test_02043_18_04939_STRUCT_0061" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_STRUCT_0061"][self.checkLevel].casefold().strip() != "n/a":
                        GeneralStructureTester.Test_02043_18_04939_STRUCT_0061(self.DOC3Workbook, self)
                        QApplication.processEvents()

            elif archi_type == "Archi NEA R1":
                if "Test_02043_18_04939_STRUCT_0059" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_STRUCT_0059"][self.checkLevel].casefold().strip() != "n/a":
                        GeneralStructureTester.Test_02043_18_04939_STRUCT_0059(self.DOC3Workbook, self)
                        QApplication.processEvents()

                if "Test_02043_18_04939_STRUCT_0060" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_STRUCT_0060"][self.checkLevel].casefold().strip() != "n/a":
                        GeneralStructureTester.Test_02043_18_04939_STRUCT_0060(self.DOC3Workbook, self)
                        QApplication.processEvents()

                if "Test_02043_18_04939_STRUCT_0062" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_STRUCT_0062"][self.checkLevel].casefold().strip() != "n/a":
                        GeneralStructureTester.Test_02043_18_04939_STRUCT_0062(self.DOC3Workbook, self)
                        QApplication.processEvents()

                if "Test_02043_18_04939_STRUCT_0063" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_STRUCT_0063"][self.checkLevel].casefold().strip() != "n/a":
                        GeneralStructureTester.Test_02043_18_04939_STRUCT_0063(self.DOC3Workbook, self)
                        QApplication.processEvents()

            elif archi_type == "Archi NEA R2":
                pass

            # DOC3
            if "Test_02043_18_04939_STRUCT_0100" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0100"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0100(self.DOC3Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0110" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0110"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0110(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0120" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0120"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0120(self.DOC3Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0130" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0130"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0130(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0140" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0140"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0140(self.DOC3Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0150" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0150"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0150(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0160" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0160"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0160(self.DOC3Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0170" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0170"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0170(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0180" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0180"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0180(self.DOC3Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0190" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0190"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0190(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0200" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0200"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0200(self.DOC3Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0210" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0210"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0210(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0220" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0220"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0220(self.DOC3Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0230" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0230"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0230(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0240" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0240"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0240(self.DOC3Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0250" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0250"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0250(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0260" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0260"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0260(self.DOC3Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0270" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0270"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0270(self.excelApp, self.DOC3Workbook, self, self.DOC3Name)
                    QApplication.processEvents()

        # Wholeness
            if ok == 0 or ok == 1:

                if "Test_02043_18_04939_WHOLENESS_1050" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1050"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1050(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1055" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1055"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1055(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1060" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1060"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1060(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1061" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1061"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1061(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1062" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1062"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1062(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1070" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1070"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1070(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1240" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1240"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1240(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1600" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1600"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1600(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1600" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1600"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1601(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1602" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1602"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1602(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1603" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1603"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1603(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1604" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1604"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1604(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1605" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1605"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1605(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1606" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1606"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1606(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1607" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1607"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1607(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1608" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1608"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1608(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1609" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1609"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1609(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1610" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1610"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1610(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1611" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1611"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1611(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1612" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1612"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1612(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1613" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1613"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1613(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1615" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1615"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1615(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1616" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1616"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1616(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1617" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1617"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1617(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1618" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1618"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1618(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1619" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1619"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1619(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1620" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1620"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1620(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1621" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1621"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1621(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1622" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1622"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1622(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1623" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1623"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1623(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1624" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1624"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1624(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1625" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1625"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1625(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1626" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1626"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1626(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1627" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1627"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1627(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1628" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1628"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1628(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1629" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1629"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1629(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1630" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1630"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1630(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1631" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1631"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1631(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1632" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1632"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1632(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1650" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1650"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1650(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1651" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1651"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1651(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1652" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1652"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1652(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1653" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1653"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1653(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1654" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1654"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1654(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1655" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1655"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1655(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1656" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1656"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1656(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1657" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1657"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1657(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1658" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1658"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1658(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1659" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1659"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1659(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1660" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1660"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1660(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1661" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1661"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1661(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1662" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1662"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1662(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1663" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1663"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1663(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1664" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1664"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1664(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1684" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1684"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1684(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1685" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1685"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1685(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1686" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1686"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1686(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1687" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1687"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1687(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1688" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1688"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1688(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1689" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1689"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1689(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1690" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1690"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1690(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1691" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1691"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1691(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1692" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1692"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1692(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1693" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1693"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1693(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1700" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1700"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1700(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1701" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1701"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1701(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1702" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1702"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1702(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1703" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1703"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1703(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1704" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1704"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1704(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1705" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1705"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1705(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1706" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1706"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1706(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1707" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1707"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1707(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1708" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1708"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1708(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1709" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1709"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1709(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1710" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1710"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1710(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1711" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1711"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1711(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1712" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1712"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1712(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1713" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1713"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1713(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1714" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1714"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1714(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1715" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1715"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1715(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1716" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1716"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1716(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1717" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1717"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1717(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1718" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1718"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1718(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1719" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1719"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1719(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1750" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1750"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1750(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1751" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1751"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1751(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1752" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1752"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1752(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1753" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1753"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1753(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1754" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1754"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1754(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1755" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1755"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1755(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1756" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1756"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1756(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1757" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1757"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1757(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1758" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1758"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1758(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1759" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1759"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1759(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1800" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1800"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1800(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1801" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1801"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1801(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1802" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1802"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1802(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1803" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1803"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1803(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1810" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1810"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1810(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1811" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1811"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1811(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1812" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1812"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1812(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1813" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1813"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1813(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1814" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1814"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1814(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1815" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1815"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1815(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1820" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1820"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1820(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1821" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1821"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1821(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1822" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1822"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1822(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1823" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1823"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1823(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1824" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1824"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1824(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1825" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1825"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1825(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1830" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1830"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1830(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1831" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1831"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1831(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1840" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1840"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1840(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1841" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1841"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1841(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1


                #  Coherence checks

                if "Test_02043_18_04939_COH_2000" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2000"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2000(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2001" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2001"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2001(self.DOC3Workbook, self, self.DOC8List)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2002" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2002"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2002(self.DOC3Workbook, self, self.DOC8List)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2004" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2004"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2004(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2005" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2005"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2005(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2006" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2006"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2006(self.DOC3Workbook, self, self.DOC8List)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2007" in self.DOC9Dict:
                    if self.DOC14Name:
                        if self.DOC9Dict["Test_02043_18_04939_COH_2007"][self.checkLevel].casefold().strip() != "n/a":
                            check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2007(self.excelApp, self.DOC3Workbook, self, self.DOC14Name)
                            QApplication.processEvents()
                            if check_indicator == True:
                                ok_indicator = 1

                if "Test_02043_18_04939_COH_2008" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2008"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2008(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2009" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2009"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2009(self.DOC3Workbook, self, self.DOC8List)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2010" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2010"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2010(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2020" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2020"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2020(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2030" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2030"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2030(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2040" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2040"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2040(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2050" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2050"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2050(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2061" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2061"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2061(self.excelApp, self.DOC3Workbook, self, self.DOC7Path)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                #check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2091(self.DOC3Workbook, self)

                if "Test_02043_18_04939_COH_2100" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2100"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2100(self.DOC3Workbook, self, self.DOC8List)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2110" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2110"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2110(self.DOC3Workbook, self, self.DOC8List)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2140" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2140"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2140(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2150" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2150"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2150(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2160" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2160"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2160(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2190" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2190"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2190(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2210" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2210"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2210(self.DOC3Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2230" in self.DOC9Dict:
                    if self.Doc15Path is not None and self.Doc15Path != "":
                        if self.DOC9Dict["Test_02043_18_04939_COH_2230"][self.checkLevel].casefold().strip() != "n/a":
                            check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2230(self.DOC3Workbook, self, self.subfamily_name, self.Doc15List)
                            QApplication.processEvents()
                            if check_indicator == True:
                                ok_indicator = 1

                if diversity_management == "Codes LCDV":

                    if "Test_02043_18_04939_COH_2240" in self.DOC9Dict:
                        if self.DOC9Dict["Test_02043_18_04939_COH_2240"][self.checkLevel].casefold().strip() != "n/a":
                            check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2240(self.DOC3Workbook, self, self.DOC13List)
                            QApplication.processEvents()
                            QApplication.processEvents()
                            if check_indicator == True:
                                ok_indicator = 1

                    if "Test_02043_18_04939_COH_2251" in self.DOC9Dict:
                        if self.DOC9Dict["Test_02043_18_04939_COH_2251"][self.checkLevel].casefold().strip() != "n/a":
                            check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2251(self.DOC3Workbook, self,self.DOC13List)
                            QApplication.processEvents()
                            QApplication.processEvents()
                            if check_indicator == True:
                                ok_indicator = 1

                elif diversity_management == "Codes EC":

                    if "Test_02043_18_04939_COH_2260" in self.DOC9Dict:
                        if self.DOC9Dict["Test_02043_18_04939_COH_2260"][self.checkLevel].casefold().strip() != "n/a":
                            check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2260(self.DOC3Workbook, self,self.DOC13List_2)
                            QApplication.processEvents()
                            if check_indicator == True:
                                ok_indicator = 1

                    if "Test_02043_18_04939_COH_2270" in self.DOC9Dict:
                        if self.DOC9Dict["Test_02043_18_04939_COH_2270"][self.checkLevel].casefold().strip() != "n/a":
                            check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2270(self.DOC3Workbook, self,self.DOC13List_2)
                            QApplication.processEvents()
                            if check_indicator == True:
                                ok_indicator = 1

                self.coverage = IndicatorTester.coverageIndicator(self.DOC3Workbook, self) * 100
                self.tab1.textbox_coverage.setText(str(self.coverage)[0:4] + "%")
                self.IncrementProgressBar()
                QApplication.processEvents()

                self.convergence = IndicatorTester.convergenceIndicator(self.DOC3Workbook, self, self.DOC3Path) * 100
                self.tab1.textbox_convergence.setText(str(self.convergence)[0:4] + "%")
                self.IncrementProgressBar()
                QApplication.processEvents()

                if ok_indicator == 1:
                    self.tab1.colorTextBox1.setStyleSheet("background-color: red")
                    self.status = "Failed"

                else:
                    self.tab1.colorTextBox1.setStyleSheet("background-color: green")
                    self.status = "Passed"


                self.end_time = time.time()


                if self.DOC3Path.split('.')[-1] == "xls":
                    ExcelEdit.ExcelWrite_del_information(self.return_list, self.DOC3Path, self, self.DOC3Workbook)
                elif self.DOC3Path.split('.')[-1] in ["xlsx","xlsm"]:
                    ExcelEdit.ExcelWrite2(self.return_list, self.DOC3Path, self, self.DOC3Path)

                self.tab1.pbar.setValue(100)

                if not self.DOC4Exists and not self.DOC5Exists:
                    win32api.MessageBox(0, 'File analysis completed !', 'Information')
                    self.tab1.buttonNew.setEnabled(True)
                    self.tab1.save_config.setEnabled(True)

            elif ok == 1:
                self.tab1.colorTextBox1.setStyleSheet("background-color: red")
                self.status = "Fail"
                self.tab1.buttonNew.setEnabled(True)
                self.tab1.pbar.setValue(100)
                self.end_time = time.time()
                ExcelEdit.WriteReportInformationSheet(self.DOC3Workbook, self)
                self.DOC3Workbook.Save()


        if self.DOC4Exists is True and not self.flag_opened_file and not self.flag_subfamily_odx:

            self.suppressionHeaderRow = 0
            self.tableHeaderRow = 2
            self.diagNeedsHeaderRow = 0
            self.effClientsHeaderRow = 0
            self.fearedEventHeaderRow = 0
            self.systemHeaderRow = 0
            self.opSitHeaderRow = 0
            self.reqTechHeaderRow = 0

            self.suppressionFirstInfoRow = 1
            self.tableFirstInfoRow = 3
            self.diagNeedsFirstInfoRow = 1
            self.effClientsFirstInfoRow = 1
            self.fearedEventFirstInfoRow = 1
            self.systemFirstInfoRow = 1
            self.opSitFirstInfoRow = 1
            self.reqTechFirstInfoRow = 1

            self.return_list = []
            self.DOC4Name = self.download_file(self.DOC4Link)
            self.DOC5Name = self.download_file(self.DOC5Link)

            self.DOC4Path = self.tab1.myTextBox2.toPlainText()

            try:
                extension = self.DOC4Path.split(".")[-1]
                if extension == "xls":
                    self.DOC4Workbook = xlrd.open_workbook(self.DOC4Path, formatting_info=True)
                else:
                    self.DOC4Workbook = xlrd.open_workbook(self.DOC4Path)
            except:
                self.tab1.textbox.setText("ERROR: when trying to parse the plan type TSD Vehicle Function file " + self.DOC4Path.split('/')[-1])
                return

            if self.DOC4Workbook == None:
                return
            check = False
            check_indicator = False
            ok_indicator = 0
            ok = 0

            FileMeasure.resetFlags(self)
            FileMeasure.DOC4Info1(self.DOC4Workbook, self)
            self.opening_time = time.time()

            # GeneralStructure

            if "Test_02043_18_04939_STRUCT_0000" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0000"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0000(self.DOC4Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            # if "Test_02043_18_04939_STRUCT_0005" in self.DOC9Dict:
            #     if self.DOC9Dict["Test_02043_18_04939_STRUCT_0005"][self.checkLevel].casefold().strip() != "n/a":
            #         GeneralStructureTester.Test_02043_18_04939_STRUCT_0005(self.DOC4Workbook, self)

            if "Test_02043_18_04939_STRUCT_0010" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0010"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0010(self.DOC4Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0011" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0011"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0011(self.DOC4Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0020" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0020"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0020(self.DOC4Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0025" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0025"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0025(self.DOC4Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0030" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0030"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0030(self.DOC4Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0035" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0035"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0035(self.DOC4Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0040" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0040"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0040(self.DOC4Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0046" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0046"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0046(self.DOC4Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0051" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0051"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0051(self.DOC4Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0052" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0052"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0052(self.DOC4Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0053" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0053"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0053(self.DOC4Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0054" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0054"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0054(self.DOC4Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0055" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0055"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0055(self.DOC4Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0056" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0056"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0056(self.DOC4Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0057" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0057"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0057(self.DOC4Workbook, self)
                    QApplication.processEvents()

            if archi_type == "Archi 2010":
                if "Test_02043_18_04939_STRUCT_0058" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_STRUCT_0058"][self.checkLevel].casefold().strip() != "n/a":
                        GeneralStructureTester.Test_02043_18_04939_STRUCT_0058(self.DOC4Workbook, self)
                        QApplication.processEvents()
                if "Test_02043_18_04939_STRUCT_0061" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_STRUCT_0061"][self.checkLevel].casefold().strip() != "n/a":
                        GeneralStructureTester.Test_02043_18_04939_STRUCT_0061(self.DOC4Workbook, self)
                        QApplication.processEvents()
            elif archi_type == "Archi NEA R1":
                if "Test_02043_18_04939_STRUCT_0059" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_STRUCT_0059"][self.checkLevel].casefold().strip() != "n/a":
                        GeneralStructureTester.Test_02043_18_04939_STRUCT_0059(self.DOC4Workbook, self)
                        QApplication.processEvents()
                if "Test_02043_18_04939_STRUCT_0060" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_STRUCT_0060"][self.checkLevel].casefold().strip() != "n/a":
                        GeneralStructureTester.Test_02043_18_04939_STRUCT_0060(self.DOC4Workbook, self)
                        QApplication.processEvents()
                if "Test_02043_18_04939_STRUCT_0062" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_STRUCT_0062"][self.checkLevel].casefold().strip() != "n/a":
                        GeneralStructureTester.Test_02043_18_04939_STRUCT_0062(self.DOC4Workbook, self)
                        QApplication.processEvents()
                if "Test_02043_18_04939_STRUCT_0063" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_STRUCT_0063"][self.checkLevel].casefold().strip() != "n/a":
                        GeneralStructureTester.Test_02043_18_04939_STRUCT_0063(self.DOC4Workbook, self)
                        QApplication.processEvents()
            elif archi_type == "Archi NEA R2":
                pass

        # DOC4
            if "Test_02043_18_04939_STRUCT_0400" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0400"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0400(self.DOC4Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0410" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0410"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0410(self.excelApp, self.DOC4Workbook, self, self.DOC4Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0420" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0420"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0420(self.DOC4Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0430" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0430"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0430(self.excelApp, self.DOC4Workbook, self, self.DOC4Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0440" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0440"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0440(self.DOC4Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0450" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0450"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0450(self.excelApp, self.DOC4Workbook, self, self.DOC4Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0460" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0460"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0460(self.DOC4Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0470" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0470"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0470(self.excelApp, self.DOC4Workbook, self, self.DOC4Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0480" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0480"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0480(self.DOC4Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0490" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0490"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0490(self.excelApp, self.DOC4Workbook, self, self.DOC4Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0500" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0500"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0500(self.DOC4Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0510" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0510"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0510(self.excelApp, self.DOC4Workbook, self, self.DOC4Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0520" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0520"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0520(self.DOC4Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            # if "Test_02043_18_04939_STRUCT_0530" in self.DOC9Dict:
            #     if self.DOC9Dict["Test_02043_18_04939_STRUCT_0530"][self.checkLevel].casefold().strip() != "n/a":
            #         GeneralStructureTester.Test_02043_18_04939_STRUCT_0530(self.excelApp, self.DOC4Workbook, self, self.DOC4Name)

            # Wholeness

            if ok == 1 or ok == 0:

                if "Test_02043_18_04939_WHOLENESS_1050" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1050"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1050(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1055" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1055"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1055(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1060" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1060"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1060(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1061" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1061"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1061(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1062" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1062"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1062(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1070" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1070"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1070(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1240" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1240"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1240(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1300" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1300"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1300(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1301" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1301"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1301(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1302" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1302"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1302(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1303" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1303"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1303(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1304" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1304"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1304(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1305" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1305"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1305(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1306" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1306"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1306(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1307" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1307"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1307(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1308" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1308"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1308(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1309" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1309"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1309(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1310" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1310"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1310(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1311" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1311"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1311(self.DOC4Workbook, self)
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1312" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1312"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1312(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1313" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1313"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1313(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1314" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1314"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1314(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1315" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1315"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1315(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1316" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1316"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1316(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1317" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1317"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1317(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1318" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1318"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1318(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1319" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1319"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1319(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1320" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1320"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1320(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1321" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1321"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1321(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1322" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1322"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1322(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1323" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1323"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1323(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1324" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1324"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1324(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1325" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1325"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1325(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1326" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1326"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1326(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1327" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1327"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1327(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1328" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1328"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1328(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1329" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1329"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1329(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1330" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1330"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1330(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1331" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1331"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1331(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1332" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1332"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1332(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1333" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1333"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1333(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1334" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1334"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1334(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1350" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1350"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1350(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1351" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1351"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1351(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1352" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1352"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1352(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1353" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1353"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1353(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1354" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1354"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1354(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1355" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1355"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1355(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1356" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1356"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1356(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1357" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1357"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1357(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1358" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1358"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1358(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1359" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1359"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1359(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1360" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1360"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1360(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1361" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1361"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1361(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1400" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1400"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1400(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1401" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1401"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1401(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1402" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1402"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1402(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1403" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1403"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1403(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1430" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1430"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1430(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1431" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1431"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1431(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1432" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1432"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1432(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1433" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1433"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1433(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1434" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1434"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1434(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1435" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1435"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1435(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1450" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1450"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1450(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1451" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1451"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1451(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1452" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1452"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1452(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1453" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1453"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1453(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1454" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1454"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1454(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1455" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1455"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1455(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1456" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1456"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1456(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1500" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1500"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1500(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1501" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1501"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1501(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1550" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1550"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1550(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1551" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1551"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1551(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1552" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1552"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1552(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

            # Coherence checks

                if "Test_02043_18_04939_COH_2000" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2000"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2000(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2001" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2001"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2001(self.DOC4Workbook, self, self.DOC8List)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2004" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2004"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2004(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2005" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2005"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2005(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2007" in self.DOC9Dict:
                    if self.DOC14Name:
                        if self.DOC9Dict["Test_02043_18_04939_COH_2007"][self.checkLevel].casefold().strip() != "n/a":
                            check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2007(self.excelApp, self.DOC4Workbook, self, self.DOC14Name)
                            QApplication.processEvents()
                            if check_indicator == True:
                                ok_indicator = 1

                if "Test_02043_18_04939_COH_2010" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2010"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2010(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2020" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2020"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2020(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2030" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2030"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2030(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2040" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2040"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2040(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2050" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2050"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2050(self.DOC4Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2070" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2070"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2070(self.excelApp, self.DOC4Workbook, self, self.DOC7Path)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                #check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2091(self.DOC4Workbook, self)

                if "Test_02043_18_04939_COH_2120" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2120"][self.checkLevel].casefold().strip() != "n/a":
                        if self.DOC5Exists:
                            check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2120(self.excelApp, self.DOC4Workbook, self, self.DOC5Name)
                            QApplication.processEvents()
                            if check_indicator == True:
                                ok_indicator = 1

                if diversity_management == "Codes LCDV":

                    if "Test_02043_18_04939_COH_2241" in self.DOC9Dict:
                        if self.DOC9Dict["Test_02043_18_04939_COH_2241"][self.checkLevel].casefold().strip() != "n/a":
                            check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2241(self.DOC4Workbook, self,self.DOC13List)
                            QApplication.processEvents()
                            if check_indicator == True:
                                ok_indicator = 1

                elif diversity_management == "Codes EC":

                    if "Test_02043_18_04939_COH_2261" in self.DOC9Dict:
                        if self.DOC9Dict["Test_02043_18_04939_COH_2261"][self.checkLevel].casefold().strip() != "n/a":
                            check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2261(self.DOC4Workbook, self,self.DOC13List_2)
                            QApplication.processEvents()
                            if check_indicator == True:
                                ok_indicator = 1


                # self.coverage = IndicatorTester.coverageIndicator(self.DOC4Workbook, self) * 100
                # self.tab1.textbox_coverage_1.setText(str(self.coverage)[0:4] + "%")
                # self.IncrementProgressBar()
                # QApplication.processEvents()
                #
                # self.convergence = IndicatorTester.convergenceIndicator(self.DOC4Workbook, self, self.DOC4Path) * 100
                # self.tab1.textbox_convergence_1.setText(str(self.convergence)[0:4] + "%")
                # self.IncrementProgressBar()
                # QApplication.processEvents()

                if ok_indicator == 1:
                    self.tab1.colorTextBox2.setStyleSheet("background-color: red")
                    self.status = "Failed"

                else:
                    self.tab1.colorTextBox2.setStyleSheet("background-color: green")
                    self.status = "Passed"

                self.end_time = time.time()

                if self.DOC4Path.split('.')[-1] == "xls":
                    ExcelEdit.ExcelWrite_del_information(self.return_list, self.DOC4Path, self, self.DOC4Workbook)
                elif self.DOC4Path.split('.')[-1] in ["xlsx", "xlsm"]:
                    ExcelEdit.ExcelWrite2(self.return_list, self.DOC4Path, self, self.DOC4Path)

                self.tab1.pbar.setValue(100)
                if not self.DOC5Exists:
                    win32api.MessageBox(0, 'File analysis completed !', 'Information')
                    self.tab1.buttonNew.setEnabled(True)
                    self.tab1.save_config.setEnabled(True)


        if self.DOC5Exists is True and not self.flag_opened_file and not self.flag_subfamily_odx:

            self.suppressionHeaderRow = 0
            self.tableHeaderRow = 1
            self.codeHeaderRow = 0
            self.dataCodesHeaderRow = 0
            self.readDataIOHeaderRow = 0
            self.notEmbDiagHeaderRow = 0
            self.techEffHeaderRow = 0
            self.effClientsHeaderRow = 0
            self.fearedEventHeaderRow = 0
            self.partsHeaderRow = 0
            self.variantHeaderRow = 0
            self.situationHeaderRow = 0
            self.degradedModeHeaderRow = 0

            self.suppressionFirstInfoRow = 1
            self.tableFirstInfoRow = 3
            self.codeFirstInfoRow = 1
            self.dataCodesFirstInfoRow = 1
            self.readDataIOFirstInfoRow = 1
            self.notEmbDiagFirstInfoRow = 1
            self.techEffFirstInfoRow = 1
            self.effClientsFirstInfoRow = 1
            self.fearedEventFirstInfoRow = 1
            self.partsFirstInfoRow = 1
            self.variantFirstInfoRow = 1
            self.situationFirstInfoRow = 1
            self.degradedModeFirstInfoRow = 1

            self.return_list = []
            self.DOC5Name = self.download_file(self.DOC5Link)

            self.DOC5Path = self.tab1.myTextBox3.toPlainText()

            try:
                extension = self.DOC5Path.split(".")[-1]
                if extension == "xls":
                    self.DOC5Workbook = xlrd.open_workbook(self.DOC5Path, formatting_info=True)
                else:
                    self.DOC5Workbook = xlrd.open_workbook(self.DOC5Path)
            except:
                self.tab1.textbox.setText("ERROR: when trying to parse the plan type TSD Système file file " + self.DOC5Path.split('/')[-1])
                return


            if self.DOC5Workbook == None:
                return
            check = False
            check_indicator = False
            ok_indicator = 0
            ok = 0

            FileMeasure.resetFlags(self)
            FileMeasure.DOC5Info1(self.DOC5Workbook, self)
            self.opening_time = time.time()

            # GeneralStructure

            if "Test_02043_18_04939_STRUCT_0000" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0000"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0000(self.DOC5Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0005" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0005"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0005(self.DOC5Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0010" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0010"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0010(self.DOC5Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0011" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0011"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0011(self.DOC5Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0020" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0020"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0020(self.DOC5Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0025" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0025"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0025(self.DOC5Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0030" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0030"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0030(self.DOC5Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0035" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0035"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0035(self.DOC5Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0040" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0040"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0040(self.DOC5Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0046" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0046"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0046(self.DOC5Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0051" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0051"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0051(self.DOC5Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0052" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0052"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0052(self.DOC5Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0053" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0053"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0053(self.DOC5Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0054" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0054"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0054(self.DOC5Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0055" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0055"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0055(self.DOC5Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0056" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0056"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0056(self.DOC5Workbook, self)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0057" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0057"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0057(self.DOC5Workbook, self)
                    QApplication.processEvents()


            if archi_type == "Archi 2010":
                if "Test_02043_18_04939_STRUCT_0058" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_STRUCT_0058"][self.checkLevel].casefold().strip() != "n/a":
                        GeneralStructureTester.Test_02043_18_04939_STRUCT_0058(self.DOC5Workbook, self)
                        QApplication.processEvents()
                if "Test_02043_18_04939_STRUCT_0061" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_STRUCT_0061"][self.checkLevel].casefold().strip() != "n/a":
                        GeneralStructureTester.Test_02043_18_04939_STRUCT_0061(self.DOC5Workbook, self)
                        QApplication.processEvents()
            elif archi_type == "Archi NEA R1":
                if "Test_02043_18_04939_STRUCT_0059" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_STRUCT_0059"][self.checkLevel].casefold().strip() != "n/a":
                        GeneralStructureTester.Test_02043_18_04939_STRUCT_0059(self.DOC5Workbook, self)
                        QApplication.processEvents()
                if "Test_02043_18_04939_STRUCT_0060" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_STRUCT_0060"][self.checkLevel].casefold().strip() != "n/a":
                        GeneralStructureTester.Test_02043_18_04939_STRUCT_0060(self.DOC5Workbook, self)
                        QApplication.processEvents()
                if "Test_02043_18_04939_STRUCT_0062" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_STRUCT_0062"][self.checkLevel].casefold().strip() != "n/a":
                        GeneralStructureTester.Test_02043_18_04939_STRUCT_0062(self.DOC5Workbook, self)
                        QApplication.processEvents()
                if "Test_02043_18_04939_STRUCT_0063" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_STRUCT_0063"][self.checkLevel].casefold().strip() != "n/a":
                        GeneralStructureTester.Test_02043_18_04939_STRUCT_0063(self.DOC5Workbook, self)
                        QApplication.processEvents()
            elif archi_type == "Archi NEA R2":
                pass

            # DOC5
            if "Test_02043_18_04939_STRUCT_0700" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0700"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0700(self.DOC5Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0710" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0710"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0710(self.excelApp, self.DOC5Workbook, self, self.DOC5Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0720" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0720"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0720(self.DOC5Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0730" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0730"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0730(self.excelApp, self.DOC5Workbook, self, self.DOC5Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0740" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0740"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0740(self.DOC5Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0750" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0750"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0750(self.excelApp, self.DOC5Workbook, self, self.DOC5Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0760" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0760"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0760(self.DOC5Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0770" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0770"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0770(self.excelApp, self.DOC5Workbook, self, self.DOC5Name)
                    QApplication.processEvents()


            if "Test_02043_18_04939_STRUCT_0780" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0780"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0780(self.DOC5Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0790" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0790"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0790(self.excelApp, self.DOC5Workbook, self, self.DOC5Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0800" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0800"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0800(self.DOC5Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0810" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0810"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0810(self.excelApp, self.DOC5Workbook, self, self.DOC5Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0820" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0820"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0820(self.DOC5Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0830" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0830"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0830(self.excelApp, self.DOC5Workbook, self, self.DOC5Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0840" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0840"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0840(self.DOC5Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0850" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0850"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0850(self.excelApp, self.DOC5Workbook, self, self.DOC5Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0860" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0860"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0860(self.DOC5Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0870" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0870"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0870(self.excelApp, self.DOC5Workbook, self, self.DOC5Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0880" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0880"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0880(self.DOC5Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0890" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0890"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0890(self.excelApp, self.DOC5Workbook, self, self.DOC5Name)
                    QApplication.processEvents()

            if "Test_02043_18_04939_STRUCT_0900" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0900"][self.checkLevel].casefold().strip() != "n/a":
                    check = GeneralStructureTester.Test_02043_18_04939_STRUCT_0900(self.DOC5Workbook, self)
                    QApplication.processEvents()
                    if check == True:
                        ok = 1

            if "Test_02043_18_04939_STRUCT_0910" in self.DOC9Dict:
                if self.DOC9Dict["Test_02043_18_04939_STRUCT_0910"][self.checkLevel].casefold().strip() != "n/a":
                    GeneralStructureTester.Test_02043_18_04939_STRUCT_0910(self.excelApp, self.DOC5Workbook, self, self.DOC5Name)
                    QApplication.processEvents()


            # Wholeness
            if ok == 0 or ok == 1:

                if "Test_02043_18_04939_WHOLENESS_1050" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1050"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1050(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1055" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1055"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1055(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1060" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1060"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1060(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1061" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1061"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1061(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1062" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1062"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1062(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1070" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1070"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1070(self.DOC5Workbook, self)
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1240" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1240"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1240(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1900" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1900"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1900(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1901" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1901"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1901(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1902" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1902"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1902(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1903" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1903"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1903(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1904" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1904"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1904(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1905" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1905"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1905(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1906" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1906"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1906(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1907" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1907"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1907(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1908" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1908"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1908(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1909" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1909"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1909(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1910" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1910"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1910(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1911" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1911"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1911(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1912" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1912"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1912(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1913" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1913"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1913(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1914" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1914"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1914(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1915" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1915"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1915(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1916" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1916"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1916(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1917" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1917"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1917(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1918" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1918"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1918(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1919" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1919"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1919(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1920" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1920"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1920(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1921" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1921"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1921(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1922" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1922"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1922(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1923" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1923"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1923(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1924" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1924"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1924(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1925" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1925"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1925(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1926" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1926"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1926(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1927" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1927"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1927(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1950" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1950"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1950(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1951" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1951"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1951(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1952" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1952"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1952(self.DOC5Workbook, self)
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1953" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1953"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1953(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1954" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1954"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1954(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1955" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1955"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1955(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1956" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1956"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1956(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1957" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1957"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1957(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1958" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1958"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1958(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1959" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1959"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1959(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1960" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1960"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1960(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1961" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1961"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1961(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1962" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1962"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1962(self.DOC5Workbook, self)
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1963" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1963"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1963(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1964" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1964"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1964(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1965" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1965"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1965(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1966" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1966"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1966(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1967" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1967"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1967(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1968" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1968"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1968(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_1969" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_1969"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_1969(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2000" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2000"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2000(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2001" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2001"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2001(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2002" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2002"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2002(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2003" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2003"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2003(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2004" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2004"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2004(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2005" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2005"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2005(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2006" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2006"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2006(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2007" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2007"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2007(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2008" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2008"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2008(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2009" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2009"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2009(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2010" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2010"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2010(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2011" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2011"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2011(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2050" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2050"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2050(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2051" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2051"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2051(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2052" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2052"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2052(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2053" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2053"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2053(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2054" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2054"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2054(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2055" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2055"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2055(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2056" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2056"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2056(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2060" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2060"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2060(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2061" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2061"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2061(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2062" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2062"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2062(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2070" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2070"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2070(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2071" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2071"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2071(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2072" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2072"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2072(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2080" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2080"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2080(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2081" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2081"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2081(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2082" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2082"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2082(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2083" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2083"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2083(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2084" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2084"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2084(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2090" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2090"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2090(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2091" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2091"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2091(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2092" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2092"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2092(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2100" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2100"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2100(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2101" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2101"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2101(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2102" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2102"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2102(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2110" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2110"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2110(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2111" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2111"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2111(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2112" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2112"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2112(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2120" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2120"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2120(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_WHOLENESS_2121" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_WHOLENESS_2121"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = WholenessTester.Test_02043_18_04939_WHOLENESS_2121(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                # Coherence checks

                if "Test_02043_18_04939_COH_2000" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2000"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2000(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2001" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2001"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2001(self.DOC5Workbook, self, self.DOC8List)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2002" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2002"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2002(self.DOC5Workbook, self, self.DOC8List)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2004" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2004"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2004(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2005" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2005"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2005(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2006" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2006"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2006(self.DOC5Workbook, self, self.DOC8Name)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2007" in self.DOC9Dict:
                    if self.DOC14Name:
                        if self.DOC9Dict["Test_02043_18_04939_COH_2007"][self.checkLevel].casefold().strip() != "n/a":
                            check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2007(self.excelApp, self.DOC5Workbook, self, self.DOC14Name)
                            QApplication.processEvents()
                            if check_indicator == True:
                                ok_indicator = 1

                if "Test_02043_18_04939_COH_2008" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2008"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2008(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2009" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2009"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2009(self.DOC5Workbook, self, self.DOC8Name)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2010" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2010"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2010(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2020" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2020"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2020(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2030" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2030"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2030(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2040" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2040"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2040(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2050" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2050"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2050(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2080" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2080"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2080(self.excelApp, self.DOC5Workbook, self, self.DOC7Path)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                #check_indicator =  Coherence_checksTester.Test_02043_18_04939_COH_2091(self.DOC5Workbook, self)

                if "Test_02043_18_04939_COH_2130" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2130"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2130(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2170" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2170"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2170(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2180" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2180"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2180(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2200" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2200"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2200(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2220" in self.DOC9Dict:
                    if self.DOC9Dict["Test_02043_18_04939_COH_2220"][self.checkLevel].casefold().strip() != "n/a":
                        check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2220(self.DOC5Workbook, self)
                        QApplication.processEvents()
                        if check_indicator == True:
                            ok_indicator = 1

                if "Test_02043_18_04939_COH_2230" in self.DOC9Dict:
                    if self.Doc15Path is not None and self.Doc15Path != "":
                        if self.DOC9Dict["Test_02043_18_04939_COH_2230"][self.checkLevel].casefold().strip() != "n/a":
                            check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2230(self.DOC5Workbook, self,self.subfamily_name, self.Doc15List)
                            QApplication.processEvents()
                            if check_indicator == True:
                               ok_indicator = 1

                if diversity_management == "Codes LCDV":

                    if "Test_02043_18_04939_COH_2240" in self.DOC9Dict:
                        if self.DOC9Dict["Test_02043_18_04939_COH_2240"][self.checkLevel].casefold().strip() != "n/a":
                            check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2240(self.DOC5Workbook, self, self.DOC13List)
                            QApplication.processEvents()
                            if check_indicator == True:
                                ok_indicator = 1

                    if "Test_02043_18_04939_COH_2251" in self.DOC9Dict:
                        if self.DOC9Dict["Test_02043_18_04939_COH_2251"][self.checkLevel].casefold().strip() != "n/a":
                            check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2251(self.DOC5Workbook, self,self.DOC13List)
                            QApplication.processEvents()
                            if check_indicator == True:
                                ok_indicator = 1

                elif diversity_management == "Codes EC":

                    if "Test_02043_18_04939_COH_2260" in self.DOC9Dict:
                        if self.DOC9Dict["Test_02043_18_04939_COH_2260"][self.checkLevel].casefold().strip() != "n/a":
                            check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2260(self.DOC5Workbook, self,self.DOC13List_2)
                            QApplication.processEvents()
                            if check_indicator == True:
                                ok_indicator = 1

                    if "Test_02043_18_04939_COH_2270" in self.DOC9Dict:
                        if self.DOC9Dict["Test_02043_18_04939_COH_2270"][self.checkLevel].casefold().strip() != "n/a":
                            check_indicator = Coherence_checksTester.Test_02043_18_04939_COH_2270(self.DOC5Workbook, self,self.DOC13List_2)
                            QApplication.processEvents()
                            if check_indicator == True:
                                ok_indicator = 1


                self.coverage = IndicatorTester.coverageIndicator(self.DOC5Workbook, self) * 100
                self.tab1.textbox_coverage_2.setText(str(self.coverage)[0:4] + "%")
                self.IncrementProgressBar()

                self.convergence = IndicatorTester.convergenceIndicator(self.DOC5Workbook, self, self.DOC5Path) * 100
                self.tab1.textbox_convergence_2.setText(str(self.convergence)[0:4] + "%")
                self.IncrementProgressBar()

                if ok_indicator == 1:
                    self.tab1.colorTextBox3.setStyleSheet("background-color: red")
                    self.status = "Failed"

                else:
                    self.tab1.colorTextBox3.setStyleSheet("background-color: green")
                    self.status = "Passed"


                self.end_time = time.time()

                if self.DOC5Path.split('.')[-1] == "xls":
                    ExcelEdit.ExcelWrite_del_information(self.return_list, self.DOC5Path, self, self.DOC5Workbook)
                elif self.DOC5Path.split('.')[-1] in ["xlsx", "xlsm"]:
                    ExcelEdit.ExcelWrite2(self.return_list, self.DOC5Path, self, self.DOC5Path)

                self.tab1.pbar.setValue(100)
                self.tab1.buttonNew.setEnabled(True)
                self.tab1.save_config.setEnabled(True)
                win32api.MessageBox(0, 'File analysis completed !', 'Information')

if __name__ == '__main__':


    try:
        FindWindow(None, appName)
        windll.user32.MessageBoxW(0, "Application already running", "Warning", 0|48)

    except:
        app = QApplication(sys.argv)
        apel = Test()
        myQLabel = QLabel()
        sys.exit(app.exec_())
