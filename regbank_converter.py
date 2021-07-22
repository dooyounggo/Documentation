# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'regbank_converter.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!

import os
from PyQt5 import QtCore, QtGui, QtWidgets
from generate_datasheet import generate_datasheet


class Ui_Dialog(object):
    def setupUi(self, Dialog, app):
        Dialog.setObjectName("Dialog")
        Dialog.resize(480, 240)
        self.ConvertButton = QtWidgets.QPushButton(Dialog)
        self.ConvertButton.setGeometry(QtCore.QRect(190, 110, 100, 40))
        self.ConvertButton.setObjectName("ConvertButton")
        self.RegbankButton = QtWidgets.QPushButton(Dialog)
        self.RegbankButton.setGeometry(QtCore.QRect(100, 110, 50, 20))
        self.RegbankButton.setObjectName("RegbankButton")
        self.StyleButton = QtWidgets.QPushButton(Dialog)
        self.StyleButton.setGeometry(QtCore.QRect(330, 110, 50, 20))
        self.StyleButton.setObjectName("StyleButton")
        self.StyleLabel = QtWidgets.QLabel(Dialog)
        self.StyleLabel.setGeometry(QtCore.QRect(250, 40, 210, 60))
        self.StyleLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.StyleLabel.setWordWrap(True)
        self.StyleLabel.setObjectName("StyleLabel")
        self.RegbankLabel = QtWidgets.QLabel(Dialog)
        self.RegbankLabel.setGeometry(QtCore.QRect(20, 40, 210, 60))
        self.RegbankLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.RegbankLabel.setWordWrap(True)
        self.RegbankLabel.setObjectName("RegbankLabel")
        self.TextBrowser = QtWidgets.QTextBrowser(Dialog)
        self.TextBrowser.setGeometry(QtCore.QRect(20, 160, 440, 70))
        self.TextBrowser.setObjectName("TextBrowser")
        self.RegDescLabel = QtWidgets.QLabel(Dialog)
        self.RegDescLabel.setGeometry(QtCore.QRect(50, 10, 150, 20))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.RegDescLabel.setFont(font)
        self.RegDescLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.RegDescLabel.setObjectName("RegDescLabel")
        self.DataStyleLabel = QtWidgets.QLabel(Dialog)
        self.DataStyleLabel.setGeometry(QtCore.QRect(280, 10, 150, 20))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.DataStyleLabel.setFont(font)
        self.DataStyleLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.DataStyleLabel.setObjectName("DataStyleLabel")

        self.app = app
        self.regbank_path = 'regbank.docx'
        self.style_path = 'style.docx'

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

        self.RegbankButton.clicked.connect(self.regbank_button_clicked)
        self.StyleButton.clicked.connect(self.style_button_clicked)
        self.ConvertButton.clicked.connect(self.convert_button_clicked)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Register Bank Description Converter"))
        self.ConvertButton.setToolTip(_translate("Dialog", "Convert register bank description."))
        self.ConvertButton.setText(_translate("Dialog", "Convert"))
        self.RegbankButton.setToolTip(_translate("Dialog", "Load Magillem-generated register description."))
        self.RegbankButton.setText(_translate("Dialog", "Load"))
        self.StyleButton.setToolTip(_translate("Dialog", "Load Telechips datasheet style."))
        self.StyleButton.setText(_translate("Dialog", "Load"))
        self.StyleLabel.setText(_translate("Dialog", "style.docx"))
        self.RegbankLabel.setText(_translate("Dialog", "regbank.docx"))
        self.RegDescLabel.setText(_translate("Dialog", "Register Description"))
        self.DataStyleLabel.setText(_translate("Dialog", "Datasheet Style"))

    def regbank_button_clicked(self):
        fname = QtWidgets.QFileDialog.getOpenFileName(caption='Select Register Bank Description',
                                                      filter='MS word docx (*.docx)')[0]
        if fname:
            self.regbank_path = fname
            self.RegbankLabel.setText(fname)
            self.uiprint('Register description file path: ' + fname, color=(0, 0, 0.5))

    def style_button_clicked(self):
        fname = QtWidgets.QFileDialog.getOpenFileName(caption='Select Datasheet Style',
                                                      filter='MS word docx (*.docx)')[0]
        if fname:
            self.style_path = fname
            self.StyleLabel.setText(fname)
            self.uiprint('Datasheet style file path: ' + fname, color=(0, 0.5, 0))

    def convert_button_clicked(self):
        fname = QtWidgets.QFileDialog.getSaveFileName(caption='Choose Output Directory and Name',
                                                      filter='MS word docx (*.docx)',
                                                      directory='./datasheet.docx')[0]
        if fname:
            if not os.path.exists(self.regbank_path):
                self.uiprint(f'Error: Register description does not exists ({self.regbank_path}).', color=(0.5, 0, 0))
                return
            if not os.path.exists(self.style_path):
                self.uiprint(f'Error: Datasheet style does not exists ({self.style_path}).', color=(0.5, 0, 0))
                return
            generate_datasheet(self.regbank_path, self.style_path, fname)
            self.uiprint('Conversion complete!')

    def uiprint(self, string_like, color=(0, 0, 0)):
        string = str(string_like)
        color_code = ''.join([f'{int(min(rgb, 1.0)*255):02x}' for rgb in color])
        colored_string = f'<span style=\" color: #{color_code};\">{string}</span>'
        self.TextBrowser.append(colored_string)
        self.app.processEvents()
        print(string)


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog, app)
    Dialog.show()
    sys.exit(app.exec_())

