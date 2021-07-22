# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'regbank_converter.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(300, 200)
        self.ConvertButton = QtWidgets.QPushButton(Dialog)
        self.ConvertButton.setGeometry(QtCore.QRect(100, 140, 100, 40))
        self.ConvertButton.setObjectName("ConvertButton")
        self.RegbankButton = QtWidgets.QPushButton(Dialog)
        self.RegbankButton.setGeometry(QtCore.QRect(55, 90, 50, 20))
        self.RegbankButton.setObjectName("RegbankButton")
        self.StyleButton = QtWidgets.QPushButton(Dialog)
        self.StyleButton.setGeometry(QtCore.QRect(195, 90, 50, 20))
        self.StyleButton.setObjectName("StyleButton")
        self.StyleLabel = QtWidgets.QLabel(Dialog)
        self.StyleLabel.setGeometry(QtCore.QRect(160, 20, 120, 60))
        self.StyleLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.StyleLabel.setObjectName("StyleLabel")
        self.RegbankLabel = QtWidgets.QLabel(Dialog)
        self.RegbankLabel.setGeometry(QtCore.QRect(20, 20, 120, 60))
        self.RegbankLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.RegbankLabel.setObjectName("RegbankLabel")

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.ConvertButton.setText(_translate("Dialog", "Convert"))
        self.RegbankButton.setText(_translate("Dialog", "Load"))
        self.StyleButton.setText(_translate("Dialog", "Load"))
        self.StyleLabel.setText(_translate("Dialog", "style.docx"))
        self.RegbankLabel.setText(_translate("Dialog", "regbank.docx"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())

