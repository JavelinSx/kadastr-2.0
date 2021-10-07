import os
import shutil
import sqlite3
import sys

from docxtpl import DocxTemplate
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtCore import QDate, QRegExp
from PyQt5.QtGui import QColor, QBrush, QTextCharFormat, QRegExpValidator, QIcon
from PyQt5.QtWidgets import QTableWidgetItem, QMessageBox, QCompleter, QInputDialog
import subprocess

import fileUi.newForm
import fileUi.formAddCity


class startWindow(QtWidgets.QMainWindow, fileUi.newForm.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        # window setting
        self.showMaximized()
        self.setWindowIcon(QIcon('img/main.ico'))
        self.setupUi(self)
        self.tableWidget.setSortingEnabled(True)
        self.tableWidget.hideColumn(9)

        # var
        self.comboBoxCityIndex = None
        self.newCity = None
        self.allInfoClient = None
        self.allInfoAddress = None
        self.allInfoWork = None
        self.allInfoCity = None
        self.lastIndex = None
        self.dbName = "kadastr 2.0.db"
        # get info in datebase
        self.getAllClientInfo()
        self.getAllInfoCity()
        # completer

        # regExp
        self.regExpPhone = QRegExp('^\d\d\d\d\d\d\d\d\d\d\d$')
        self.intValidatorPhone = QRegExpValidator(self.regExpPhone, self.lineEditTelefone)
        self.lineEditTelefone.setValidator(self.intValidatorPhone)

        # buttons

        # page client
        self.pushButtonDelete.clicked.connect(self.deleteClient)
        self.pushButtonSearch.clicked.connect(self.searchClient)
        self.pushButtonUpdate.clicked.connect(self.updateTableClient)
        self.pushButtonViewForgotten.clicked.connect(self.getForgottenClient)
        self.pushButtonViewReady.clicked.connect(self.getInfoCompleteClient)
        # page add client
        self.updateComboBoxCity()
        self.pushButtonAdd.clicked.connect(self.addClient)
        self.pushButtonAddCity.clicked.connect(self.insertCity)
        # page update client
        self.pushButtonOpenFolder.clicked.connect(self.openFolder)
        self.pushButtonEdit.clicked.connect(self.updateInfoClient)
        self.pushButtonClone.clicked.connect(self.cloneClientInfo)


    def updateComboBoxCity(self):
        for i in range(len(self.allInfoCity)):
            self.comboBoxCity.addItem(self.allInfoCity[i][1])

    def msgInfo(self, text):
        QMessageBox.information(self, "Информация",
                                text)

    def querySelect(self, sql, param=None):

        con = sqlite3.connect(self.dbName)
        with con:
            if param is None:
                cur = con.cursor()
                cur.execute(sql)
                res = cur.fetchall()


            else:
                cur = con.cursor()
                cur.execute(sql, param)
                res = cur.fetchall()

        return res

    def queryInsert(self, sql, param=None):

        con = sqlite3.connect(self.dbName)
        with con:
            if param is None:
                cur = con.cursor()
                cur.execute(sql)
                con.commit()
                cur.execute("SELECT last_insert_rowid()")
                index = cur.fetchone()
                self.lastIndex = index
            else:
                cur = con.cursor()
                cur.execute(sql, param)
                con.commit()
                cur.execute("SELECT last_insert_rowid()")
                index = cur.fetchone()
                self.lastIndex = index
        if con:
            con.close()

    def clearAddClientPage(self):
        self.comboBoxProvideServices.setCurrentIndex(0)
        self.lineEditCity.clear()
        self.lineEditAddress.clear()
        self.lineEditSurnameAdd.clear()
        self.lineEditNameAdd.clear()
        self.lineEditMiddleNameAdd.clear()
        self.lineEditTelefoneAdd.clear()

    def getIdService(self, id):

        sqlIdService = "SELECT service from info_client where id = ?"

        return self.querySelect(sqlIdService, (id,))

    def getIDWorkInfo(self, id):

        sqlIdAddressInfo = "SELECT doc_info from info_client where id = ?"

        return self.querySelect(sqlIdAddressInfo, (id,))

    def getIdAddressInfo(self, id):

        sqlIdAddressInfo = "SELECT address_info from info_client where id = ?"

        return self.querySelect(sqlIdAddressInfo, (id,))

    def getIdDocInfo(self, id):

        sqlIdDocInfo = "SELECT doc_info from info_client where id = ?"

        return self.querySelect(sqlIdDocInfo, (id,))

    def getAllClientInfo(self):
        # list of page client

        sqlInfoClient = "SELECT * from info_client"
        self.allInfoClient = self.querySelect(sqlInfoClient)

        sqlInfoAddress = "SELECT * from address_info_client"
        self.allInfoAddress = self.querySelect(sqlInfoAddress)

        sqlInfoWork = "SELECT * from work_info_client"
        self.allInfoWork = self.querySelect(sqlInfoWork)

    def getAllInfoCity(self):
        sqlInfoCity = "SELECT *  from city"
        self.allInfoCity = self.querySelect(sqlInfoCity)

    def insertAddressInfo(self):

        city = self.lineEditCity.text().title()
        address = self.lineEditAddress.text().title()
        sqlAddressInfo = "INSERT INTO address_info_client(city,address) VALUES (?,?)"
        self.lastIndexAddressInfo = self.queryInsert(sqlAddressInfo, (city, address,))

    def insertDocInfo(self):

        passDocSeries = self.lineEditPassDocSeries.text()
        passDocDate = self.dateEditPassDocDate.text()
        passDocInfo = self.lineEditPassDocInfo.text()
        passSnils = self.lineEditPassSnils.text()
        sqlDocInfo = "INSERT INTO doc_info_client(series,date,info,snils) VALUES (?,?,?,?)"
        self.queryInsert(sqlDocInfo, (passDocSeries, passDocDate, passDocInfo, passSnils))

    def insertWorkInfo(self):

        prepayment = self.checkBoxPrepayment.checkState()
        remains = self.checkBoxRemains.checkState()
        work = self.comboBoxWork.currentIndex()
        dateWork = self.dateEditDateWork.text()
        status = self.comboBoxStatus.currentIndex()
        dateStatus = self.comboBoxStatus.currentIndex()
        info = self.textEditInfo.toPlainText()
        sqlWorkInfo = "INSERT INTO work_info_client(prepayment,remains,work,date_work,status,date_status,info) VALUES " \
                      "(?,?,?,?,?,?,?) "
        self.queryInsert(sqlWorkInfo, (prepayment, remains, work, dateWork, status, dateStatus, info))

    def insertInfoClient(self):
        self.insertAddressInfo()
        lastIndexAddressInfo = self.lastIndex
        self.insertWorkInfo()
        lastIndexWorkInfo = self.lastIndex
        self.insertDocInfo()
        lastIndexDocInfo = self.lastIndex
        service = self.comboBoxProvideServices.currentIndex()
        surName = self.lineEditSurnameAdd.text()
        name = self.lineEditNameAdd.text()
        middleName = self.lineEditMiddleNameAdd.text()
        telefone = self.lineEditTelefoneAdd.text()
        sqlInfoClient = "INSERT INTO info_client(sur_name,name,middle_name,telefone,address_info,doc_info,work_info," \
                        "service) VALUES (?,?,?,?,?,?,?,?) "
        self.queryInsert(sqlInfoClient, (
            surName, name, middleName, telefone, int(lastIndexAddressInfo[0]), int(lastIndexDocInfo[0]),
            int(lastIndexWorkInfo[0]), int(service)))

    def insertCity(self):
        check = False
        allCity = []
        sqlCItyIndex = "INSERT INTO city (city_name) values(?)"
        for i in range(self.comboBoxCity.count()):
            allCity.append(self.comboBoxCity.itemText(i))
        dlg = QInputDialog(self)
        dlg.setInputMode(QInputDialog.TextInput)
        dlg.setWindowTitle("Добавить населенный пункт")
        dlg.setLabelText("")
        dlg.resize(300, 100)
        ok = dlg.exec_()
        city = dlg.textValue()
        print(self.allInfoCity)
        if ok:

            for i in range(len(allCity)):
                if city == allCity[i]:
                    check = False
                else:
                    check = True
            if check:
                self.comboBoxCity.addItem(city)

                self.queryInsert(sqlCItyIndex, (city,))




    def updateTableClient(self):
        self.getAllClientInfo()
        lenInfoClient = len(self.allInfoClient)

        for i in range(lenInfoClient):
            self.tableWidget.insertRow(self.tableWidget.rowCount())

            statusWorkText = self.comboBoxStatus.itemText(self.allInfoWork[i][3])
            addressCityText = self.allInfoAddress[i][1]
            surNameText = self.allInfoClient[i][1]
            nameText = self.allInfoClient[i][2]
            middleNameText = self.allInfoClient[i][3]
            telefoneText = self.allInfoClient[i][4]
            serviceText = self.comboBoxProvideServices.itemText(self.allInfoClient[i][8])

            item = QTableWidgetItem()
            item.setText(statusWorkText)  # status
            self.tableWidget.setItem(i, 0, item)

            item = QTableWidgetItem()
            item.setText(addressCityText)  # city
            self.tableWidget.setItem(i, 1, item)

            item = QTableWidgetItem()
            item.setText(surNameText)  # surName
            self.tableWidget.setItem(i, 2, item)

            item = QTableWidgetItem()
            item.setText(nameText)  # name
            self.tableWidget.setItem(i, 3, item)

            item = QTableWidgetItem()
            item.setText(middleNameText)  # middleName
            self.tableWidget.setItem(i, 4, item)

            item = QTableWidgetItem()
            item.setText(telefoneText)  # telefone
            self.tableWidget.setItem(i, 5, item)

            item = QTableWidgetItem()
            item.setText(serviceText)  # service
            self.tableWidget.setItem(i, 6, item)

    def addClient(self):
        self.insertInfoClient()
        self.clearAddClientPage()

    def openFolder(self):
        pass

    def updateInfoClient(self):
        pass

    def deleteClient(self):
        pass

    def searchClient(self):
        pass

    def getForgottenClient(self):
        pass

    def getInfoCompleteClient(self):
        pass

    def cloneClientInfo(self):
        pass


def main():
    import ctypes
    myappid = 'mycompany.myproduct.subproduct.version'  # arbitrary string
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

    app = QtWidgets.QApplication(sys.argv)
    window = startWindow()
    window.show()
    app.setStyle('Fusion')
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
