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
        self.tableWidget.hideColumn(7)
        os.chdir("//lena/межевание/test")
        # var
        self.newCity = None
        self.clientInfo = None
        self.allInfoClient = None
        self.allInfoAddress = None
        self.allInfoWork = None
        self.allInfoCity = None
        self.lastIndex = None
        self.pathFolderClient = None
        self.dbName = "kadastr 2.0.db"

        # get info in date base
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
        self.pushButtonChangeCity.clicked.connect(self.updateInfoCity)
        self.pushButtonAdd.clicked.connect(self.addClient)
        self.pushButtonAddCity.clicked.connect(self.insertCity)
        # page update client
        self.pushButtonOpenFolder.clicked.connect(self.openFolder)
        self.pushButtonEdit.clicked.connect(self.updateInfoClient)
        self.pushButtonClone.clicked.connect(self.cloneClientInfo)

        #test

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
                cur.execute("PRAGMA foreign_keys=ON")
                cur.execute(sql)
                con.commit()
                cur.execute("SELECT last_insert_rowid()")
                index = cur.fetchone()
                self.lastIndex = index
            else:
                cur = con.cursor()
                cur.execute("PRAGMA foreign_keys=ON")
                cur.execute(sql, param)
                con.commit()
                cur.execute("SELECT last_insert_rowid()")
                index = cur.fetchone()
                self.lastIndex = index
        if con:
            con.close()



    def getIdService(self, id):

        sqlIdService = "SELECT service from info_client where id = ?"

        return self.querySelect(sqlIdService, (id,))

    def getIdDocInfoClient(self, id_client):

        sqlIdDocInfo = "SELECT * from doc_info_client where id_client = ?"

        return self.querySelect(sqlIdDocInfo, (id_client,))

    def getIdCity(self, name):
        print(self.allInfoCity)
        for i in range(len(self.allInfoCity)):
            if self.allInfoCity[i][1] == name:
                return self.allInfoCity[i][0]

    def getWorkInfoClient(self, id_client):

        sqlIdAddressInfo = "SELECT * from work_info_client where id_client = ?"

        return self.querySelect(sqlIdAddressInfo, (id_client,))

    def getAddressInfoClient(self, id_client):

        sqlIdAddressInfo = "SELECT * from address_info_client where id_client = ?"

        return self.querySelect(sqlIdAddressInfo, (id_client,))

    def getForgottenClient(self):
        pass

    def getInfoCompleteClient(self):
        pass

    def getNameCity(self, id_city):
        for i in range(len(self.allInfoCity)):
            if id_city == self.allInfoCity[i][0]:
                return self.allInfoCity[i][1]

    def getAllClientInfo(self):
        # list of page client

        sqlInfoClient = "SELECT * from info_client"
        self.allInfoClient = self.querySelect(sqlInfoClient)

        sqlInfoAddress = "SELECT * from address_info_client"
        self.allInfoAddress = self.querySelect(sqlInfoAddress)

        sqlInfoWork = "SELECT * from work_info_client"
        self.allInfoWork = self.querySelect(sqlInfoWork)

    def getClientInfo(self, id_client):
        sqlInfoClient = "SELECT * from info_client where id=?"
        self.clientInfo = self.querySelect(sqlInfoClient, (id_client,))

    def getAllInfoCity(self):
        sqlInfoCity = "SELECT * from city"
        self.allInfoCity = self.querySelect(sqlInfoCity)
        self.allOnlyCity = []
        for index in range(len(self.allInfoCity)):
            self.allOnlyCity.append(self.allInfoCity[index][1])

        self.comboBoxCity.addItems(self.allOnlyCity)


    def insertAddressInfo(self, id_client):
        city = self.comboBoxCity.currentText()
        address = self.lineEditAddress.text().title()
        sqlAddressInfo = "INSERT INTO address_info_client(id_client,id_city,address) VALUES (?,?,?)"

        self.lastIndexAddressInfo = self.queryInsert(sqlAddressInfo, (id_client, self.getIdCity(city), address,))

    def insertDocInfo(self, id_client):

        passDocSeries = self.lineEditPassDocSeries.text()
        passDocDate = self.dateEditPassDocDate.text()
        passDocInfo = self.lineEditPassDocInfo.text()
        passSnils = self.lineEditPassSnils.text()
        sqlDocInfo = "INSERT INTO doc_info_client(id_client,series,date,info,snils) VALUES (?,?,?,?,?)"
        self.queryInsert(sqlDocInfo, (id_client, passDocSeries, passDocDate, passDocInfo, passSnils))

    def insertWorkInfo(self, id_client):

        prepayment = self.checkBoxPrepayment.checkState()
        remains = self.checkBoxRemains.checkState()
        work = self.comboBoxWork.currentIndex()
        dateWork = self.dateEditDateWork.text()
        status = self.comboBoxStatus.currentIndex()
        dateStatus = self.comboBoxStatus.currentIndex()
        info = self.textEditInfo.toPlainText()
        sqlWorkInfo = "INSERT INTO work_info_client(id_client,prepayment,remains,work,date_work,status,date_status,info) VALUES " \
                      "(?,?,?,?,?,?,?,?) "
        self.queryInsert(sqlWorkInfo, (id_client, prepayment, remains, work, dateWork, status, dateStatus, info))

    def insertInfoClient(self):

        self.createFolder()

        service = self.comboBoxProvideServices.currentIndex()
        surName = self.lineEditSurnameAdd.text()
        name = self.lineEditNameAdd.text()
        middleName = self.lineEditMiddleNameAdd.text()
        telefone = self.lineEditTelefoneAdd.text()
        sqlInfoClient = "INSERT INTO info_client(sur_name,name,middle_name,telefone,path_folder,service) VALUES (?,?," \
                        "?,?,?,?) "
        self.queryInsert(sqlInfoClient, (
            surName, name, middleName, telefone, os.path.abspath(self.pathFolderClient), int(service)))
        lastIndexCleint = self.lastIndex

        self.insertAddressInfo(lastIndexCleint[0])

        self.insertWorkInfo(lastIndexCleint[0])

        self.insertDocInfo(lastIndexCleint[0])

    def insertCity(self):
        check = False
        sqlInsertCity = "INSERT INTO city(city_name) values(?)"
        dlg = QInputDialog(self)
        dlg.setInputMode(QInputDialog.TextInput)
        dlg.setWindowTitle("Добавить населенный пункт")
        dlg.setLabelText("")
        dlg.resize(300, 100)
        ok = dlg.exec_()
        city = dlg.textValue()
        if ok:
            for i in range(len(self.allOnlyCity)):
                if self.allOnlyCity[i] == city:
                    self.msgInfo("Найден похожий населенный пункт")
                    check = False
                    break
                else:
                    check = True
        if check:
            self.queryInsert(sqlInsertCity, (city,))
            self.msgInfo("Населенный пункт добавлен")
            self.comboBoxCity.clear()
            self.getAllInfoCity()


    def updateTableClient(self):
        self.getAllClientInfo()
        self.tableWidget.setRowCount(0)
        lenInfoClient = len(self.allInfoClient)

        for i in range(lenInfoClient):
            self.tableWidget.insertRow(self.tableWidget.rowCount())
            id = str(self.allInfoClient[i][0])
            statusWorkText = self.comboBoxStatus.itemText(self.getWorkInfoClient(id)[0][4])
            addressCityText = self.getNameCity(self.allInfoAddress[i][1])
            surNameText = self.allInfoClient[i][1]
            nameText = self.allInfoClient[i][2]
            middleNameText = self.allInfoClient[i][3]
            telefoneText = self.allInfoClient[i][4]
            serviceText = self.comboBoxProvideServices.itemText(self.allInfoClient[i][6])

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

            item = QTableWidgetItem()
            item.setText(id)  # id
            self.tableWidget.setItem(i, 7, item)

    def updateInfoClient(self):
        pass

    def updateInfoCity(self):
        check = False
        index = ''
        sqlUpdateCity = "UPDATE city SET city_name = ? where id = ?"
        dlg = QInputDialog(self)
        dlg.setInputMode(QInputDialog.TextInput)
        dlg.setWindowTitle("Добавить населенный пункт")
        dlg.setLabelText("")
        dlg.setTextValue(self.comboBoxCity.currentText())

        dlg.resize(300, 100)
        ok = dlg.exec_()
        city = dlg.textValue()
        if ok:
            for i in range(len(self.allOnlyCity)):
                if self.allOnlyCity[i] == city:
                    self.msgInfo("Найден похожий населенный пункт")
                    check = False
                    break
                else:
                    index = self.comboBoxCity.currentIndex()+1
                    check = True
        if check:
            self.queryInsert(sqlUpdateCity, (city, index))
            self.msgInfo("Населенный пункт изменен")
            self.comboBoxCity.clear()
            self.getAllInfoCity()

    def createFolder(self):
        path = os.path.join(os.getcwd(),
                            self.comboBoxProvideServices.currentText(),
                            self.comboBoxCity.currentText(),
                            self.lineEditAddress.text() + " " + self.lineEditSurname.text())
        os.makedirs(path)
        os.startfile(path)
        self.pathFolderClient = path

    def openFolder(self):
        try:
            os.startfile(os.path.abspath(self.pathFolderClient))
        except FileNotFoundError as text:
            QMessageBox.critical(self, "Ошибка ", str(text), QMessageBox.Ok)

    def deleteFolder(self, path):
        shutil.rmtree(path)


    def addClient(self):
        self.insertInfoClient()
        self.clearAddClientPage()

    def deleteClient(self):
        buttonReply = QMessageBox.question(self, 'Подтверждение действия', "Удалить выбранную запись?",
                                           QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if buttonReply == QMessageBox.Yes:
            try:
                id_client = self.tableWidget.item(self.tableWidget.currentRow(),7).text()
                self.getClientInfo(int(id_client))
                clientPathFolder = self.clientInfo[0][5]
                self.deleteFolder(clientPathFolder)
                sql = "DELETE FROM info_client WHERE id=?"
                self.queryInsert(sql, (id_client,))

            except PermissionError:
                self.msgError("Удаление невозможно, закройте файлы находящиеся в папке")

    def searchClient(self):
        pass

    def cloneClientInfo(self):
        pass


    def clearAddClientPage(self):
        self.comboBoxProvideServices.setCurrentIndex(0)
        self.comboBoxCity.setCurrentIndex(0)
        self.lineEditAddress.clear()
        self.lineEditSurnameAdd.clear()
        self.lineEditNameAdd.clear()
        self.lineEditMiddleNameAdd.clear()
        self.lineEditTelefoneAdd.clear()

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
