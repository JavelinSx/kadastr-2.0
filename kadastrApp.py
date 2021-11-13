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
        os.chdir("D:/testDir")
        # var
        self.infoClientForSelectCalendar = None
        self.infoAddressForSelectCalendar = None
        self.infoWorkForSelectCalendar = None
        self.idClientForSelectcalendar = None
        self.workForCalendar = None
        self.statusForCalendar = None
        self.dateStatusForCalendar = None
        self.createPath = None
        self.infoClient = None
        self.infoPassportClient = None
        self.infoWorkClient = None
        self.indexClient = None
        self.newCity = None
        self.pathFolderNew = None
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
        self.pushButtonOpenFolder.clicked.connect(lambda: self.openFolder(self.infoClient[0][5]))
        self.pushButtonEdit.clicked.connect(self.updateFullInfoForClient)
        self.tableWidget.doubleClicked.connect(self.fillClientInfo)
        # test
        self.updateTableClient()

<<<<<<< HEAD
=======
        #test
>>>>>>> 64a251773191da07d6f1401b7188d549653d2c0a

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

    def createDateObject(self, date):
        dateBuild = QDate()
        year = int(date[6:10])
        mnt = int(date[3:5])
        day = int(date[0:2])
        dateBuild.setDate(year, mnt, day)
        return dateBuild

    #

    def getInfoForSelectCalendar(self, id_client): # Доделать
        sqlInfo = ""
        sqlAddress = ""
        sqlWork = ""
        self.infoClientForSelectCalendar = self.querySelect(sqlInfo, id_client)
        self.infoAddressForSelectCalendar = self.querySelect(sqlAddress, id_client)
        self.infoWorkForSelectCalendar = self.querySelect(sqlWork, id_client)

    def getIdClientForSelectCalendar(self, date): # Доделать
        sqlDate = ""
        self.idClientForSelectcalendar = self.querySelect(sqlDate, date)

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

        return self.querySelect(sqlInfoClient, (id_client,))

    def getAllInfoCity(self):
        self.comboBoxCity.clear()
        sqlInfoCity = "SELECT * from city"
        self.allInfoCity = self.querySelect(sqlInfoCity)
        self.allOnlyCity = []
        for index in range(len(self.allInfoCity)):
            self.allOnlyCity.append(self.allInfoCity[index][1])

        self.comboBoxCity.addItems(self.allOnlyCity)
<<<<<<< HEAD

    def getDateStatusForCalendar(self):
        sqlDateStatus = "SELECT date_status from work_info_client"
        self.dateStatusForCalendar = self.querySelect(sqlDateStatus)
=======
>>>>>>> 64a251773191da07d6f1401b7188d549653d2c0a

    def getWorkForCalendar(self, date):
        sqlWorkInfo = "SELECT work from work_info_client where date_status = ?"
        self.workForCalendar = self.querySelect(sqlWorkInfo, date)

    def getStatusForCalendar(self):
        sqlStatusInfo = "SELECT status from work_info_client"
        self.statusForCalendar = self.querySelect(sqlStatusInfo)

    #

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
        work = 0
        dateWork = self.dateEditDateWork.text()
        status = 1
        dateStatus = self.dateEditReception.text()
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
            surName, name, middleName, telefone, os.path.abspath(self.createPath), int(service)))
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

    #

    def updateTableClient(self):
        self.calendarWork()
        self.getAllClientInfo()
        self.tableWidget.setRowCount(0)
        lenInfoClient = len(self.allInfoClient)

        for i in range(lenInfoClient):
            self.tableWidget.insertRow(self.tableWidget.rowCount())
            id = str(self.allInfoClient[i][0])
            status = self.getWorkInfoClient(id)
            print(status)
            statusWorkText = self.comboBoxWork.itemText(status[0][4])
            addressCityText = self.comboBoxCity.itemText(self.allInfoAddress[i][2])

            surNameText = self.allInfoClient[i][1]
            nameText = self.allInfoClient[i][2]
            middleNameText = self.allInfoClient[i][3]
            telefoneText = self.allInfoClient[i][4]
            serviceText = self.comboBoxProvideServices.itemText(int(self.allInfoClient[i][6]))

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

    def updateInfoClient(self, sur_name, name, middle_name, telefone, path_folder, service, id_client):
        sqlUpdateInfoClient = "UPDATE info_client SET sur_name = ?, name = ?, middle_name = ?, telefone = ?, path_folder = ?, service = ? where id = ?"
        self.queryInsert(sqlUpdateInfoClient, (sur_name, name, middle_name, telefone, path_folder, service, id_client))

    def updateAddressClient(self, id_city, address, id_client):

        sqlUpdateInfoAddress = "UPDATE address_info_client SET id_city = ?, address = ? where id_client = ?"
        self.queryInsert(sqlUpdateInfoAddress, (id_city, address, id_client))

    def updateDocClient(self, series, date, info, snils, id_client):
        sqlUpdateInfoDoc = "UPDATE doc_info_client SET series = ?, date = ?, info = ?, snils = ? where id_client = ?"
        self.queryInsert(sqlUpdateInfoDoc, (series, date, info, snils, id_client))

    def updateWorkClient(self, prepayment, remains, work, date_work, status, date_status, info, id_client):
        sqlUpdateInfoWork = "UPDATE work_info_client SET prepayment = ?, remains = ?, work = ?, date_work = ?, status = ?, date_status = ?, info = ? where id_client = ? "
        self.queryInsert(sqlUpdateInfoWork,
                         (prepayment, remains, work, date_work, status, date_status, info, id_client))

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
                    index = self.comboBoxCity.currentIndex() + 1
                    check = True
        if check:
            self.queryInsert(sqlUpdateCity, (city, index))
            self.msgInfo("Населенный пункт изменен")
            self.comboBoxCity.clear()
            self.getAllInfoCity()

    def updateFullInfoForClient(self):
        self.changeFolder()
        id_city = self.comboBoxCityChange.currentIndex()
        address = self.lineEditAddressChange.text()
        self.updateAddressClient(id_city, address, self.indexClient)

        sur_name = self.lineEditSurname.text()
        name = self.lineEditName.text()
        middle_name = self.lineEditMiddleName.text()
        telefone = self.lineEditTelefone.text()
        path_folder = self.pathFolderNew
        service = self.comboBoxService.currentIndex()
        self.updateInfoClient(sur_name, name, middle_name, telefone, path_folder, service, self.indexClient)

        series = self.lineEditPassDocSeries.text()
        datePass = self.dateEditPassDocDate.text()
        infoPass = self.lineEditPassDocInfo.text()
        snils = self.lineEditPassSnils.text()

        self.updateDocClient(series, datePass, infoPass, snils, self.indexClient)

<<<<<<< HEAD
        prepayment = self.checkBoxPrepayment.checkState()
        remains = self.checkBoxPrepayment.checkState()
        work = self.comboBoxWork.currentIndex()
        date_work = self.dateEditDateWork.text()
        status = self.comboBoxStatus.currentIndex()
        date_status = self.dateEditReception.text()
        infoClient = self.textEditInfo.toPlainText()
        self.updateWorkClient(prepayment, remains, work, date_work, status, date_status, infoClient, self.indexClient)

    #

    def fillClientInfo(self):
        # info data
        self.getAllInfoCity()
        self.clearChangePage()
        self.comboBoxCityChange.addItems(self.allOnlyCity)
        self.tabWidget.setCurrentIndex(2)
        row = self.tableWidget.currentRow()
        self.indexClient = self.tableWidget.item(row, 7).text()

        self.infoClient = self.getClientInfo(self.indexClient)
        self.infoPassportClient = self.getIdDocInfoClient(self.indexClient)
        self.infoWorkClient = self.getWorkInfoClient(self.indexClient)
        infoAddressClient = self.getAddressInfoClient(self.indexClient)
        # info city
        city = infoAddressClient[0][2]
        address = infoAddressClient[0][3]

        # info client

        surName = self.infoClient[0][1]
        name = self.infoClient[0][2]
        middleName = self.infoClient[0][3]
        telefone = self.infoClient[0][4]
        path = self.infoClient[0][5]
        service = self.infoClient[0][6]
        # passport info

        series = self.infoPassportClient[0][2]
        datePassport = self.createDateObject(self.infoPassportClient[0][3])
        infoPassport = self.infoPassportClient[0][4]
        snils = self.infoPassportClient[0][5]

        # work info

        prepaymant = self.infoWorkClient[0][2]
        paymant = self.infoWorkClient[0][3]
        wokrStatus = self.infoWorkClient[0][4]
        dateWorkStatus = self.createDateObject(self.infoWorkClient[0][5])
        status = self.infoWorkClient[0][6]
        dateStatus = self.createDateObject(self.infoWorkClient[0][7])
        info = self.infoWorkClient[0][8]

        self.comboBoxService.setCurrentIndex(int(service))
        self.comboBoxCityChange.setCurrentIndex(city)
        self.lineEditAddressChange.setText(address)
        self.lineEditPathChange.setText(path)

        self.lineEditSurname.setText(surName)
        self.lineEditName.setText(name)
        self.lineEditMiddleName.setText(middleName)
        self.lineEditTelefone.setText(telefone)

        self.lineEditPassDocSeries.setText(series)
        self.dateEditPassDocDate.setDate(datePassport)
        self.lineEditPassDocInfo.setText(infoPassport)
        self.lineEditPassSnils.setText(snils)

        self.checkBoxPrepayment.setCheckState(int(prepaymant))
        self.checkBoxRemains.setCheckState(int(paymant))
        self.comboBoxWork.setCurrentIndex(wokrStatus)
        self.dateEditDateWork.setDate(dateWorkStatus)
        self.comboBoxStatus.setCurrentIndex(status)
        self.dateEditReception.setDate(dateStatus)
        self.textEditInfo.setText(info)

    #
=======
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

    def fillClientInfo(self):
        #info data

        infoClient = self.getClientInfo(int(self.tableWidget.item(self.tableWidget.currentRow(), 7).text()))
        infoPassportClient = []
        infoWorkClient = []

        #info client

        surName = ''
        name = ''
        middleName = ''
        telefone = ''

        #passport info

        series = ''
        datePassport = ''
        infoPassport = ''
        snils = ''

        #work info

        paymant = ''
        prepaymant = ''
        wokrStatus = ''
        dateWorkStatus = ''
        status = ''
        dateStatus = ''
        info = ''

        pass
>>>>>>> 64a251773191da07d6f1401b7188d549653d2c0a

    def createFolder(self):
        path = os.path.join(os.getcwd(),
                            self.comboBoxProvideServices.currentText(),
                            self.comboBoxCity.currentText(),
                            self.lineEditAddress.text() + " " + self.lineEditSurname.text())
        os.makedirs(path)
        self.createPath = path

    def openFolder(self, path):

        os.startfile(os.path.abspath(path))

    def deleteFolder(self, path):
        shutil.rmtree(path)

    def changeFolder(self):
        path = os.path.join(os.getcwd(),
                            self.comboBoxService.currentText(),
                            self.comboBoxCityChange.currentText(),
                            self.lineEditAddressChange.text() + " " + self.lineEditSurname.text())
        pathOld = self.infoClient[0][5]
        print(self.infoClient)
        print(path)
        print(pathOld)
        if path != pathOld:
            os.makedirs(path)
            os.startfile(path)
            self.pathFolderNew = path
            for item in os.listdir(pathOld):
                s = os.path.join(os.path.join(pathOld, item))
                d = os.path.join(path, item)
                if os.path.isdir(s):
                    shutil.copytree(s, d)
                else:
                    shutil.copy2(s, d)
            buttonReply = QMessageBox.question(self, 'Подтверждение действия', "Удалить выбранную запись?",
                                               QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if buttonReply == QMessageBox.Yes:
                self.deleteFolder(pathOld)
        else:
            self.pathFolderNew = self.infoClient[0][5]

    #

    def addClient(self):
        self.insertInfoClient()
        self.clearAddClientPage()

    def deleteClient(self):
        buttonReply = QMessageBox.question(self, 'Подтверждение действия', "Удалить выбранную запись?",
                                           QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if buttonReply == QMessageBox.Yes:
            try:
<<<<<<< HEAD
                id_client = self.tableWidget.item(self.tableWidget.currentRow(), 7).text()
=======
                id_client = self.tableWidget.item(self.tableWidget.currentRow(),7).text()
>>>>>>> 64a251773191da07d6f1401b7188d549653d2c0a

                clientPathFolder = self.getClientInfo(int(id_client))[0][5]
                self.deleteFolder(clientPathFolder)
                sql = "DELETE FROM info_client WHERE id=?"
                self.queryInsert(sql, (id_client,))

            except PermissionError:
                self.msgError("Удаление невозможно, закройте файлы находящиеся в папке")

    def searchClient(self):
        pass


    #

    def clearAddClientPage(self):
        self.comboBoxProvideServices.setCurrentIndex(0)
        self.comboBoxCity.setCurrentIndex(0)
        self.lineEditAddress.clear()
        self.lineEditSurnameAdd.clear()
        self.lineEditNameAdd.clear()
        self.lineEditMiddleNameAdd.clear()
        self.lineEditTelefoneAdd.clear()

    def clearChangePage(self):

        self.comboBoxCityChange.clear()
        self.lineEditSurname.clear()
        self.lineEditName.clear()
        self.lineEditMiddleName.clear()
        self.lineEditTelefone.clear()

        self.lineEditPassDocSeries.clear()
        self.dateEditPassDocDate.clear()
        self.lineEditPassDocInfo.clear()
        self.lineEditPassSnils.clear()

        self.dateEditDateWork.clear()

        self.dateEditReception.clear()
        self.textEditInfo.clear()

        self.checkBoxRemains.setCheckState(0)
        self.checkBoxPrepayment.setCheckState(0)
        self.comboBoxService.setCurrentIndex(0)

    #

    def calendarWork(self):
        dateBuild = QDate()
        color = QColor()
        brush = QBrush()
        form = QTextCharFormat()
        self.getDateStatusForCalendar()

        for date in self.dateStatusForCalendar:

            self.getWorkForCalendar(date)
            print(self.workForCalendar)
            if self.workForCalendar[0][0] == 1:
                print("hello1")
                color.setRgb(222, 245, 140)
            elif self.workForCalendar[0][0] == 0:
                print("hello2")
                color.setRgb(237, 140, 28)


            year = int(str(date[0])[6:10])
            mnt = int(str(date[0])[3:5])
            day = int(str(date[0])[0:2])
            dateBuild.setDate(year, mnt, day)
            brush.setColor(color)
            form.setBackground(brush)
            self.calendarWidget.setDateTextFormat(dateBuild, form)

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
