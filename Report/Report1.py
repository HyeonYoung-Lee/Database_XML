import sys
import pymysql
import datetime
import csv
import json
import xml.etree.ElementTree as ET
from decimal import *
from PyQt5.QtWidgets import *

class DB_Utils:
    def queryExecutor(self, sql, params):
        conn = pymysql.connect(host='localhost', user='guest', password='bemyguest', db='classicmodels', charset='utf8')

        try:
            with conn.cursor(pymysql.cursors.DictCursor) as cursor:
                cursor.execute(sql, params)
                rows = cursor.fetchall()
                return rows
        except Exception as e:
            print(e)
            print(type(e))
        finally:
            conn.close()

class DB_Queries:

    def selectCustomerName(self):
        sql = "SELECT name FROM customers"
        params = ()

        util = DB_Utils()
        rows = util.queryExecutor(sql=sql, params=params)
        return rows

    def selectCustomerCountry(self):
        sql = "SELECT country FROM customers"
        params = ()

        util = DB_Utils()
        rows = util.queryExecutor(sql=sql, params=params)
        return rows

    def selectCustomerCity(self):
        sql = "SELECT city FROM customers"
        params = ()

        util = DB_Utils()
        rows = util.queryExecutor(sql=sql, params=params)
        return rows

    def selectCityFromCountry(self, country):
        sql = "SELECT city FROM customers WHERE country = %s"
        params = (country)

        util = DB_Utils()
        rows = util.queryExecutor(sql=sql, params=params)
        return rows

    def selectOrders(self, selectedCombo, comboValue):
        sql = ''

        if comboValue == 'ALL':
            sql = f"SELECT O.orderNo, O.orderDate, O.requiredDate, O.shippedDate, O.status, C.name as customer, O.comments FROM ( SELECT customerId, name FROM customers WHERE {selectedCombo} is not null) C, orders O WHERE C.customerId = O.customerId ORDER BY O.orderNo"
            params = ()

        else:
            sql = f"SELECT O.orderNo, O.orderDate, O.requiredDate, O.shippedDate, O.status, C.name as customer, O.comments FROM ( SELECT customerId, name FROM customers WHERE {selectedCombo} = %s) C, orders O WHERE C.customerId = O.customerId ORDER BY O.orderNo"
            params = (comboValue)

        util = DB_Utils()
        rows = util.queryExecutor(sql=sql, params=params)
        return rows

    def searchOrders(self, orderNo):
        sql = "SELECT O.orderLineNo, O.productCode, P.name as productName, O.quantity, O.priceEach, (O.quantity * O.priceEach) as '상품주문액' FROM (SELECT orderLineNo, quantity, priceEach, productCode FROM orderDetails WHERE orderNo = %s) O, products P where O.productCode = P.productCode ORDER BY orderLineNo"
        params = (orderNo)

        util = DB_Utils()
        rows = util.queryExecutor(sql=sql, params=params)
        return rows

class MainWindow(QWidget):
    selectedCombo = 'name'          # name, country, city
    comboValue = 'ALL'

    def __init__(self):
        super().__init__()
        self.setupUI()

    def setupUI(self):

        # DB 검색문 실행
        query = DB_Queries()

        name_rows = query.selectCustomerName()
        nameItem = []
        for name in name_rows:
            nameItem.append(name['name'])
        nameItem.sort()
        nameItem.insert(0, 'ALL')

        country_rows = query.selectCustomerCountry()
        countryItem = set()
        for country in country_rows:
            countryItem.add(country['country'])
        countryItem = list(countryItem)
        countryItem.sort()
        countryItem.insert(0, 'ALL')

        city_rows = query.selectCustomerCity()
        cityItem = set()
        for city in city_rows:
            cityItem.add(city['city'])
        cityItem = list(cityItem)
        cityItem.sort()
        cityItem.insert(0, 'ALL')

        ### Window ###
        self.setWindowTitle("주문 검색")
        self.setGeometry(0, 0, 1100, 620)

        ### Widget ###
        # Label
        customerLbl = QLabel('고객 :')
        countryLbl = QLabel('국가 :')
        cityLbl = QLabel('도시 :')
        countLbl = QLabel('검색된 주문의 개수 :')
        self.count = QLabel('')

        # combo box
        self.customerCombo = QComboBox(self)
        self.customerCombo.addItems(nameItem)
        self.customerCombo.activated.connect(self.customerComboBox_Activated)

        self.countryCombo = QComboBox(self)
        self.countryCombo.addItems(countryItem)
        self.countryCombo.activated.connect(self.countryComboBox_Activated)

        self.cityCombo = QComboBox(self)
        self.cityCombo.addItems(cityItem)
        self.cityCombo.activated.connect(self.cityComboBox_Activated)

        groupBox = QGroupBox('주문 검색')

        # Button
        searchBtn = QPushButton("검색", self)
        searchBtn.clicked.connect(self.searchBtn_Clicked)
        initBtn = QPushButton("초기화", self)
        initBtn.clicked.connect(self.initBtn_Clicked)

        # Table
        self.tableWidget = QTableWidget(100, 8)
        self.tableWidget.cellDoubleClicked.connect(self.cellDoubleClicked_event)

        ### Layout ###
        topInnerLeftLayout = QGridLayout()
        topInnerLeftLayout.addWidget(customerLbl, 0, 0)
        topInnerLeftLayout.addWidget(self.customerCombo, 0, 1)
        topInnerLeftLayout.addWidget(countryLbl, 0, 2)
        topInnerLeftLayout.addWidget(self.countryCombo, 0, 3)
        topInnerLeftLayout.addWidget(cityLbl, 0, 4)
        topInnerLeftLayout.addWidget(self.cityCombo, 0, 5)
        topInnerLeftLayout.addWidget(countLbl, 1, 0)
        topInnerLeftLayout.addWidget(self.count, 1, 1)

        topInnerRightLayout = QVBoxLayout()
        topInnerRightLayout.addWidget(searchBtn)
        topInnerRightLayout.addWidget(initBtn)

        groupBox.setLayout(topInnerLeftLayout)

        topLayout = QHBoxLayout()
        topLayout.addWidget(groupBox, 5)
        topLayout.addLayout(topInnerRightLayout, 1)

        bottomLayout = QVBoxLayout()
        bottomLayout.addWidget(self.tableWidget)

        layout = QVBoxLayout()
        layout.addLayout(topLayout)
        layout.addLayout(bottomLayout)

        self.setLayout(layout)

    def customerComboBox_Activated(self):
        self.selectedCombo = 'name'
        self.comboValue = self.customerCombo.currentText()

    def countryComboBox_Activated(self):
        self.selectedCombo = 'country'
        self.comboValue = self.countryCombo.currentText()

        # reset city for country
        query = DB_Queries()
        city_rows = query.selectCityFromCountry(self.comboValue)
        cityItem = set()
        for city in city_rows:
            cityItem.add(city['city'])
        cityItem = list(cityItem)
        cityItem.sort()
        cityItem.insert(0, 'ALL')

        self.cityCombo.clear()
        self.cityCombo.addItems(cityItem)
        self.cityCombo.activated.connect(self.cityComboBox_Activated)

    def cityComboBox_Activated(self):
        self.selectedCombo = 'city'
        self.comboValue = self.cityCombo.currentText()

    def searchBtn_Clicked(self):
        # query
        query = DB_Queries()
        orders = query.selectOrders(self.selectedCombo, self.comboValue)

        self.tableWidget.clearContents()
        self.tableWidget.setRowCount(len(orders))
        try:
            self.tableWidget.setColumnCount(len(orders[0]))
            columnNames = list(orders[0].keys())
            self.count.setText(str(len(orders)))

        except IndexError as e:
            self.tableWidget.setColumnCount(0)
            columnNames = []
            self.count.setText('0')

        self.tableWidget.setHorizontalHeaderLabels(columnNames)
        self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)

        for rowIdx, orders in enumerate(orders):
            for columnIdx, (k, v) in enumerate(orders.items()):
                if v == None:
                    continue
                elif isinstance(v, datetime.date):
                    item = QTableWidgetItem(v.strftime('%Y-%m-%d'))
                else:
                    item = QTableWidgetItem(str(v))

                self.tableWidget.setItem(rowIdx, columnIdx, item)

        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.resizeRowsToContents()

    def initBtn_Clicked(self):
        # reset city
        query = DB_Queries()
        city_rows = query.selectCustomerCity()
        cityItem = set()
        for city in city_rows:
            cityItem.add(city['city'])
        cityItem = list(cityItem)
        cityItem.sort()
        cityItem.insert(0, 'ALL')

        self.cityCombo.clear()
        self.cityCombo.addItems(cityItem)

        self.customerCombo.setCurrentIndex(0)
        self.countryCombo.setCurrentIndex(0)
        self.cityCombo.setCurrentIndex(0)
        self.tableWidget.clearContents()

        self.selectedCombo = 'customer'
        self.comboValue = 'ALL'
        self.count.setText('')

    def cellDoubleClicked_event(self, row, col):
        data = self.tableWidget.item(row, 0)
        orderNo = data.text()
        self.subwindow = SubWindow(orderNo)
        self.subwindow.show()


class SubWindow(QWidget):
    orderNo = 0
    sumPrice = 0
    saveMsg = ''
    def __init__(self, orderNo):
        super().__init__()
        self.orderNo = orderNo
        self.setupUI()

    def setupUI(self):
        # DB 검색문 실행
        query = DB_Queries()
        self.order_rows = query.searchOrders(orderNo=self.orderNo)

        ### Window ###
        self.setWindowTitle("주문 상세")
        self.setGeometry(0, 0, 1000, 520)

        ### Widget ###
        # Label
        orderNumLbl = QLabel('주문번호 :')
        self.orderNum = QLabel('')
        quantityLbl = QLabel('상품개수 :')
        self.quantity = QLabel('')
        priceLbl = QLabel('주문액 :')
        self.price = QLabel('')

        # button
        self.raidioBtnCSV = QRadioButton('CSV', self)
        self.raidioBtnCSV.setChecked(True)
        self.saveMsg = 'csv'
        self.raidioBtnCSV.clicked.connect(self.radioBtn_Clicked)
        self.raidioBtnJSON = QRadioButton('JSON', self)
        self.raidioBtnJSON.clicked.connect(self.radioBtn_Clicked)
        self.raidioBtnXML = QRadioButton('XML', self)
        self.raidioBtnXML.clicked.connect(self.radioBtn_Clicked)

        self.saveBtn = QPushButton('저장', self)
        self.saveBtn.clicked.connect(self.saveBtn_Clicked)

        # groupbox
        topGroupBox = QGroupBox('주문 상세 내역')
        bottomGroupBox = QGroupBox('파일 출력')

        # Table
        self.tableWidget = QTableWidget(100, 8)

        ### Layout ###
        topInnerLayout = QHBoxLayout()
        topInnerLayout.addWidget(orderNumLbl)
        topInnerLayout.addWidget(self.orderNum)
        topInnerLayout.addWidget(quantityLbl)
        topInnerLayout.addWidget(self.quantity)
        topInnerLayout.addWidget(priceLbl)
        topInnerLayout.addWidget(self.price)

        topGroupBox.setLayout(topInnerLayout)

        bottomInnerLayout = QHBoxLayout()
        bottomInnerLayout.addWidget(self.raidioBtnCSV)
        bottomInnerLayout.addWidget(self.raidioBtnJSON)
        bottomInnerLayout.addWidget(self.raidioBtnXML)
        bottomInnerLayout.addWidget(self.saveBtn)

        bottomGroupBox.setLayout(bottomInnerLayout)

        layout = QVBoxLayout()
        layout.addWidget(topGroupBox)
        layout.addWidget(self.tableWidget)
        layout.addWidget(bottomGroupBox)

        self.setLayout(layout)

        ### set values ###
        # table
        self.tableWidget.clearContents()
        self.tableWidget.setRowCount(len(self.order_rows))
        try:
            self.tableWidget.setColumnCount(len(self.order_rows[0]))
            columnNames = list(self.order_rows[0].keys())
            self.quantity.setText(str(len(self.order_rows)))

        except IndexError as e:
            self.tableWidget.setColumnCount(0)
            columnNames = []
            self.quantity.setText('0')

        self.tableWidget.setHorizontalHeaderLabels(columnNames)
        self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)

        for rowIdx, orders in enumerate(self.order_rows):
            for columnIdx, (k, v) in enumerate(orders.items()):
                if k == '상품주문액' and v != None:
                    self.sumPrice += v

                if v == None:
                    continue
                else:
                    item = QTableWidgetItem(str(v))

                self.tableWidget.setItem(rowIdx, columnIdx, item)

        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.resizeRowsToContents()

        self.orderNum.setText(str(self.orderNo))
        self.price.setText(str(self.sumPrice))

    def radioBtn_Clicked(self):
        if self.raidioBtnCSV.isChecked():
            self.saveMsg = 'csv'
        elif self.raidioBtnJSON.isChecked():
            self.saveMsg = 'json'
        else:
            self.saveMsg = 'xml'

    def saveBtn_Clicked(self):
        if self.saveMsg == 'csv':
            self.save_CSV()

        elif self.saveMsg == 'json':
            self.save_JSON()

        elif self.saveMsg == 'xml':
            self.save_XML()

    def save_CSV(self):
        with open(f'./{self.orderNo}.csv', 'w', encoding='utf-8', newline='') as f:
            wr = csv.writer(f)

            columnNames = list(self.order_rows[0].keys())
            wr.writerow(columnNames)

            for row in self.order_rows:
                item = list(row.values())
                wr.writerow(item)

    def save_JSON(self):
        for order in self.order_rows:
            for k, v in order.items():
                if isinstance(v, Decimal):
                    order[k] = str(v)

        with open(f'./{self.orderNo}.json', 'w', encoding='utf-8') as f:
            json.dump(self.order_rows, f, ensure_ascii=False)

    def save_XML(self):
        rootElement = ET.Element('TABLE')

        for row in self.order_rows:
            rowElement = ET.Element('ROW')
            rootElement.append(rowElement)

            for columnName in list(row.keys()):
                item = row[columnName]
                if item == None:
                    rowElement.attrib[columnName] = ''
                elif type(item) == int or type(item) == Decimal:
                    rowElement.attrib[columnName] = str(item)
                else:
                    rowElement.attrib[columnName] = item

        ET.ElementTree(rootElement).write(f'./{self.orderNo}.xml', encoding='utf-8', xml_declaration=True)

def main():
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec_())

main()