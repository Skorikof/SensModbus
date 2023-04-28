from PyQt5.QtWidgets import QHeaderView, QFrame, QMainWindow, QWidget, QGridLayout,\
    QTableWidget, QTableWidgetItem, QInputDialog
from PyQt5.QtCore import QObject, pyqtSignal, Qt
from PyQt5.QtGui import QFont, QIcon


class MySignals(QObject):
    win_closed = pyqtSignal()
    query_saveValidNum = pyqtSignal(int, int, int)


class TableValidNumbers(QMainWindow):
    def __init__(self, logger):
        super(TableValidNumbers, self).__init__()
        self.logger = logger
        self.signals = MySignals()
        self.InitUi()

    def InitUi(self):
        try:
            self.setWindowIcon(QIcon('ico/settings.png'))
            self.setWindowTitle('Таблица валидных номеров')
            self.resize(300, 400)
            self.setWindowFlag(Qt.WindowStaysOnTopHint)
            self.setWindowModality(Qt.WindowModal)
            self.centralwidget = QWidget()
            self.centralwidget.setObjectName(u"centralwidget")
            self.gridLayout = QGridLayout(self.centralwidget)
            self.gridLayout.setObjectName(u"gridLayout")
            self.tableNumbers = QTableWidget()
            self.tableNumbers.setObjectName(u"tableNumbers")
            self.tableNumbers.setColumnCount(2)
            self.tableNumbers.setRowCount(30)
            self.tableNumbers.setHorizontalHeaderLabels(['Slave ID', 'Серийный номер'])
            temp_lst = []
            for i in range(31):
                temp_lst.append('')
            self.tableNumbers.setVerticalHeaderLabels(temp_lst)
            self.tableNumbers.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            font = QFont('Calibri', 12)
            self.tableNumbers.setFont(font)
            font = QFont('Calibri', 10, QFont.Bold)
            self.tableNumbers.horizontalHeader().setFont(font)
            header = self.tableNumbers.horizontalHeader()
            header.setFrameStyle(QFrame.Box | QFrame.Plain)
            header.setLineWidth(1)
            self.tableNumbers.setHorizontalHeader(header)
            self.gridLayout.addWidget(self.tableNumbers)
            self.tableNumbers.cellClicked.connect(self.cellClick)
            self.setCentralWidget(self.centralwidget)
            self.setVisible(True)

        except Exception as e:
            self.logger.error(str(e))

    def fillTable(self, number_values):
        try:
            for i in range(len(number_values)):
                item = QTableWidgetItem(str(i + 2))
                self.tableNumbers.setItem(i, 0, item)
                item = QTableWidgetItem(str(number_values[i]))
                self.tableNumbers.setItem(i, 1, item)
            print('заполняем таблицы: {}'.format(number_values))
            pass

        except Exception as e:
            self.logger.error(str(e))

    def cellClick(self, r, c):
        try:
            item = self.tableNumbers.item(r, 0)
            id_temp = int(item.text())
            addr_temp = id_temp + 4128
            item = self.tableNumbers.item(r, 1)
            val_temp = int(item.text())
            number, ok = QInputDialog.getInt(self, 'Данные', 'Введите серийный номер для SlaveID = ' + str(id_temp),
                                             value=int(val_temp), min=0, max=65535, step=1)
            if ok:
                self.signals.query_saveValidNum.emit(id_temp, addr_temp, number)
        except Exception as e:
            self.logger.error(str(e))
