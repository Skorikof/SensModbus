import sys
from PyQt5.QtWidgets import QHeaderView, QFrame, QFileDialog, QProgressDialog,\
    QMainWindow, QApplication, QWidget, QGridLayout, QLabel,\
    QTableWidget, QTableWidgetItem, QStatusBar, QAction, QMenuBar, QMenu

from PyQt5.QtCore import QObject, pyqtSignal, Qt, QRect
from PyQt5.QtGui import QFont, QIcon
from functools import partial
import struct


class MySignals(QObject):
    win_closed = pyqtSignal()
    change_Sens_ID = pyqtSignal(int)
    query_save_file = pyqtSignal(int, object)
    cancel_Save_File = pyqtSignal()
    query_clear_mem = pyqtSignal()
    query_null_sens = pyqtSignal()
    clear_table = pyqtSignal()
    query_serial_number = pyqtSignal()
    query_TimeTx = pyqtSignal()
    query_ChannelNumber = pyqtSignal()
    query_TransPower = pyqtSignal()
    query_AddrModulation = pyqtSignal()
    query_FreqOffset = pyqtSignal()


class WindowSensor(QMainWindow):
    def __init__(self, logger):
        super(WindowSensor, self).__init__() 
        self.signals = MySignals()
        self.logger = logger
        self.InitUi()

    def InitUi(self):
        try:
            self.ID_sens = 0
            self.type_sensor = 'none'
            self.list_id = []
            self.setWindowIcon(QIcon('ico/settings.png'))
            str_title = 'Ct'
            if self.type_sensor == 'force':
                str_title = 'Датчик усилия.'
            if self.type_sensor == 'temper':
                str_title = 'Датчик температуры.'

            self.setWindowTitle('Датчик')
            self.centralwidget = QWidget()
            self.centralwidget.setObjectName(u"centralwidget")
            self.gridLayout = QGridLayout(self.centralwidget)
            self.gridLayout.setObjectName(u"gridLayout")
            self.labelTxt1 = QLabel(self.centralwidget)
            self.labelTxt1.setObjectName(u"labelTxt1")
            font = QFont()
            font.setPointSize(12)
            self.labelTxt1.setFont(font)
            self.labelTxt1.setStyleSheet(u"color:rgb(0, 0, 127)")
            self.labelTxt1.setAlignment(Qt.AlignCenter)
            self.labelTxt1.setText('Текущие данные')
            self.gridLayout.addWidget(self.labelTxt1, 0, 0, 1, 1)
            self.tableCurrent = QTableWidget()
            self.tableCurrent.setObjectName(u"tableCurrent")
            self.tableCurrent.setRowCount(3)
            self.tableCurrent.setColumnCount(2)
            self.tableCurrent.setHorizontalHeaderLabels(['Значение', 'В единицах АЦП'])
            self.tableCurrent.setVerticalHeaderLabels(['Усилие,кгс/см2',
                                                       'Температура воздуха,⁰С', 'Температура датчика,⁰С'])
            self.tableCurrent.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            font = QFont('Calibri', 14)
            self.tableCurrent.setFont(font)
            font = QFont('Calibri', 11, QFont.Bold)
            self.tableCurrent.horizontalHeader().setFont(font)
            header = self.tableCurrent.horizontalHeader()
            header.setFrameStyle(QFrame.Box | QFrame.Plain)
            header.setLineWidth(1)
            self.tableCurrent.setHorizontalHeader(header)
            self.gridLayout.addWidget(self.tableCurrent, 1, 0, 1, 1)
            self.tableKoeff = QTableWidget()
            self.tableKoeff.setObjectName(u"tableCoeff")
            self.tableKoeff.setRowCount(3)
            self.tableKoeff.setColumnCount(1)
            self.tableKoeff.setHorizontalHeaderLabels(['В единицах АЦП'])
            self.tableKoeff.setVerticalHeaderLabels(['Смещение нуля', 'Вес бита', 'Коэффициент усиления'])
            self.tableKoeff.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            font = QFont('Calibri', 14)
            self.tableKoeff.setFont(font)
            font = QFont('Calibri', 10, QFont.Bold)
            self.tableKoeff.horizontalHeader().setFont(font)
            header = self.tableKoeff.horizontalHeader()
            header.setFrameStyle(QFrame.Box | QFrame.Plain)
            header.setLineWidth(1)
            self.tableKoeff.setHorizontalHeader(header)
            self.gridLayout.addWidget(self.tableKoeff, 3, 0, 1, 1)
            self.labelTxt2 = QLabel(self.centralwidget)
            self.labelTxt2.setObjectName(u"labelTxt2")
            font = QFont()
            font.setPointSize(12)
            self.labelTxt2.setFont(font)
            self.labelTxt2.setStyleSheet(u"color:rgb(0, 0, 127)")
            self.labelTxt2.setAlignment(Qt.AlignCenter)
            self.labelTxt2.setText('Значения коэффициентов')
            self.gridLayout.addWidget(self.labelTxt2, 2, 0, 1, 1)
            self.statusbar = QStatusBar()
            self.statusbar.setObjectName("statusbar")
            self.setStatusBar(self.statusbar)
            self.setCentralWidget(self.centralwidget)
            # делаем окно поверх всех
            self.setWindowFlag(Qt.WindowStaysOnTopHint)
            self.resize(500, 400)
            self.createMenuSens()

            self.menuSerialNumber = QAction('Изменить серийный номер', self)
            self.menuSerialNumber.triggered.connect(self.changeSerialNumber)
            self.menuSens.addAction(self.menuSerialNumber)

            self.menuTimeTransmission = QAction('Время между передачей', self)
            self.menuTimeTransmission.triggered.connect(self.editTimeTx)
            self.menuSens.addAction(self.menuTimeTransmission)

            self.menuNumberChannel = QAction('Номер канала', self)
            self.menuNumberChannel.triggered.connect(self.changeChannelNumber)
            self.menuSens.addAction(self.menuNumberChannel)

            self.menuTransPower = QAction('Мощность передатчика', self)
            self.menuTransPower.triggered.connect(self.changeTransPower)
            self.menuSens.addAction(self.menuTransPower)

            self.menuAddrModulation = QAction('Адрес модуляции', self)
            self.menuAddrModulation.triggered.connect(self.changeAddrModulation)
            self.menuSens.addAction(self.menuAddrModulation)

            self.menuFreqOffset = QAction('Корректировка по частоте', self)
            self.menuFreqOffset.triggered.connect(self.changeFreqOffset)
            self.menuSens.addAction(self.menuFreqOffset)

            self.menuSaveFile = QAction('Записать файл калибровок', self)
            self.menuSaveFile.triggered.connect(self.saveFileToDevice)
            self.menuSens.addAction(self.menuSaveFile)

            self.menuClearMemory = QAction('Стереть страницу памяти', self)
            self.menuClearMemory.triggered.connect(self.clearMemory)
            self.menuSens.addAction(self.menuClearMemory)

            self.menuNull = QAction('Обнулить датчик', self)
            self.menuNull.triggered.connect(self.nullSens)
            self.menuSens.addAction(self.menuNull)
            self.progress_dialog = QProgressDialog()
            self.progress_dialog.setMinimum(0)
            self.progress_dialog.setMaximum(100)
            self.progress_dialog.setValue(0)
            self.progress_dialog.setWindowModality(Qt.WindowModal)
            self.progress_dialog.setCancelButtonText("Отмена")
            self.progress_dialog.canceled.connect(self.cancelSaveFile)
            self.progress_dialog.setWindowFlag(Qt.WindowStaysOnTopHint)
            self.progress_dialog.hide()
            self.setVisible(False)

        except Exception as e:
            self.logger.error(str(e))

    def editTimeTx(self):
        self.signals.query_TimeTx.emit()
    
    def changeChannelNumber(self):
        self.signals.query_ChannelNumber.emit()

    def changeSerialNumber(self):
        self.signals.query_serial_number.emit()

    def changeTransPower(self):
        self.signals.query_TransPower.emit()

    def changeAddrModulation(self):
        self.signals.query_AddrModulation.emit()

    def changeFreqOffset(self):
        self.signals.query_FreqOffset.emit()

    def nullSens(self):        
        self.signals.query_null_sens.emit()
        
    def createMenuSens(self):
        try:
            self.menubar = QMenuBar(self)
            self.menubar.setGeometry(QRect(0, 0, 468, 21))
            self.menubar.setObjectName("menubar")
            self.setMenuBar(self.menubar)
            self.menuSens = QMenu('&Настройки', self)
            self.menuNumSens = self.menuSens.addMenu('&Изменить номер')
            self.menubar.addMenu(self.menuSens)
            actions = []
            for i in range(0,32):
                action = QAction(str(i), self)
                action.setEnabled(True)
                for id in self.list_id:
                    if i == id:
                        action.setEnabled(False)
                action.triggered.connect(partial(self.changeIDSens, action))
                actions.append(action)
            self.menuNumSens.addActions(actions)

        except Exception as e:
            self.logger.error(str(e))

    def initWindow(self, ID_sens, type_sens, list_id, number_sens):
        try:
            self.ID_sens = ID_sens
            self.type_sensor = type_sens
            self.list_id = list_id
            self.serial_number_sens = number_sens
            str_title = ''
            b = True
            if self.type_sensor == 'force':
                str_title = 'Датчик усилия (Serial № : {})'.format(self.serial_number_sens)
            if self.type_sensor == 'temper':
                str_title = 'Датчик температуры (Serial № : {})'.format(self.serial_number_sens)
                b = False

            self.menuSaveFile.setEnabled(b)
            self.menuClearMemory.setEnabled(b)
            self.menuNull.setEnabled(b)

            actions = []
            for i in range(0, 32):
                action = QAction(str(i), self)
                action.setEnabled(True)
                for id in self.list_id:
                    if i == id:
                        action.setEnabled(False)
                action.triggered.connect(partial(self.changeIDSens, action))
                actions.append(action)
            self.menuNumSens.clear()
            self.menuNumSens.addActions(actions)
            self.setWindowTitle('{} Slave ID={}'.format(str_title, ID_sens))
            self.progress_dialog.hide()
            self.progress_dialog.close()

        except Exception as e:
            self.logger.error(str(e))

    def clearTables(self):
        self.signals.clear_table.emit()
        self.setVisible(True)
        try:
            for i in range(3):
                item = QTableWidgetItem('')
                item.setFlags(Qt.ItemIsEnabled)
                self.tableCurrent.setItem(i, 0, item)
                self.tableCurrent.setItem(i, 1, item)
                self.tableKoeff.setItem(i, 0, item)
            pass
        except Exception as e:
            print('{}'.format(e))

    def saveFileToDevice(self):
        try:
            fname = QFileDialog.getOpenFileName(self, 'Open file', '*.klb')[0]
            if len(fname) > 0:
                self.setfile = SettingsSens(self.logger)

                temp_a = str(fname).split('/')
                temp_f = str(temp_a[len(temp_a)-1])
                self.progress_dialog.setWindowTitle("{}".format(temp_f))
                self.progress_dialog.setLabelText("Запись регистров")
                self.progress_dialog.setValue(0)
                self.progress_dialog.show()
                b_array = self.setfile.readFileKlb(fname)
                self.signals.query_save_file.emit(self.ID_sens, b_array)

        except Exception as e:
            self.logger.error(str(e))
    
    def clearMemory(self):
        self.signals.query_clear_mem.emit()

    def cancelSaveFile(self):
        self.progress_dialog.close()
        self.signals.cancel_Save_File.emit()

    def changeIDSens(self,act):
        num=int(act.text())
        print('Запрос на изменение номера: ' + str(num))
        self.signals.change_Sens_ID.emit(num)

    def closeEvent(self, event):
        self.signals.win_closed.emit()
        self.setVisible(False)


class SettingsPull(object):
    def __init__(self):
        self.value_pull = 0
        self.value_pull_acp = 0


class Pull(object):
    def __init__(self):
        self.value_temper = 0
        self.value_temper_acp = 0
        self.pulls = []


class SettingsSens(object):
    def __init__(self, logger):
        self.logger = logger
        self.number_sens = 0
        self.data = []

    def readFileKlb(self, nam_f):
        try:
            flag_find_number = False
            flag_data_for_temper = False
            flag_data_for_pull = False
            index_temper = -1
            index_p = -1
            with open(nam_f, 'r') as fread:
                # считываем все строки
                lines = fread.readlines()
                # итерация по строкам
                for line in lines:
                    if not flag_find_number:
                        n = line.find('***')
                        if n >= 0:
                            temp_str = line.split('***')
                            temp_str1 = temp_str[1].split('.')
                            self.number_sens = int(temp_str1[0])
                            flag_find_number = True
                    else:
                        flag_info_txt = False
                        n = line.find('Темпер')
                        if n >= 0:
                            flag_info_txt = True
                            flag_data_for_temper = True
                            flag_data_for_pull = False
                        n = line.find('Усил')
                        if n >= 0:
                            flag_info_txt = True
                            flag_data_for_temper = False
                            flag_data_for_pull = True
                            index_p += 1
                            index_pull =- 1
                        if not flag_info_txt:
                            if flag_data_for_temper:
                                self.data.append(Pull())
                                index_temper += 1
                                line = line.replace(',', '.')
                                temp_str = line.split(' ')
                                self.data[index_temper].value_temper = float(temp_str[0])
                                self.data[index_temper].value_temper_acp = int(temp_str[1])
                            if flag_data_for_pull:
                                index_pull += 1
                                line = line.replace(',', '.')
                                temp_str = line.split(' ')
                                self.data[index_p].pulls.append(SettingsPull())
                                self.data[index_p].pulls[index_pull].value_pull = float(temp_str[0])
                                self.data[index_p].pulls[index_pull].value_pull_acp = int(temp_str[1])
                                pass

            b_array = self.createBitsArray()
            return b_array

        except Exception as e:
            self.logger.error(str(e))

    def createBitsArray(self):
        try:
            str_temp = ''
            if len(str(self.number_sens)) < 8:
                for i in range(8-len(str(self.number_sens))):
                    str_temp += ' '
                str_temp += str(self.number_sens)

            byte_str = []
            for symb in str_temp:
                str_d = self.byteToStr(ord(symb))
                byte_str.append(str_d)

            str_d = self.byteToStr(ord('T'))
            byte_str.append(str_d)
            str_d = self.byteToStr(ord(' '))
            byte_str.append(str_d)

            for i in range(len(self.data)):
                #Значения температуры
                temp_val = self.data[i].value_temper
                ba = struct.pack('>f', temp_val)
                for j in range(4):
                    val_h = self.byteToStr(ba[j])
                    byte_str.append(val_h)

                #Значения температуры в битах АЦП
                temp_val = self.data[i].value_temper_acp
                ba = struct.pack('>H', temp_val)
                for j in range(2):
                    val_h = self.byteToStr(ba[j])
                    byte_str.append(val_h)

                if i == len(self.data)-1:
                    str_d = self.byteToStr(ord('E'))
                    byte_str.append(str_d)
                else:
                    str_d = self.byteToStr(ord(' '))
                    str_d.zfill(2)
                    byte_str.append(str_d)
                pass

            #Данные по усилиям при определенной температуре
            for i in range(len(self.data)):
                str_d = self.byteToStr(ord('F'))
                byte_str.append(str_d)

                str_d = self.byteToStr(ord(' '))
                byte_str.append(str_d)

                for j in range(len(self.data[i].pulls)):
                    temp_val = self.data[i].pulls[j].value_pull
                    ba = struct.pack('>f', temp_val)
                    for k in range(4):
                        val_h = self.byteToStr(ba[k])
                        byte_str.append(val_h)

                    temp_val = self.data[i].pulls[j].value_pull_acp
                    ba = struct.pack('>i', temp_val)
                    for k in range(1, 4):
                        val_h = self.byteToStr(ba[k])
                        byte_str.append(val_h)

                    if j == len(self.data[i].pulls)-1:
                        str_d = self.byteToStr(ord('E'))
                        byte_str.append(str_d)
                    else:
                        str_d = self.byteToStr(ord(' '))
                        byte_str.append(str_d)
                    pass

            len_ba = len(byte_str)
            if not len_ba/4 == len_ba//4:
                byte_str.append('ff')

            for i in range(0,len(byte_str), 2):
                a = byte_str[i]
                b = byte_str[i+1]
                byte_str[i] = b
                byte_str[i+1] = a

            byte2_array = []
            for i in range(0, len(byte_str), 2):

                if len(byte_str[i]) < 2:
                    byte_str[i].zfill(2)
                if len(byte_str[i+1]) < 2:
                    byte_str[i].zfill(2)

                str_hex = byte_str[i]+byte_str[i+1]
                c = bytes.fromhex(str_hex)
                val_d = int.from_bytes(c, 'big', signed=False)
                byte2_array.append(val_d)
            return byte2_array

        except Exception as e:
            self.logger.error(str(e))

    def byteToStr(self, val_d):
        str_hex = (hex(val_d))[2:]
        if len(str_hex) < 2:
            str_hex = '0' + str_hex
        return str_hex

    def printStructure(self):
        for i in range(len(self.data)):
            print('====================')
            print('Температура: {}  {}'.format(self.data[i].value_temper, self.data[i].value_temper_acp))
            print('====================')
            for j in range(len(self.data[i].pulls)):
                print('Усилия: {}  {}'.format(self.data[i].pulls[j].value_pull, self.data[i].pulls[j].value_pull_acp))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    WS = WindowSensor(1)
    WS.setVisible(True)
    sys.exit(app.exec_())
