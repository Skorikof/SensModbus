import configparser
import serial
import sys
import LogPrg
from View import Ui_MainWindow
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMessageBox, QAction, QMenu, QTableWidget, QFrame, QTableWidgetItem, QHeaderView, QLabel
from PyQt5.QtCore import Qt, QObject, QThreadPool, pyqtSignal
from PyQt5.QtGui import QFont, QIcon
from pymodbus.client.sync import ModbusSerialClient as ModbusClient
from Threads import Reader, Writer, FileWriter
from struct import pack, unpack
from ViewSens import WindowSensor
import datetime
from functools import partial
import time
from ValidNumbers import TableValidNumbers


class WindowSignals(QObject):
    signalStart = pyqtSignal()
    signalStop = pyqtSignal()
    signalExit = pyqtSignal()

    signalStartSens = pyqtSignal()
    signalStopSens = pyqtSignal()
    signalExitSens = pyqtSignal()

    signalReadKoeff = pyqtSignal()

    signalChangeID = pyqtSignal(int)
    signalCancelSaveFile = pyqtSignal()
    signalReadBaseTx = pyqtSignal()
    signalReadSensInfo = pyqtSignal(bool)
    signalReadValidNumber = pyqtSignal(bool)
    signalReadSensTx = pyqtSignal()

    signalChangeBroadcastRange = pyqtSignal(bool)


class Process(object):
    def __init__(self):
        self.flag_read_sens_info = False
        self.flag_read_single_sens = False
        self.flag_stop = False
        self.select_id = 0
        self.timeBaseTx = 0  # время между передачей
        self.select_index = 0
        self.timeSensTx = 0
        self.trans_power = 0
        self.adr_modulation = 0
        self.freqOffset = 0
        self.numberChannelBase = 0  # номер канала базовой станции
        self.numberChannelSens = 0  # номер канала текущего датчика


class TableWidget(QTableWidget):
    def __init__(self, parent=None):
        super(TableWidget, self).__init__(parent)
        self.mouse_press = None

    def mousePressEvent(self, event):
        print(int(event.button()))
        self.mouse_press = None
        if event.button() == Qt.LeftButton:
            self.mouse_press = "mouse left press"
        if int(event.button()) == Qt.RightButton:
            self.mouse_press = "mouse right press"
        if int(event.button()) == Qt.MidButton:
            self.mouse_press = "mouse middle press"
        super(TableWidget, self).mousePressEvent(event)


class ApplicationWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super(ApplicationWindow, self).__init__(parent)

        self.signals = WindowSignals()

        self.win_view = Ui_MainWindow()
        self.win_view.setupUi(self)

        self.logger = LogPrg.get_logger(__name__)

        self.setWindowIcon(QIcon('ico/sensor.png'))
        self.resize(1050, 508)
        self.setWindowTitle('Настройка датчиков по MODBUS v.1.05')

        self.exit_action = QAction(self)
        self.exit_action.setText("Выход")
        self.exit_action.setData("exit")
        self.exit_action.triggered.connect(self.exit_prg)
        self.win_view.menubar.addAction(self.exit_action)

        self.set_prg = SettingsPrg()
        self.createMenuPorts()

        self.scan_action = QAction('Сканировать', self)
        self.scan_action.triggered.connect(self.scan_sens)
        self.win_view.menubar.addAction(self.scan_action)

        self.menuStation = QMenu('Базовая станция', self)

        self.correctTime_action = QAction('Корректировка времени', self)
        self.correctTime_action.triggered.connect(self.correcttime)

        self.time_transmission_action = QAction('Время между передачей', self)
        self.time_transmission_action.triggered.connect(self.editTimeTx)

        self.num_channel_action = QAction('Номер канала', self)
        self.num_channel_action.triggered.connect(self.editNumChannelBase)

        self.transmitter_power_action = QAction('Мощность передатчика', self)
        self.transmitter_power_action.triggered.connect(self.editTransmitPower)

        self.modulation_address_action = QAction('Адрес модуляции', self)
        self.modulation_address_action.triggered.connect(self.editAddrModulation)

        self.frequency_adjustment_action = QAction('Корректировка по чатоте', self)
        self.frequency_adjustment_action.triggered.connect(self.editFrequency)

        self.valid_pairs_action = QAction('Таблица валидных номеров', self)
        self.valid_pairs_action.triggered.connect(self.validPairs)
        
        self.menuStation.addAction(self.correctTime_action)
        self.menuStation.addAction(self.time_transmission_action)
        self.menuStation.addAction(self.num_channel_action)
        self.menuStation.addAction(self.transmitter_power_action)
        self.menuStation.addAction(self.modulation_address_action)
        self.menuStation.addAction(self.frequency_adjustment_action)
        self.menuStation.addAction(self.valid_pairs_action)
        self.menuStation.setEnabled(False)
        
        self.win_view.menubar.addMenu(self.menuStation)

        self.tableWidget = TableWidget()
        self.tableWidget.setObjectName(u"tableWidget")

        self.tableWidget.setColumnCount(11)

        self.tableWidget.setHorizontalHeaderLabels(['Название', 'Серийный\nномер', 'Номер в\nподсети', 'Усилие,\nкгс/см2',
                                                    'Темпер,\n⁰С', 'Темпер\nдатчика,⁰С', 'Uпит.,\nВ',
                                                    'RSSI', 'Смещение', 'Время,\nмс', 'Статус'])
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        font = QFont('Calibri', 14)
        self.tableWidget.setFont(font)
        font = QFont('Calibri', 11, QFont.Bold)
        self.tableWidget.horizontalHeader().setFont(font)

        header = self.tableWidget.horizontalHeader()
        header.setFrameStyle(QFrame.Box | QFrame.Plain)
        header.setLineWidth(1)
        self.tableWidget.setHorizontalHeader(header)

        self.tableWidget.cellDoubleClicked[int, int].connect(self.clickedRowColumn)

        self.win_view.gridLayout.addWidget(self.tableWidget, 0, 0, 1, 1)
        self.is_scan = False
        self.tableWidget.mouse_press = None
        self.process = Process()
        self.win_sens = WindowSensor(self.logger)
        self.win_sens.signals.win_closed.connect(self.closeWindowSensor)
        self.win_sens.signals.change_Sens_ID.connect(self.changeSensID)
        self.win_sens.signals.query_save_file.connect(self.querySaveFile)
        self.win_sens.signals.cancel_Save_File.connect(self.cancelSaveFile)
        self.win_sens.signals.query_clear_mem.connect(self.queryClearMemory)
        self.win_sens.signals.query_null_sens.connect(self.nullSens)
        self.win_sens.signals.clear_table.connect(self.clearTable)
        self.win_sens.signals.query_serial_number.connect(self.editSerialNumber)
        self.win_sens.signals.query_TimeTx.connect(self.editSensTimeTx)
        self.win_sens.signals.query_ChannelNumber.connect(self.editNumChannelSens)
        self.win_sens.signals.query_TransPower.connect(self.editTransmitPowerSens)
        self.win_sens.signals.query_AddrModulation.connect(self.editAddrModulationSens)
        self.win_sens.signals.query_FreqOffset.connect(self.editFrequencySens)

        self.win_sens.tableKoeff.cellChanged.connect(self.editCellKoeff)
        self.win_sens.progress_dialog.close()
        self.pool_Exchange = QThreadPool()
        self.pool_Exchange.setMaxThreadCount(5)
        print("Multithreading with maximum %d threads" % self.pool_Exchange.maxThreadCount())
        self.pool_Exchange.clear()

        # self.initStatusBar()

    def initStatusBar(self):
        try:
            self.lbl_info_base = QLabel("Ответ Базы:")
            self.lbl_info_base.setStyleSheet('border: 0; color:  green; ')
            self.lbl_info_base.setFont(QFont('Calibri', 12))

            self.lbl_info_channel = QLabel("Канал:")
            self.lbl_info_channel.setStyleSheet('border: 0; color:  black;')
            self.lbl_info_channel.setFont(QFont('Calibri', 12))

            self.lbl_info_power = QLabel("Мощность:")
            self.lbl_info_power.setStyleSheet('border: 0; color:  black;')
            self.lbl_info_power.setFont(QFont('Calibri', 12))

            self.lbl_info_addrMod = QLabel("Адрес:")
            self.lbl_info_addrMod.setStyleSheet('border: 0; color:  black;')
            self.lbl_info_addrMod.setFont(QFont('Calibri', 12))

            self.lbl_info_freq = QLabel("Смещение:")
            self.lbl_info_freq.setStyleSheet('border: 0; color:  black;')
            self.lbl_info_freq.setFont(QFont('Calibri', 12))

            self.statusBar().reformat()
            self.statusBar().setStyleSheet('border: 0; background-color: #FFF8DC;')
            self.statusBar().setStyleSheet("QStatusBar::item {border: none;}")

            self.statusBar().addPermanentWidget(QFrame.VLine())  # <---
            self.statusBar().addPermanentWidget(self.lbl_info_base, stretch=1)
            self.statusBar().addPermanentWidget(QFrame.VLine())  # <---
            self.statusBar().addPermanentWidget(self.lbl_info_channel, stretch=1)
            self.statusBar().addPermanentWidget(QFrame.VLine())  # <---
            self.statusBar().addPermanentWidget(self.lbl_info_power, stretch=1)
            self.statusBar().addPermanentWidget(QFrame.VLine())  # <---
            self.statusBar().addPermanentWidget(self.lbl_info_addrMod, stretch=1)
            self.statusBar().addPermanentWidget(QFrame.VLine())  # <---
            self.statusBar().addPermanentWidget(self.lbl_info_freq, stretch=1)

        except Exception as e:
            self.logger.error(str(e))

    def editNumChannelBase(self):
        try:
            number, ok = QtWidgets.QInputDialog.getInt(self.win_sens, 'Данные',
                                                       "Введите номер канала базы (текущее значение: {})".format(
                                                           self.process.numberChannelBase),
                                                       value=self.process.numberChannelBase, min=0, max=7, step=1)
            if ok:
                val = []
                val.append(number)
                self.writeValues(1, 6, val)

        except Exception as e:
            self.logger.error(str(e))

    def editNumChannelSens(self):
        try:
            number, ok = QtWidgets.QInputDialog.getInt(self.win_sens, 'Данные',
                                                       "Введите номер канала от 0 до 7 (текущее значение: {})".format(
                                                           self.process.numberChannelSens),
                                                       value=self.process.numberChannelSens, min=0, max=7, step=1)
            if ok:
                val = []
                val.append(number)
                self.writeValues(self.set_prg.selectID, 6, val)
                time.sleep(7.0)
                self.signals.signalReadSensInfo.emit(True)

        except Exception as e:
            self.logger.error(str(e))

    def editSensTimeTx(self):
        try:
            number, ok = QtWidgets.QInputDialog.getInt(self.win_sens, 'Данные',
                                                       "Введите время между передачей (текущее значение: {})".format(
                                                           self.process.timeSensTx),
                                                       value=self.process.timeSensTx, min=5, max=65535, step=1)
            if ok:
                val = []
                val.append(number)
                self.writeValues(self.set_prg.selectID, 5, val)
                time.sleep(7.0)
                self.signals.signalReadSensTx.emit()

        except Exception as e:
            self.logger.error(str(e))

    def editTimeTx(self):
        try:
            number, ok = QtWidgets.QInputDialog.getInt(self.win_sens, 'Данные',
                                                       "Введите время между передачей (текущее значение: {})".format(
                                                           self.process.timeBaseTx),
                                                       value=self.process.timeBaseTx, min=5, max=65535, step=1)
            if ok:
                val = []
                val.append(number)
                self.writeValues(1, 5, val)
                time.sleep(5.0)
                self.signals.signalReadBaseTx.emit()

        except Exception as e:
            self.logger.error(str(e))

    def editTransmitPower(self):
        try:
            number, ok = QtWidgets.QInputDialog.getInt(self.win_sens, 'Данные',
                                               "Введите мощность передатчика от 0 до 15 (текущее значение: {})".format(
                                                   self.process.trans_power),
                                               value=self.process.trans_power, min=0, max=15, step=1)
            if ok:
                val = []
                val.append(number)
                self.writeValues(1, 7, val)
                # time.sleep(5.0)
                # self.signals.signalReadBaseTx.emit()

        except Exception as e:
            self.logger.error(str(e))

    def editTransmitPowerSens(self):
        try:
            number, ok = QtWidgets.QInputDialog.getInt(self.win_sens, 'Данные',
                                               "Введите мощность передатчика от 0 до 15 (текущее значение: {})".format(
                                                   self.process.trans_power),
                                               value=self.process.trans_power, min=0, max=15, step=1)
            if ok:
                val = []
                val.append(number)
                self.writeValues(self.process.select_id, 7, val)
                # time.sleep(5.0)
                # self.signals.signalReadBaseTx.emit()

        except Exception as e:
            self.logger.error(str(e))

    def editAddrModulation(self):
        try:
            number, ok = QtWidgets.QInputDialog.getInt(self.win_sens, 'Данные',
                                               "Введите номер адреса модуляции от 0 до 12 (текущее значение: {})".format(
                                                   self.process.adr_modulation),
                                               value=self.process.adr_modulation, min=0, max=12, step=1)
            if ok:
                val = []
                val.append(number)
                self.writeValues(1, 11, val)
                # time.sleep(5.0)
                # self.signals.signalReadBaseTx.emit()

        except Exception as e:
            self.logger.error(str(e))

    def editAddrModulationSens(self):
        try:
            number, ok = QtWidgets.QInputDialog.getInt(self.win_sens, 'Данные',
                                               "Введите номер адреса модуляции от 0 до 12 (текущее значение: {})".format(
                                                   self.process.adr_modulation),
                                               value=self.process.adr_modulation, min=0, max=12, step=1)
            if ok:
                val = []
                val.append(number)
                self.writeValues(self.process.select_id, 11, val)
                # time.sleep(5.0)
                # self.signals.signalReadBaseTx.emit()

        except Exception as e:
            self.logger.error(str(e))

    def editFrequency(self):
        try:
            number, ok = QtWidgets.QInputDialog.getInt(self.win_sens, 'Данные',
                                                       "Введите смещение по частоте от -16384 до 16384,"
                                                       "1 = 0.645 Гц (текущее значение: {})".format(
                                                           self.process.freqOffset),
                                                       value=self.process.freqOffset, min=-16384, max=16384, step=1)
            if ok:
                val = []
                val.append(number)
                self.writeValues(1, 12, val)
                # time.sleep(5.0)
                # self.signals.signalReadBaseTx.emit()

        except Exception as e:
            self.logger.error(str(e))

    def editFrequencySens(self):
        try:
            number, ok = QtWidgets.QInputDialog.getInt(self.win_sens, 'Данные',
                                                       "Введите смещение по частоте от -16384 до 16384,"
                                                       "1 = 0.645 Гц (текущее значение: {})".format(
                                                           self.process.freqOffset),
                                                       value=self.process.freqOffset, min=-16384, max=16384, step=1)
            if ok:
                val = []
                val.append(number)
                self.writeValues(self.process.select_id, 12, val)
                # time.sleep(5.0)
                # self.signals.signalReadBaseTx.emit()

        except Exception as e:
            self.logger.error(str(e))

    def validPairs(self):
        try:
            self.win_numbers = TableValidNumbers(self.logger)
            self.win_numbers.signals.query_saveValidNum.connect(self.saveValidNumber)
            self.signals.signalReadValidNumber.emit(True)

        except Exception as e:
            self.logger.error(str(e))

    def saveValidNumber(self, id, addr, val_s):
        try:
            val = []
            val.append(val_s)
            self.writeValues(1, addr, val)
            time.sleep(2.0)
            self.signals.signalReadValidNumber.emit(True)

        except Exception as e:
            self.logger.error(str(e))

    def editSerialNumber(self):
        try:
            number, ok = QtWidgets.QInputDialog.getInt(self.win_sens, 'Данные', "Введите серийный номер",
                                                       value=self.set_prg.select_serial_number, min=0, max=65535, step=1)
            if ok:
                val = []
                val.append(number)
                self.writeValues(self.set_prg.selectID, 1, val)
                time.sleep(7.0)

                self.signals.signalReadSensInfo.emit(True)

        except Exception as e:
            self.logger.error(str(e))

    def readSerialNumber(self, data):
        try:
            val = data.registers[1]
            self.set_prg.select_serial_number = val
            self.set_prg.serial_numbers[self.process.select_index] = val

            self.tableWidget.setItem(self.process.select_index, 1, val)

            val1 = data.registers[5]
            self.process.timeSensTx = val1

            val1 = data.registers[6]
            self.process.numberChannelSens = val1

            print('Serial № ' + str(val))
            print('Channel № ' + str(val1))

            nam_type = 'температуры'
            if self.set_prg.sensTypes[self.process.select_index] == 'force':
                nam_type = 'усилия'
            str_title = 'Датчик {0}, (Serial № : {1}) , Slave ID={2}'.format(nam_type, val, self.process.select_id)
            self.win_sens.setWindowTitle(str_title)

        except Exception as e:
            self.logger.error(str(e))

    def cancelSerialNumber(self):
        self.win_edit.setVisible(False)

    def clearTable(self):
        self.set_prg.isEditableKoeff = True

    def createMenuPorts(self):
        try:
            flag_checked = False
            self.menuSettings = QMenu('&Настройки', self)
            self.menuNumberComm=self.menuSettings.addMenu('&COM-порт')

            self.send_1 = QAction('Записывать 1 в регистр 1015', self)
            self.send_1.setCheckable(True)
            self.send_1.setChecked(False)
            if self.set_prg.send_1:
                self.send_1.setChecked(True)
            self.send_1.triggered.connect(self.send_1_Check)

            self.menuSettings.addAction(self.send_1)
            self.win_view.menubar.addMenu(self.menuSettings)
            actions = []
            mnu_nam = self.set_prg.available_ports
            for comms in mnu_nam:
                number_port = comms[3:]
                action = QAction(comms, self)
                action.isChecked()
                action.setCheckable(True)
                if str.upper(action.text()) == str.upper(self.set_prg.port):
                    action.setChecked(True)
                    flag_checked = True
                    self.setClient()
                else:
                    action.setChecked(False)
                action.triggered.connect(partial(self.menuComPortClick, action))
                actions.append(action)
            self.menuNumberComm.addActions(actions)

            if len(actions) == 0:
                self.showMsg('error', 'Подключение', 'В системе отсутствуют COM-порты...', True)
            else:
                if flag_checked == False:
                    action = actions[0]
                    self.menuComPortClick(action)

        except Exception as e:
            self.logger.error(str(e))

    def send_1_Check(self):
        try:
            self.set_prg.send_1 = not self.set_prg.send_1
            if self.set_prg.send_1:
                self.send_1.setChecked(True)
            else:
                self.send_1.setChecked(False)
            self.saveFileSettings()

        except Exception as e:
            self.logger.error(str(e))

    def changeSensID(self, num_id):
        try:
            val = []
            val.append(num_id)
            self.writeValues(self.set_prg.selectID, 0, val)
            self.signals.signalChangeID.emit(num_id)

        except Exception as e:
            self.logger.error(str(e))

    def menuComPortClick(self, act):
        try:
            self.menuNumberComm.clear()
            actions = []
            mnu_nam = self.set_prg.available_ports
            for comms in mnu_nam:
                number_port = comms[3:]
                action = QAction(comms, self)
                action.setCheckable(True)
                if str.upper(action.text()) == str.upper(act.text()):
                    action.setChecked(True)
                    self.set_prg.port = action.text()
                else:
                    action.setChecked(False)
                action.triggered.connect(partial(self.menuComPortClick, action))
                actions.append(action)
            self.menuNumberComm.addActions(actions)
            self.saveFileSettings()
            self.setClient()
            print('Select port: {}'.format(act.text()))

        except Exception as e:
            self.logger.error(str(e))

    def saveFileSettings(self):
        try:
            with open('Settings.ini', 'w') as f:
                f.write('[ComPort]'+'\n')
                f.write('NumberPort={}'.format(self.set_prg.port[3:])+'\n')
                f.write('PortSettings={0},{1},{2},{3}\n'.format(self.set_prg.baudrate, self.set_prg.parity,
                                                                self.set_prg.databits, self.set_prg.stopbits))
                f.write('Send_1_to_1015={}\n'.format(self.set_prg.send_1))
                f.write('Timeout(ms)={}\n'.format(self.set_prg.timeout_s))
                f.write('TimeBetweenSend(ms)={}\n'.format(self.set_prg.timeout_BS))

        except Exception as e:
            self.logger.error(str(e))

    def setClient(self):
        try:
            if self.set_prg.timeout_s == 0:
                self.client = ModbusClient(method='rtu', port=self.set_prg.port,
                                           parity=self.set_prg.parity,
                                           baudrate=self.set_prg.baudrate,
                                           databits=self.set_prg.databits,
                                           stopbits=self.set_prg.stopbits, strict=False)
            else:
                self.client = ModbusClient(method='rtu', port=self.set_prg.port,
                                           parity=self.set_prg.parity,
                                           baudrate=self.set_prg.baudrate,
                                           databits=self.set_prg.databits,
                                           stopbits=self.set_prg.stopbits, strict=False,
                                           timeout=self.set_prg.timeout_s/1000)

        except Exception as e:
            self.logger.error(str(e))

    def correcttime(self):
        try:
            timestamp_linux = datetime.datetime(year=1970, month=1, day=1, hour=0, minute=0, second=0)
            timestamp_current = datetime.datetime.now()
            d = timestamp_current-timestamp_linux
            total_sec = int(d.total_seconds())
            total_sec_hex = hex(total_sec)
            total_sec_hex = total_sec_hex[2:]
            ba = bytearray.fromhex(total_sec_hex)
            val_array = []

            b = []
            b.append(ba[0])
            b.append(ba[1])
            c = bytes(b)
            val_d = int.from_bytes(c, 'big', signed=False)
            val_array.append(val_d)

            b = []
            b.append(ba[2])
            b.append(ba[3])
            c = bytes(b)
            val_d = int.from_bytes(c, 'big', signed=False)
            val_array.append(val_d)

            self.writeValues(1, 4160, val_array)

        except Exception as e:
            self.logger.error(str(e))
       
    def exit_prg(self):
        self.close()

    def closeEvent(self, event): 
        self.signals.signalExit.emit()
        self.signals.signalExitSens.emit()
        self.signals.signalCancelSaveFile.emit()
        self.pool_Exchange.waitForDone()
        self.client.close()
        print('exit prg')

    def showWindowSensor(self):
        self.setEnabled(False)
        self.win_sens = WindowSensor(4)
        self.win_sens.signals.win_closed.connect(self.closeWindowSensor)
        self.win_sens.setVisible(True)

    def closeWindowSensor(self):
        self.signals.signalExitSens.emit()
        if self.set_prg.selectType == 'force':
            self.showMsg('info', 'Настройки', 'Отключите перемычку и нажмите Ок!', False)
        self.signals.signalStart.emit()
        self.setEnabled(True)
    
    def startStopSettings(self, val_int):
        val_a = []
        val_a.append(val_int)
        self.writeValues(1, 4117, val_a)

    def scan_sens(self):
        try:
            if self.set_prg.isRunExchange == False:
                self.port_connect=self.client.connect()
                if self.port_connect:
                    self.set_prg.isRunExchange = True
                    self.start_exchange()
                else:
                    self.showMsg('error', 'Подключение',
                                 'Отсутствует подключение по порту {} или порт занят другим процессом'.format(
                                     self.set_prg.port), False)
            else:
                self.stopExange()

        except Exception as e:
            self.logger.error(str(e))

    def stopExange(self):
        try:
            self.set_prg.isRunExchange = False
            self.signals.signalExit.emit()
            self.pool_Exchange.waitForDone()
            self.tableWidget.setRowCount(0)
            self.client.close()
            self.scan_action.setText('Сканировать')
            self.menuSettings.setEnabled(True)
            self.menuStation.setEnabled(False)

        except Exception as e:
            self.logger.error(str(e))

    def showMsg(self, type_m, text_title, text_info, flag_exit):
        try:
            msg = QMessageBox()
            if type_m == 'error':
                a = QMessageBox.Critical
            if type_m == 'info':
                a = QMessageBox.Information
            if type_m == 'warning':
                a = QMessageBox.Warning
            msg.setIcon(a)
            msg.setStandardButtons(QMessageBox.Ok)
            msg.setText(text_title)
            msg.setInformativeText(text_info)
            msg.setWindowFlag(Qt.WindowStaysOnTopHint)
            retval = msg.exec_()
            if flag_exit:
                sys.exit(0)

        except Exception as e:
            self.logger.error(str(e))

    def start_exchange(self):
        try:
            self.menuSettings.setEnabled(False)
            self.scan_action.setText('Прервать сканирование')
            self.menuStation.setEnabled(True)

            if self.set_prg.send_1:
                self.startStopSettings(1)
            else:
                self.startStopSettings(0)

            thread_base_info = Reader('read_base_info', 1, self.client, self.set_prg.timeout_BS, self.set_prg.send_1)
            thread_base_info.signals.result.connect(self.result_read)
            thread_base_info.signals.error_read.connect(self.error_read)
            thread_base_info.signals.time_received.connect(self.time_from_baseStation)
            thread_base_info.signals.count_sensors.connect(self.count_sensors_from_baseStation)
            thread_base_info.signals.result_types.connect(self.result_read_types)
            thread_base_info.signals.timeBaseTx.connect(self.result_timeBaseTx)
            thread_base_info.signals.result_ValidNumbers.connect(self.result_ValidNumbers)

            thread_base_info.signals.msgToStatusBar.connect(self.msgFromThread)
            thread_base_info.signals.result_infoBase.connect(self.resultReadInfoBase)

            self.signals.signalStart.connect(thread_base_info.startThread)
            self.signals.signalStop.connect(thread_base_info.stopThread)
            self.signals.signalExit.connect(thread_base_info.exitThread)
            self.signals.signalReadBaseTx.connect(thread_base_info.readBaseTx)
            self.signals.signalReadValidNumber.connect(thread_base_info.readValidNumbers)
            self.signals.signalChangeBroadcastRange.connect(thread_base_info.changeBroadCastRange)

            self.signals.signalReadBaseTx.emit()

            self.pool_Exchange.start(thread_base_info)
            self.signals.signalStart.emit()

            self.signals.signalChangeBroadcastRange.emit(self.set_prg.send_1)

        except Exception as e:
            self.logger.error(str(e))

    def resultReadInfoBase(self, obj):
        try:
            self.process.numberChannelBase = obj.registers[6]
            self.process.trans_power = obj.registers[7]
            self.process.adr_modulation = obj.registers[11]
            self.process.freqOffset = obj.registers[12]

        except Exception as e:
            self.logger.error(str(e))

    def clickedRowColumn(self, r, c):
        try:
            self.set_prg.selectType = self.set_prg.sensTypes[r]
            self.set_prg.select_serial_number = self.set_prg.serial_numbers[r]

            self.signals.signalStop.emit()

            num_sens = int(self.tableWidget.item(r, 2).text())

            if self.set_prg.sensTypes[r] == 'force':
                self.showMsg('info', 'Настройки', 'Подключите перемычку и нажмите Ок!', False)

            self.set_prg.selectID = num_sens
            self.setEnabled(False)
            self.process.select_id = num_sens
            self.process.select_index = r
            self.win_sens.initWindow(self.process.select_id, self.set_prg.sensTypes[r], self.set_prg.list_sens_id,
                                     self.set_prg.select_serial_number)
            self.win_sens.clearTables()
            self.win_sens.setVisible(True)

            thread_sens = Reader('read_sensor', self.process.select_id, self.client, self.set_prg.timeout_BS,
                                 self.set_prg.send_1)
            thread_sens.signals.result_sens_current.connect(self.result_sens_read)
            thread_sens.signals.result_sens_koeff.connect(self.result_koeff_read)
            thread_sens.signals.error_readSens.connect(self.error_readSens)
            thread_sens.signals.editableKoeff.connect(self.setEditableKoeff)
            thread_sens.signals.result_SerialNumber.connect(self.readSerialNumber)
            thread_sens.signals.timeSensorTx.connect(self.result_timeSensorTx)

            self.signals.signalReadKoeff.connect(thread_sens.readKoeff)
            self.signals.signalStartSens.connect(thread_sens.startThread) 
            self.signals.signalStopSens.connect(thread_sens.stopThread)             
            self.signals.signalExitSens.connect(thread_sens.exitThread)
            self.signals.signalChangeID.connect(thread_sens.changeID)
            self.signals.signalReadSensInfo.connect(thread_sens.readSensInfo)
            self.signals.signalReadSensTx.connect(thread_sens.readSensTx)

            self.pool_Exchange.start(thread_sens)
            self.signals.signalStartSens.emit()
            self.signals.signalReadSensInfo.emit(True)
            self.set_prg.isEditableKoeff=False

        except Exception as e:
            print(str(e))

    def result_timeSensorTx(self, val):
        self.process.timeSensTx = val

    def result_timeBaseTx(self, val):
        self.process.timeBaseTx = val
    
    def result_ValidNumbers(self, obj):
        list_numb = []
        for i in range(len(obj.registers)):
            if i > 1:
                val = obj.registers[i]
                list_numb.append(val)
        self.win_numbers.fillTable(list_numb)

    def cancelSaveFile(self):
        self.signals.signalCancelSaveFile.emit()
        self.win_sens.progress_dialog.close()

    def showPercent(self, val_p):
        if val_p > 0:
            self.win_sens.progress_dialog.show()
            self.win_sens.progress_dialog.setValue(val_p)

    def count_sensors_from_baseStation(self, count_s, list_s):
        print('Sensors: {}'.format(count_s))
        self.tableWidget.setRowCount(count_s)
        self.set_prg.list_sens_id = list_s

    def time_from_baseStation(self, obj_time):
        try:
            self.setStyleStBar('info')
            # self.lbl_info_base.setText('Время с базовой станции: ' + obj_time.strftime('%d-%m-%Y %H:%M:%S'))
            self.win_view.statusbar.showMessage('Время с базовой станции: ' + obj_time.strftime('%d-%m-%Y %H:%M:%S'))

        except Exception as e:
            self.logger.error(str(e))

    def setStyleStBar(self, str_style):
        if str_style == 'info':
            self.setStyleSheet("QStatusBar{padding-left:8px;background:rgba(0,0,0,0);"
                               "color:green;font: 75 12pt \"Calibri\";font-weight:bold;}")
        if str_style == 'error':
            self.setStyleSheet("QStatusBar{padding-left:8px;background:rgba(0,0,0,0);"
                               "color:red;font: 75 12pt \"Calibri\";font-weight:bold;}")

    def setEditableKoeff(self):
        self.set_prg.isEditableKoeff = True

    def editCellKoeff(self, row_e, col_e):
        try:
            if not self.set_prg.isEditableKoeff:
                print('Edit: {} {}'.format(row_e, col_e))
            else:
                # смещение нуля
                if row_e == 0:
                    val_str = self.win_sens.tableKoeff.item(row_e, col_e).text()
                    try:
                        val_d = int(val_str)
                        val_regs = []
                        b = pack('>i', val_d)
                        byt = []
                        byt.append(b[0])
                        byt.append(b[1])
                        val_d = int.from_bytes(byt, 'big', signed=False)
                        val_regs.append(val_d)
                        byt = []
                        byt.append(b[2])
                        byt.append(b[3])
                        val_d = int.from_bytes(byt, 'big', signed=False)
                        val_regs.append(val_d)
                        self.writeValues(self.set_prg.selectID, 23, val_regs)

                    except Exception as e:
                        self.showMsg('warning', 'Внимание', 'Ошибка при вводе значения \n{}'.format(e), False)
                        self.logger.error(str(e))
                    self.set_prg.isEditableKoeff = False
                    self.signals.signalReadKoeff.emit()

                # вес бита АЦП
                if row_e == 1:
                    val_str = self.win_sens.tableKoeff.item(row_e, col_e).text()
                    try:
                        val_d = float(val_str)
                        val_regs = []
                        b = pack('>f', val_d)
                        byt = []
                        byt.append(b[0])
                        byt.append(b[1])
                        val_d = int.from_bytes(byt, 'big', signed=False)
                        val_regs.append(val_d)
                        byt = []
                        byt.append(b[2])
                        byt.append(b[3])
                        val_d = int.from_bytes(byt, 'big', signed=False)
                        val_regs.append(val_d)
                        self.writeValues(self.set_prg.selectID, 25, val_regs)

                    except Exception as e:
                        self.showMsg('warning', 'Внимание', 'Ошибка при вводе значения \n{}'.format(e), False)
                        self.logger.error(str(e))

                    self.set_prg.isEditableKoeff = False
                    self.signals.signalReadKoeff.emit()

                # коэффициент усиления
                if row_e == 2:
                    values_k = [1, 2, 4, 8, 16, 32, 64, 128]
                    val_str = self.win_sens.tableKoeff.item(row_e, col_e).text()
                    try:
                        flag_c = False
                        val_d = int(val_str)
                        for i in range(len(values_k)):
                            if val_d == values_k[i]:
                                flag_c = True
                                index_v = i
                                i = len(values_k) + 1
                        if flag_c:
                            val_regs = []
                            val_d = index_v
                            val_regs.append(val_d)
                            self.writeValues(self.set_prg.selectID, 27, val_regs)
                        else:
                            self.showMsg('warning', 'Внимание', 'Допустимые значения: 1,2,4,8,16,32,64,128', False)

                    except Exception as e:
                        self.showMsg('warning', 'Внимание', 'Ошибка при вводе значения \n{}'.format(e), False)
                        self.logger.error(str(e))

                    self.set_prg.isEditableKoeff = False
                    self.signals.signalReadKoeff.emit()

                print('Edit by Enter: {} {}'.format(row_e, col_e))

        except Exception as e:
            self.logger.error(str(e))

    def writeValues(self, id, start_addr, values):
        try:
            self.set_prg.flag_OkWrite = False
            thread_Writer = Writer(id, start_addr, values, self.client)
            thread_Writer.signals.error_write.connect(self.error_write)
            thread_Writer.signals.finish.connect(self.finish_write)
            self.pool_Exchange.start(thread_Writer)

        except Exception as e:
            self.logger.error(str(e))

    def result_read_types(self, num_str, num_id, status, result):
        try:
            val_u = unpack('H', pack('<H', result.registers[10]))[0]
            val_str = 'усилие'
            self.set_prg.serial_numbers[num_str-1] = result.registers[1]

            if val_u == 3 or val_u == 7 or val_u == 11:
                # датчик усилия
                self.set_prg.sensTypes[num_str-1] = 'force'
            else:
                # датчик температуры
                self.set_prg.sensTypes[num_str-1] = 'temper'
                val_str = 'темпер-ра'

            # заполнение типа датчика
            item = QTableWidgetItem(val_str)
            item.setFlags(Qt.ItemIsEnabled)
            item.setTextAlignment(Qt.AlignHCenter)
            self.tableWidget.setItem(num_str-1, 0, item)

            # заполнение серийного номера
            val_str = str(result.registers[1])
            item = QTableWidgetItem(val_str)
            item.setFlags(Qt.ItemIsEnabled)
            item.setTextAlignment(Qt.AlignHCenter)
            self.tableWidget.setItem(num_str-1, 1, item)

            # заполнение статуса датчика
            if status == 0:
                item = QTableWidgetItem(QIcon("ico/ok_32.png"), '')
            else:
                item = QTableWidgetItem(QIcon("ico/error_32.png"), '')

            item.setFlags(Qt.ItemIsEnabled)
            item.setTextAlignment(Qt.AlignHCenter)
            self.tableWidget.setItem(num_str - 1, 10, item)

        except Exception as e:
            self.logger.error(str(e))

    def result_read(self, num_str, num_id, result):
        try:


            # заполнение Slave ID
            item = QTableWidgetItem(str(num_id))
            item.setFlags(Qt.ItemIsEnabled)
            item.setTextAlignment(Qt.AlignHCenter)
            self.tableWidget.setItem(num_str-1, 2, item)

            # Заполнение текущего усилия
            val_u = unpack('f', pack('<HH', result.registers[3], result.registers[2]))[0]
            item = QTableWidgetItem(str(round(val_u, 2)))
            item.setFlags(Qt.ItemIsEnabled)
            item.setTextAlignment(Qt.AlignHCenter)
            self.tableWidget.setItem(num_str-1, 3, item)

            # Заполнение текущей температуры
            val_u = unpack('f', pack('<HH', result.registers[5], result.registers[4]))[0]
            item = QTableWidgetItem(str(round(val_u, 2)))
            item.setFlags(Qt.ItemIsEnabled)
            item.setTextAlignment(Qt.AlignHCenter)
            self.tableWidget.setItem(num_str-1, 4, item)

            # Заполнение текущей температуры датчика
            val_u = unpack('f', pack('<HH', result.registers[7], result.registers[6]))[0]
            item = QTableWidgetItem(str(round(val_u, 2)))
            item.setFlags(Qt.ItemIsEnabled)
            item.setTextAlignment(Qt.AlignHCenter)
            self.tableWidget.setItem(num_str-1, 5, item)

            # Заполнение напряжения питания
            val_u = unpack('f', pack('<HH', result.registers[14], result.registers[13]))[0]
            item = QTableWidgetItem(str(round(val_u, 2)))
            item.setFlags(Qt.ItemIsEnabled)
            item.setTextAlignment(Qt.AlignHCenter)
            self.tableWidget.setItem(num_str-1, 6, item)

            if len(result.registers) == 19:
                val_u = unpack('h', pack('<H', result.registers[16]))[0]
                item = QTableWidgetItem(str(val_u))
                item.setFlags(Qt.ItemIsEnabled)
                item.setTextAlignment(Qt.AlignHCenter)
                self.tableWidget.setItem(num_str-1, 7, item)

                val_u = unpack('h', pack('<H', result.registers[17]))[0]
                item = QTableWidgetItem(str(val_u))
                item.setFlags(Qt.ItemIsEnabled)
                item.setTextAlignment(Qt.AlignHCenter)
                self.tableWidget.setItem(num_str-1, 8, item)

                val_u = result.registers[18]
                item = QTableWidgetItem(str(val_u))
                item.setFlags(Qt.ItemIsEnabled)
                item.setTextAlignment(Qt.AlignHCenter)
                self.tableWidget.setItem(num_str-1, 9, item)

        except Exception as e:
            self.logger.error(str(e))

    def result_sens_read(self, tag, result):
        try:
            # текущее усилие
            val_u = unpack('f', pack('<HH', result.registers[3], result.registers[2]))[0]
            item = QTableWidgetItem(str(round(val_u, 2)))
            item.setFlags(Qt.ItemIsEnabled)
            item.setTextAlignment(Qt.AlignHCenter)
            self.win_sens.tableCurrent.setItem(0, 0, item)

            # текущая температура воздуха
            val_u = unpack('f', pack('<HH', result.registers[5], result.registers[4]))[0]
            item = QTableWidgetItem(str(round(val_u, 2)))
            item.setFlags(Qt.ItemIsEnabled)
            item.setTextAlignment(Qt.AlignHCenter)
            self.win_sens.tableCurrent.setItem(1, 0, item)

            # текущая температура датчика усилия
            val_u = unpack('f', pack('<HH', result.registers[7], result.registers[6]))[0]
            item = QTableWidgetItem(str(round(val_u, 2)))
            item.setFlags(Qt.ItemIsEnabled)
            item.setTextAlignment(Qt.AlignHCenter)
            self.win_sens.tableCurrent.setItem(2, 0, item)

            # усилие в единицах АЦП
            val_u = unpack('i', pack('<HH', result.registers[9], result.registers[8]))[0]
            item = QTableWidgetItem(str(round(val_u, 2)))
            item.setFlags(Qt.ItemIsEnabled)
            item.setTextAlignment(Qt.AlignHCenter)
            self.win_sens.tableCurrent.setItem(0, 1, item)

            # температура воздуха в единицах АЦП
            val_u = unpack('I', pack('<HH', result.registers[11], result.registers[10]))[0]
            item = QTableWidgetItem(str(round(val_u, 2)))
            item.setFlags(Qt.ItemIsEnabled)
            item.setTextAlignment(Qt.AlignHCenter)
            self.win_sens.tableCurrent.setItem(1, 1, item)

            # температура датчика усилия в единицах АЦП
            val_u = unpack('H', pack('<H', result.registers[12]))[0]
            item = QTableWidgetItem(str(round(val_u, 2)))
            item.setFlags(Qt.ItemIsEnabled)
            item.setTextAlignment(Qt.AlignHCenter)
            self.win_sens.tableCurrent.setItem(2, 1, item)

            if len(self.win_sens.statusbar.currentMessage()) > 0:
                self.setStyleStBar('info')
                self.win_sens.statusbar.showMessage('')

        except Exception as e:
            self.logger.error(str(e))

    def result_koeff_read(self, tag, result):
        try:
            # Смещение нуля АЦП
            val_u = unpack('i', pack('<HH', result.registers[1], result.registers[0]))[0]
            item = QTableWidgetItem(str(int(val_u)))
            item.setTextAlignment(Qt.AlignHCenter)
            self.win_sens.tableKoeff.setItem(0, 0, item)

            # Вес бита АЦП
            val_u = unpack('f', pack('<HH', result.registers[3], result.registers[2]))[0]
            item = QTableWidgetItem(str(round(val_u, 6)))
            item.setTextAlignment(Qt.AlignHCenter)
            self.win_sens.tableKoeff.setItem(1, 0, item)

            # коэффициент усиления АЦП
            val_u = result.registers[4]
            val_u = 2 ** val_u
            item = QTableWidgetItem(str(int(val_u)))
            item.setTextAlignment(Qt.AlignHCenter)
            self.win_sens.tableKoeff.setItem(2, 0, item)
            if len(self.win_sens.statusbar.currentMessage()) > 0:
                self.setStyleStBar('info')
                self.win_sens.statusbar.showMessage('')

        except Exception as e:
            self.logger.error(str(e))

    def error_read(self, tag, result):
        print('Tag: {} Result: {}'.format(tag, result))
        self.setStyleStBar('error')
        # self.lbl_info_base.setText('{}'.format(result))
        self.win_view.statusbar.showMessage('{}'.format(result))
        self.saveLogFile('BASE READ', '{}'.format(result))

    def error_readSens(self, tag, result):
        self.setStyleStBarSens('error')
        # self.lbl_info_base.setText('{} {}'.format(tag, result))
        self.win_sens.statusbar.showMessage('{} {}'.format(tag, result))
        self.saveLogFile(tag, '{}'.format(result))

    def saveLogFile(self, status, msg):
        now = datetime.datetime.now()
        nam_f = str(now.day) + '-' + str(now.month) + '-' + str(now.year) + '.log'

        with open(nam_f, 'a') as f:
            str_w = now.strftime("%d-%m-%Y %H:%M:%S") + ' [' + status + '] ' + msg + '\n'
            f.write(str_w)

    def setStyleStBarSens(self, str_style):
        if str_style == 'info':
            self.win_sens.setStyleSheet("QStatusBar{padding-left:8px;background:rgba(0,0,0,0);"
                                        "color:green;font: 75 12pt \"Calibri\";font-weight:bold;}")
        if str_style == 'error':
            self.win_sens.setStyleSheet("QStatusBar{padding-left:8px;background:rgba(0,0,0,0);"
                                        "color:red;font: 75 12pt \"Calibri\";font-weight:bold;}")

    def error_write(self, result):
        print('Error write: {}'.format(result))
        self.logger.error(str(result))
    
    def finish_write(self):
        if not self.set_prg.flag_OkWrite:
            self.set_prg.flag_OkWrite = True
            print('WRITE OK')

    def finish_read(self, tag):
        print(tag + ' end')

    def querySaveFile(self, id, b_array):
        try:
            print('Query save file')
            print(b_array)
            self.signals.signalStopSens.emit()
            time.sleep(0.5)
            file_w = FileWriter(id, 48, b_array, self.client)
            file_w.signals.finish.connect(self.finishSaveFile)
            file_w.signals.percent.connect(self.showPercent)
            self.signals.signalCancelSaveFile.connect(file_w.cancelSave)
            self.pool_Exchange.start(file_w)

        except Exception as e:
            self.logger.error(str(e))

    def finishSaveFile(self):
        self.signals.signalStartSens.emit()

    def queryClearMemory(self):
        try:
            print('Clear Memory')
            byte_array = []
            for i in range(11):
                byte_array.append(0)

            self.signals.signalStopSens.emit()
            time.sleep(0.5)
            file_w = FileWriter(self.set_prg.selectID, 48, byte_array, self.client)
            file_w.signals.finish.connect(self.finishSaveFile)
            file_w.signals.percent.connect(self.showPercent)

            self.win_sens.progress_dialog.setWindowTitle("Стереть")
            self.win_sens.progress_dialog.setLabelText("Стирание памяти")
            self.win_sens.progress_dialog.setValue(0)

            self.win_sens.progress_dialog.show()

            self.signals.signalCancelSaveFile.connect(file_w.cancelSave)
            self.pool_Exchange.start(file_w)

        except Exception as e:
            self.logger.error(str(e))

    def nullSens(self):
        try:
            val_str = self.win_sens.tableCurrent.item(0, 1).text()
            val_d = int(val_str)
            val_regs = []
            b = pack('>i', val_d)
            byt = []
            byt.append(b[0])
            byt.append(b[1])
            val_d = int.from_bytes(byt, 'big', signed=False)
            val_regs.append(val_d)
            byt = []
            byt.append(b[2])
            byt.append(b[3])
            val_d = int.from_bytes(byt, 'big', signed=False)
            val_regs.append(val_d)
            self.writeValues(self.set_prg.selectID, 23, val_regs)
            self.set_prg.isEditableKoeff = False
            self.signals.signalReadKoeff.emit()

        except Exception as e:
            self.logger.error(str(e))

    def msgFromThread(self, msg_t):
        self.setStyleStBar('info')
        # self.lbl_info_base.setText('{}'.format(msg_t))
        self.win_view.statusbar.showMessage(msg_t)


class SettingsPrg(object):
    def __init__(self):
        config = configparser.ConfigParser()
        config.read("Settings.ini")

        self.port = 'COM' + config.get('ComPort', 'NumberPort')
        str_p = config.get('ComPort', 'PortSettings')
        dim_t = str_p.split(',')
        self.baudrate = int(dim_t[0])
        self.parity = dim_t[1]
        self.databits = int(dim_t[2])
        self.stopbits = int(dim_t[3])

        self.available_ports = self.scan_ports()
        self.isEditableKoeff = False
        self.isRunExchange = False
        self.selectType = 'none'
        self.sensTypes = []
        self.selectID = 0
        self.serial_numbers = []
        self.select_serial_number = 0
       
        self.send_1 = False

        for i in range(0, 33):
            self.sensTypes.append('none')
            self.serial_numbers.append(0)
        self.list_sens_id = []

        self.flag_OkWrite = False

        str_p = config.get('ComPort', 'Send_1_to_1015')
        if str_p.upper() == 'TRUE':
            self.send_1 = True
        str_p = config.get('ComPort', 'Timeout(ms)')
        self.timeout_s = int(str_p)

        str_p = config.get('ComPort', 'TimeBetweenSend(ms)')
        self.timeout_BS = int(str_p)

    def scan_ports(self):
        available = []
        for i in range(256):
            try:
                s = serial.Serial('COM'+str(i))
                available.append((s.portstr))
                s.close()
            except serial.SerialException:
                pass
        
        for s in available:
            print("%s" % (s))
        return available


def main():
    app = QtWidgets.QApplication(sys.argv)
    application = ApplicationWindow()
    application.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()

