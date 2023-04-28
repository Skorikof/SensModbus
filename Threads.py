from PyQt5.QtCore import QObject, QRunnable, pyqtSignal, pyqtSlot
from struct import pack, unpack
import datetime
import time


class ReaderSignals(QObject):
    time_received = pyqtSignal(object)
    count_sensors = pyqtSignal(int, object)
    result = pyqtSignal(int, int, object)
    result_types = pyqtSignal(int, int, int, object)
    error_read = pyqtSignal(str, object)
    result_sens_current = pyqtSignal(int, object)
    result_sens_koeff = pyqtSignal(int, object)
    finish = pyqtSignal(str)
    error_readSens = pyqtSignal(str, object)
    editableKoeff = pyqtSignal()
    timeBaseTx = pyqtSignal(int)  # время между передачей от базовой станции
    result_SerialNumber = pyqtSignal(object)
    result_ValidNumbers = pyqtSignal(object)
    timeSensorTx = pyqtSignal(int)  # время между передачей от датчика
    msgToStatusBar = pyqtSignal(str)
    result_infoBase = pyqtSignal(object)  # чтение данных от базы (регистры 0-14)


class Reader(QRunnable):
    signals = ReaderSignals()

    def __init__(self, tag, dev_id, client, timeout_s, br_range):
        super(Reader, self).__init__()
        self.cycle = True
        self.is_run = False
        self.tag = tag
        self.ID = dev_id
        self.client = client
        self.list_id_current = []
        self.list_id_previous = []
        self.isReadKoeff = False
        self.query_readBaseTx = False
        self.query_readSensSerialNumber = True
        self.query_readValidNumbers = False  # запрос на чтение валидных номеров датчиков
        self.query_readSensorTx = False

        self.num_attempt = 0  # число попыток чтения
        self.MAX_ATTEMPT = 3
        self.flag_isRead = False
        self.timeout_BS = timeout_s/1000
        self.broadcast_range = br_range
        self.first_read = False

    @pyqtSlot()
    def run(self):
        while self.cycle:
            try:
                if not self.is_run:
                    time.sleep(0.01)
                else:
                    # чтение информации от базовой станции
                    if self.tag == 'read_base_info':
                        # читаем информацию по регистрам базы (с адреса=0 14 регистров)
                        if not self.broadcast_range:
                            self.num_attempt = 0
                            while self.num_attempt < self.MAX_ATTEMPT:
                                rr = self.client.read_holding_registers(0, 14, unit=1)
                                if not rr.isError():
                                    self.num_attempt = self.MAX_ATTEMPT+1
                                    self.signals.result_infoBase.emit(rr)
                                else:
                                    self.num_attempt += 1
                                    print('ATTEMPT: {}'.format(self.num_attempt))
                                    if self.num_attempt == self.MAX_ATTEMPT:
                                        self.signals.error_read.emit(self.tag, rr)
                                    if not self.timeout_BS == 0:
                                        time.sleep(self.timeout_BS)

                        if not self.broadcast_range:
                            # чтение информации по времени между передачами по запросу
                            if self.query_readBaseTx:
                                self.query_readBaseTx = False

                                self.num_attempt = 0
                                self.flag_isRead = False

                                while self.num_attempt < self.MAX_ATTEMPT:
                                    rr = self.client.read_holding_registers(5, 1, unit=1)
                                    if not rr.isError():
                                        self.num_attempt = self.MAX_ATTEMPT+1
                                        self.flag_isRead = True
                                        val = rr.registers[0]
                                        self.signals.timeBaseTx.emit(val)
                                    else:
                                        self.num_attempt += 1
                                        print('ATTEMPT: {}'.format(self.num_attempt))
                                        if self.num_attempt == self.MAX_ATTEMPT:
                                            self.signals.error_read.emit(self.tag, rr)
                                        if not self.timeout_BS == 0:
                                            time.sleep(self.timeout_BS)

                        if not self.broadcast_range:
                            # чтение информации по валидным номерам датчиков по запросу
                            if self.query_readValidNumbers:
                                self.query_readValidNumbers = False

                                self.num_attempt = 0
                                self.flag_isRead = False

                                while self.num_attempt < self.MAX_ATTEMPT:
                                    rr = self.client.read_holding_registers(4128, 32, unit=1)
                                    if not rr.isError():
                                        self.num_attempt = self.MAX_ATTEMPT+1
                                        self.flag_isRead = True
                                        self.signals.result_ValidNumbers.emit(rr)
                                    else:
                                        self.num_attempt += 1
                                        print('ATTEMPT: {}'.format(self.num_attempt))
                                        if self.num_attempt == self.MAX_ATTEMPT:
                                            self.signals.error_read.emit(self.tag, rr)
                                        if not self.timeout_BS == 0:
                                            time.sleep(self.timeout_BS)

                        # чтение датчиков в широковещательном диапазоне
                        if self.broadcast_range:
                            if not self.first_read:
                                self.first_read = True
                                self.list_id_current.clear()
                                for i in range(2, 32):
                                    self.signals.msgToStatusBar.emit('Scan ID: ' + str(i))
                                    self.num_attempt = 0
                                    self.flag_isRead = False
                                    while self.num_attempt < self.MAX_ATTEMPT:
                                        # чтение текущих данных от датчика
                                        rr = self.client.read_holding_registers(4102, 2, unit=i)
                                        if not rr.isError():                              
                                            self.num_attempt = self.MAX_ATTEMPT+1
                                            self.flag_isRead = True
                                            self.list_id_current.append(i)
                                            print('Append sensor: {}'.format(i))
                                        else:
                                            self.num_attempt += 1
                                            if self.num_attempt == self.MAX_ATTEMPT:
                                                print('no sensor: {}'.format(i))
                                                pass
                                            if not self.timeout_BS == 0:
                                                time.sleep(self.timeout_BS)

                                self.signals.msgToStatusBar.emit('')

                            if len(self.list_id_previous) != len(self.list_id_current):
                                self.query_readSensSerialNumber = True
                                self.list_id_previous = list(self.list_id_current)
                                self.signals.count_sensors.emit(len(self.list_id_current), self.list_id_current)
                                    
                            num_str = 0

                            if self.query_readSensSerialNumber:
                                self.query_readSensSerialNumber = False
                                for i in range(len(self.list_id_current)):
                                    num_str += 1
                                    print('ReadSensInfo')
                                    self.num_attempt = 0
                                    self.flag_isRead = False
                                    while self.num_attempt < self.MAX_ATTEMPT:
                                        rr = self.client.read_holding_registers(0, 11, unit=self.list_id_current[i])
                                        if not rr.isError():
                                            self.num_attempt = self.MAX_ATTEMPT+1
                                            self.flag_isRead = True
                                            self.signals.result_types.emit(num_str, self.list_id_current[i], rr)
                                        else:
                                            self.num_attempt += 1
                                            print('ATTEMPT: {}'.format(self.num_attempt))
                                            if self.num_attempt == self.MAX_ATTEMPT:
                                                self.signals.error_read.emit(self.tag, rr)
                                            if not self.timeout_BS == 0:
                                                time.sleep(self.timeout_BS)

                            num_str = 0
                            for i in range(len(self.list_id_current)): 
                                self.num_attempt = 0
                                self.flag_isRead = False
                                num_str += 1
                                while self.num_attempt < self.MAX_ATTEMPT:
                                    # чтение текущих данных от датчика
                                    rr = self.client.read_holding_registers(4102, 15, unit=self.list_id_current[i])
                                    if not rr.isError():
                                        self.num_attempt = self.MAX_ATTEMPT + 1
                                        self.flag_isRead = True
                                        self.signals.result.emit(num_str, self.list_id_current[i], rr)
                                    else:
                                        self.num_attempt += 1
                                        print('ATTEMPT: {}'.format(self.num_attempt))
                                        if self.num_attempt == self.MAX_ATTEMPT:
                                            self.signals.error_read.emit(self.tag, rr)
                                        if not self.timeout_BS == 0:
                                            time.sleep(self.timeout_BS)

                        # чтение датчиков в обычном режиме
                        else:
                            self.num_attempt = 0
                            self.flag_isRead = False

                            while self.num_attempt < self.MAX_ATTEMPT:
                                # чтение информации о количестве подключенных датчиков от базовой станции
                                rr = self.client.read_holding_registers(4162, 2, unit=1)
                      
                                if not rr.isError():                              
                                    self.num_attempt = self.MAX_ATTEMPT+1
                                    self.flag_isRead = True
                                    str_bin = self.decToBinStr(rr.registers[0]) + self.decToBinStr(rr.registers[1])
                                    str_bin = ''.join(reversed(str_bin))
                                    list_bin = []
                                    for i in str_bin:
                                        list_bin.append(int(i))
                                else:
                                    self.num_attempt += 1
                                    print('ATTEMPT: {}'.format(self.num_attempt))
                                    if self.num_attempt == self.MAX_ATTEMPT:
                                        self.signals.error_read.emit(self.tag, rr)
                                    if not self.timeout_BS == 0:
                                        time.sleep(self.timeout_BS)

                            if self.flag_isRead:
                                self.num_attempt = 0
                                self.flag_isRead = False
                                while self.num_attempt < self.MAX_ATTEMPT:
                                    # чтение информации о статусе подключенных датчиков
                                    rr = self.client.read_holding_registers(4164, 2, unit=1)
                                    # заполнение списка id датчиков, id=1 - базовая станция
                                    if not rr.isError():
                                        self.num_attempt = self.MAX_ATTEMPT+1
                                        self.flag_isRead = True

                                        str_bin = self.decToBinStr(rr.registers[0]) + self.decToBinStr(rr.registers[1])
                                        str_bin = ''.join(reversed(str_bin))
                                        list_status = []
                                        for i in str_bin:
                                            list_status.append(int(i))

                                        self.list_id_current.clear()

                                        for i in range(2, len(list_bin)):
                                            if list_bin[i] == 1: # and list_status[i]==0:
                                                self.list_id_current.append(i)

                                        if len(self.list_id_previous) != len(self.list_id_current):
                                            self.query_readSensSerialNumber = True
                                            self.list_id_previous = list(self.list_id_current)
                                            self.signals.count_sensors.emit(len(self.list_id_current),
                                                                            self.list_id_current)
                                    else:
                                        self.num_attempt += 1
                                        print('ATTEMPT: {}'.format(self.num_attempt))
                                        if self.num_attempt == self.MAX_ATTEMPT:
                                            self.signals.error_read.emit(self.tag, rr)
                                        if not self.timeout_BS == 0:
                                            time.sleep(self.timeout_BS)

                            if self.flag_isRead:
                                self.num_attempt = 0
                                self.flag_isRead = False
                                while self.num_attempt < self.MAX_ATTEMPT:
                                    # чтение текущего времени от базовой станции
                                    rr = self.client.read_holding_registers(4160, 2, unit=1)
                                    if not rr.isError():
                                        self.num_attempt = self.MAX_ATTEMPT+1
                                        self.flag_isRead = True
                                        num_second = unpack('L', pack('=HH', rr.registers[1], rr.registers[0]))[0]
                                        timestamp = self.calcTimeFromSecondsStartLinux(num_second)
                                        self.signals.time_received.emit(timestamp)
                                    else:
                                        self.num_attempt += 1
                                        print('ATTEMPT: {}'.format(self.num_attempt))
                                        if self.num_attempt == self.MAX_ATTEMPT:
                                            self.signals.error_read.emit(self.tag, rr)
                                        if not self.timeout_BS == 0:
                                            time.sleep(self.timeout_BS)

                            if self.flag_isRead:
                                num_str = 0
                                # опрос датчиков из списка базовой станции
                                if len(self.list_id_current) > 0:

                                    if self.query_readSensSerialNumber:
                                        self.query_readSensSerialNumber = False
                                        for i in range(len(self.list_id_current)):
                                            num_str += 1
                                            print('ReadSensInfo')
                                            self.num_attempt = 0
                                            self.flag_isRead = False
                                            while self.num_attempt < self.MAX_ATTEMPT:
                                                rr = self.client.read_holding_registers(0, 11, unit=self.list_id_current[i])
                                                if not rr.isError():
                                                    self.num_attempt = self.MAX_ATTEMPT+1
                                                    self.flag_isRead = True
                                                    self.signals.result_types.emit(num_str, self.list_id_current[i],
                                                                                   list_status[i], rr)
                                                else:
                                                    self.num_attempt += 1
                                                    print('ATTEMPT: {}'.format(self.num_attempt))
                                                    if self.num_attempt == self.MAX_ATTEMPT:
                                                        self.signals.error_read.emit(self.tag, rr)
                                                    if not self.timeout_BS == 0:
                                                        time.sleep(self.timeout_BS)

                                    num_str = 0
                                    for i in range(len(self.list_id_current)): 
                                        self.num_attempt = 0
                                        self.flag_isRead = False
                                        num_str += 1
                                        while self.num_attempt < self.MAX_ATTEMPT:
                                            # чтение текущих данных от датчика
                                            rr = self.client.read_holding_registers(4102, 19, unit=self.list_id_current[i])
                                            if not rr.isError():
                                                self.num_attempt = self.MAX_ATTEMPT+1
                                                self.flag_isRead = True
                                                self.signals.result.emit(num_str, self.list_id_current[i], rr)
                                            else:
                                                self.num_attempt += 1
                                                print('ATTEMPT: {}'.format(self.num_attempt))
                                                if self.num_attempt == self.MAX_ATTEMPT:
                                                    self.signals.error_read.emit(self.tag, rr)
                                                if not self.timeout_BS == 0:
                                                    time.sleep(self.timeout_BS)

                    # чтение информации от выбранного датчика
                    if self.tag == 'read_sensor':
                        # запрос на чтение интервала между передачами
                        if self.query_readSensorTx:
                            self.query_readSensorTx = False
                            self.num_attempt = 0
                            self.flag_isRead = False
                            while self.num_attempt < self.MAX_ATTEMPT:
                                rr = self.client.read_holding_registers(5, 1, unit=self.ID)
                                if not rr.isError():
                                    self.num_attempt = self.MAX_ATTEMPT+1
                                    self.flag_isRead = True
                                    val = rr.registers[0]
                                    self.signals.timeSensorTx.emit(val)
                                else:
                                    self.num_attempt += 1
                                    print('ATTEMPT: {}'.format(self.num_attempt))
                                    if self.num_attempt == self.MAX_ATTEMPT:
                                        self.signals.error_readSens.emit(self.tag + ' ID=' + str(self.ID), rr)
                                    if not self.timeout_BS == 0:
                                        time.sleep(self.timeout_BS)

                        # запрос на чтение серийного номера
                        if self.query_readSensSerialNumber:
                            self.query_readSensSerialNumber = False
                            self.num_attempt = 0
                            self.flag_isRead = False
                            while self.num_attempt < self.MAX_ATTEMPT:
                                rr = self.client.read_holding_registers(0, 11, unit=self.ID)
                                if not rr.isError():
                                    self.num_attempt = self.MAX_ATTEMPT+1
                                    self.flag_isRead = True
                                    self.signals.result_SerialNumber.emit(rr)
                                else:
                                    self.num_attempt += 1
                                    print('ATTEMPT: {}'.format(self.num_attempt))
                                    if self.num_attempt == self.MAX_ATTEMPT:
                                        self.signals.error_readSens.emit(self.tag + ' ID=' + str(self.ID), rr)
                                    if not self.timeout_BS == 0:
                                        time.sleep(self.timeout_BS)

                        self.num_attempt = 0
                        self.flag_isRead = False
                        while self.num_attempt < self.MAX_ATTEMPT:
                            # читаем текущие значения
                            rr = self.client.read_holding_registers(4102, 15, unit=self.ID)
                            if not rr.isError():
                                self.num_attempt = self.MAX_ATTEMPT + 1
                                self.flag_isRead = True
                                self.signals.result_sens_current.emit(self.ID, rr)
                            else:
                                self.num_attempt += 1
                                print('ATTEMPT: {}'.format(self.num_attempt))
                                if self.num_attempt == self.MAX_ATTEMPT:
                                    self.signals.error_readSens.emit(self.tag + ' ID=' + str(self.ID), rr)
                                if not self.timeout_BS == 0:
                                    time.sleep(self.timeout_BS)

                        if not self.isReadKoeff:
                            self.num_attempt = 0
                            self.flag_isRead = False
                            while self.num_attempt < self.MAX_ATTEMPT:
                                # читаем коэффициенты
                                rr = self.client.read_holding_registers(23, 5, unit=self.ID)
                                if not rr.isError():
                                    self.num_attempt = self.MAX_ATTEMPT+1
                                    self.flag_isRead = True
                                    self.isReadKoeff = True
                                    self.signals.result_sens_koeff.emit(self.ID, rr)
                                    self.signals.editableKoeff.emit()
                                else:
                                    self.num_attempt += 1
                                    print('ATTEMPT: {}'.format(self.num_attempt))
                                    if self.num_attempt == self.MAX_ATTEMPT:
                                        self.signals.error_readSens.emit(self.tag + ' ID=' + str(self.ID), rr)
                                    if not self.timeout_BS == 0:
                                        time.sleep(self.timeout_BS)

            except Exception as e:
                time.sleep(0.3)
                if self.tag == 'read_base_info':
                    print('Error base info!\n{}'.format(e))

                if self.tag == 'read_sensor':
                    print('Error read sensor!\n{}'.format(e))

    def startThread(self):
        self.is_run = True

    def stopThread(self):
        self.is_run = False

    def exitThread(self):
        self.cycle = False

    def readKoeff(self):
        self.isReadKoeff = False

    def readBaseTx(self):
        self.query_readBaseTx = True

    def readSensTx(self):
        self.query_readSensorTx = True

    def readSensInfo(self, val_bool):
        self.query_readSensSerialNumber = val_bool

    def readValidNumbers(self, val_bool):
        self.query_readValidNumbers = val_bool

    def changeBroadCastRange(self, val_bool):
        self.broadcast_range = val_bool
        self.list_id_previous.clear()
        self.first_read = False

    def decToBinStr(self, val_d):
        bin_str = bin(val_d)
        bin_str = bin_str[2:]
        bin_str = bin_str.zfill(16)
        return bin_str

    def calcTimeFromSecondsStartLinux(self, sec):
        timestamp = datetime.datetime(year=1970, month=1, day=1, hour=0, minute=0, second=0)
        d = datetime.timedelta(seconds=sec)
        new_timestamp = timestamp+d
        return new_timestamp

    def changeID(self, new_id):
        print('Thread change ID')
        self.ID = new_id


class WriterSignals(QObject):
    error_write = pyqtSignal(object)
    finish = pyqtSignal()


class Writer(QRunnable):
    signals = WriterSignals()
    
    def __init__(self, dev_id, start_addr, values, client):
        super(Writer, self).__init__()
        self.id = dev_id
        self.start_addr = start_addr
        self.write_values = values
        self.client = client
        self.number_attempts = 0
        self.max_attempts = 4
        self.flag_write = False
        self.flag_OK = False

    @pyqtSlot()
    def run(self):
        try:
            while self.number_attempts <= self.max_attempts:
                rr = self.client.write_registers(self.start_addr, self.write_values, unit=self.id)
                if not rr.isError():
                    self.flag_write = True
                    self.number_attempts = self.max_attempts+1
                    if not self.flag_OK:
                        self.signals.finish.emit()
                        self.flag_OK = True
                else:
                    print('Attempts {}'.format(self.number_attempts))
                    self.number_attempts += 1
                    time.sleep(0.2)
            if not self.flag_write:
                self.signals.error_write.emit(rr)
        except Exception as e:
            self.signals.error_write.emit(e)


class FileWriterSignals(QObject):
    error_write = pyqtSignal(object)
    finish = pyqtSignal()
    percent = pyqtSignal(int)


class FileWriter(QRunnable):
    signals = FileWriterSignals()
    
    def __init__(self, dev_id, start_addr, values, client):
        super(FileWriter,self).__init__()
        self.id = dev_id
        self.start_addr = start_addr
        self.write_values = values
        self.client = client
        self.number_attempts = 0
        self.max_attempts = 4
        self.flag_write = False
        self.flag_OK = False
        self.flag_Cancel_Save = False

    @pyqtSlot()
    def run(self):
            try:
                address = self.start_addr
                i = 0
                while i < len(self.write_values):
                    if self.flag_Cancel_Save:
                        self.flag_Cancel_Save = False
                        i = len(self.write_values) + 1
                    else:
                        temp_v = []
                        temp_v.append(self.write_values[i])
                        rr = self.client.write_registers(address, temp_v, unit=self.id)
                        if not rr.isError():
                            print('Ok {}'.format(i))
                            p = int((i+1)*100/len(self.write_values))
                            self.signals.percent.emit(p)
                            self.number_attempts=0
                            i += 1
                            address += 1
                        else:
                            self.number_attempts += 1
                            if self.number_attempts > self.max_attempts:
                                print('Error {}'.format(rr))
                                i = len(self.write_values) + 1
                            else:
                                pass
                self.signals.finish.emit()

            except Exception as e:
                print('Global Error {}'.format(e))

    def cancelSave(self):
        self.flag_Cancel_Save = True
