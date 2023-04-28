import logging
from datetime import datetime
import os

_log_format = f"%(asctime)s - [%(levelname)s] - %(name)s - (%(filename)s).%(funcName)s(%(lineno)d) - %(message)s"
_date_log = str(datetime.now().day).zfill(2) + '_' + str(datetime.now().month).zfill(2) + \
            '_' + str(datetime.now().year)
_path_logs = '/log'


def get_warn_handler():
    warn_handler = logging.FileHandler(_path_logs + '/' + _date_log + '.log', encoding='utf-8')
    warn_handler.setLevel(logging.WARNING)
    warn_handler.setFormatter(logging.Formatter(_log_format))
    return warn_handler


def get_err_handler():
    err_formater = f'%(asctime)s - [%(levelname)s] - %(name)s - ' \
                   f'(%(filename)s).%(funcName)s(%(lineno)d) : %(message)s'
    err_handler = logging.FileHandler(_path_logs + '/' + _date_log + '.log', encoding='utf-8')
    err_handler.setLevel(logging.ERROR)
    err_handler.setFormatter(logging.Formatter(err_formater))
    return err_handler


def get_info_handler():
    info_handler = logging.FileHandler(_path_logs + '/' + _date_log + '.log', encoding='utf-8')
    info_handler.setLevel(logging.INFO)
    info_handler.setFormatter(logging.Formatter(_log_format))
    return info_handler


def get_logger(name):
    if not os.path.exists(_path_logs):
        access_rights = 0o755
        try:
            os.mkdir(_path_logs, access_rights)
        except OSError:
            print("Создать директорию %s не удалось" % _path_logs)

    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    logger.addHandler(get_warn_handler())
    logger.addHandler(get_err_handler())
    logger.addHandler(get_info_handler())
    return logger
