import subprocess
import os
import signal
import sys
import time
import psutil
import pyodbc
import parse_monitor_config as CONFIG
import sql_querys as QUERYS
import re
import time
from termcolor import cprint


def start_app():
    """
    Запуск приложения "Монитор парсинга" в подпроцессе.
    :return:
    """
    proc = subprocess.Popen([sys.executable, r'parse_monitor.py'])
    cprint((f'Приложение запущено, ПИД процесса {proc.pid}', time.strftime("%H:%M:%S %d.%m.%Y", time.localtime())),
           'green')
    return proc.pid


def get_last_task():
    """
    Получение из базы данных номера последнего созданного задания парсинга
    :return: номер актуального задания
    """
    conn = pyodbc.connect(CONFIG.FOR_CONNECT)  # подключение к базе
    cursor = conn.cursor()                     #Создание курсора
    cursor.execute(QUERYS.NUM_LAST_PARS_TASK)   #выполнение запроса
    response = cursor.fetchall()                #ответ запроса в виде списка списков
    try:
        return response[0][0]                   #выделение номера задания из списка списков и возвращение
    except Exception as e:
        print("Задача парсинга в 1С не найдена, ошибка - >", e)


def refresh_config(new_num):
    """
    Запись в файл конфига нового номера задания
    :param new_num:Новый номер задания
    :return: Ничего не возвращает
    """
    with open(r'parse_monitor_config.py', 'r+', encoding='utf8') as conf_file:
        context = conf_file.read()  #Считывание собержимого файла в тектовом виде
        pat = r"(?<=TASK_NUM=')[\d-]+"  #Регулярка поиска старого номера задания
        try:
            context = re.sub(pat,new_num,context)  #Найти и заменить номер задания на новый
        except Exception as e:
            print('Номер задачи не найден в файле конфига', e)
        conf_file.seek(0)      #Сместить "коретку" записи на начало файла (чтобы не дописывать в конец дубюлируя инфу).
        conf_file.write(context) #Запись конфига
    cprint((f'Конфиг обновлен, новый номер задачи {new_num}', time.strftime("%H:%M:%S %d.%m.%Y", time.localtime())),
           'cyan')


def kill_proc_tree(pid, including_parent=True):
    """
    Функция убивает процесс и порожденные име подпроцессы
    :param pid: Пид процесса, который нужно убить
    :param including_parent:
    :return:
    """
    parent = psutil.Process(pid) #Процесс
    children = parent.children(recursive=True) #Подпроцессы основного процесса
    for child in children:
        child.kill()
    gone, still_alive = psutil.wait_procs(children, timeout=5)
    if including_parent:
        parent.kill()
        parent.wait(5)
    cprint((f'Приложение остановлено ПИД процесса {pid}', time.strftime("%H:%M:%S %d.%m.%Y", time.localtime())), 'red')


def go_go_go():
    """
    Основная функция.
    Запускает подпроцесс "монитор парсинга" (файл parse_monitor.py) по конфигурации parse_monitor_config.
    После этого в бесконечном цикле проверяет, не поменялся ли номер здания парсинга. Если поменялся, то вызывается
    функция убийства всего дерева процесса приведущего монитора, меняется номер задания парсинга в конфиге и монитор
    парсинга запускается заново с новым конфигом. Небольшая задержка  переод новым запросом номера нового задания.
    Повтор...
    :return: ничего не возвращает
    """
    current_pid = start_app() #Запуск монитора приложения монитора
    try: #Если я прерву задачу с клавиатуры или другим способом, то дерево процессов запущенного монитора убивается (иначу они продолжат функционировать)
        while current_pid:  #Если текущий процесс имет номер пид
            last_task = get_last_task() #Определить номер самого нового задания парсинга
            if last_task and last_task != CONFIG.TASK_NUM:  #Если номер Нового задания отличается от того, что указан в конфиге
                kill_proc_tree(current_pid) #Предидущее приложение монтирова и все вызванные им процессы убиваются
                refresh_config(last_task)   #В конфиге обновляется номер задания
                current_pid = start_app()   #Заново запускается приложение монитора
            time.sleep(CONFIG.TIME_FIND_NEW_NUM_TASK) #Время задержци - обновления
    except Exception:
        kill_proc_tree(current_pid)



if __name__ == '__main__':
    go_go_go()
