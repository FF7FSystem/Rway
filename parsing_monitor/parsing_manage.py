import pasing_print
import pasing_verification_char
from termcolor import cprint
import json
import parse_monitor_config as CONFIG
import sys

def direct_order():
    """
    Оснавная функция упорядочивания получения статистических данных.
    TASK_NUM - номер задания. берется из файла конфига или из командной строки если этот файл запускался так
    (python parsing_manage.py 0001-0771). Оба варианта работа востребованы (из скрипта parse_monitor).
    1)prepare_data  - Получение списка задач в задании с некоторыми предварительными хар-ками, или по списку задач получения предварительных хар-к
    2)filtered_prepare_data - Из предварительного списка задач отфильровываются задачи запущенные несколько раз,
        (задачи запущенные по одному и тому же источнику, берутся задачи с максимальным количество спарсенных предложений)
    3)pasing_verification_char.main -  Запуск основной асинхронной функции парсинга статистических данных
    Результат формирования статистических данных записывается в файл в папке "parse_result", имя файла это номер
    задачи (формат JSON).
    4) Печать статистических данных в консоль через модуль pasing_print

    номер задания передается строкой типа '0001-0646'
    номера задач передается списком типа ['0001-0646-0001','0001-0646-0002','0001-0646-0003']
    :return: ничего не возвращает, но формирует файл
    """
    try:
        task_from_cmd = sys.argv[1] # Нати переданнные в командной строке аргументы при запуске (номер задания)
    except:
        task_from_cmd = False   # Если  В командной строке  номер задания не передавался
    TASK_NUM =  task_from_cmd if task_from_cmd else CONFIG.TASK_NUM # Если номер задания передан в командной строке,  иначе из конфига
    prepare_data = pasing_verification_char.task_list_prepare(task_num=TASK_NUM, for_connect=CONFIG.FOR_CONNECT)
    filtered_prepare_data = pasing_verification_char.excluding_task(prepare_data)
    pasing_verification_char.main(filtered_prepare_data, CONFIG.FOR_CONNECT)
    cprint('Файл с результатами парсинга:', 'yellow', end=' ')

    if CONFIG.PRINT_ALLTASK_STATISTIC: #Печать статистических данных в консоль через модуль pasing_print
        try:
            print(pasing_verification_char.file_path_of_result_data[0])
            with open(pasing_verification_char.file_path_of_result_data[0], 'r', encoding='utf-8') as fp:
                for_print = json.load(fp)   #Считывание файла
        except Exception as e:
            cprint(f'Ошибка открытия файла с результатами парсинга: {e}', 'red')
        pasing_print.main(for_print)    #печать через модуль

if __name__ == "__main__":
    direct_order()
