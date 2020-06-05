import pasing_print
import pasing_verification_char
from termcolor import cprint
import json


def direct_order():
    """
    task_num:
        номер задания передается строкой типа '0001-0646'
        номера задач передается списком типа ['0001-0646-0001','0001-0646-0002','0001-0646-0003']
    :return:

    """
    task_num = '0001-0661'
    # task_num = ['0001-0661-0036', '0001-0661-0035']
    for_connect = 'DRIVER={SQL Server};SERVER=10.199.13.57;DATABASE=rway;UID=vdorofeev;PWD=VD12345#' #Настройки подключения к базе RWAY, будет использовано для подключения в синхронной и асинхронной функциях
    # prepare_data  - Получение списка задач в задании с некоторыми предварительными хар-ками, или по списку задач получения предварительных хар-к
    prepare_data  = pasing_verification_char.task_list_prepare(task_num= task_num ,for_connect = for_connect)
    # pasing_verification_char.excluding_task  - Из предварительного списка задач отфильровываются задачи запущенные несколько раз
    # (задачи запущенные по одному и тому же источнику, берутся задачи с максимальным количество спарсенных предложений)
    # Там же выводится небольшая статистика по ззадачам и записывается в фал  небольшой лог (в котором указывается  сколько задач и с каким статусом на данный момент в данном задании)
    #pasing_verification_char.main -  Запуск основной асинхронной функции парсинга
    filtered_prepare_data=pasing_verification_char.excluding_task(prepare_data)
    pasing_verification_char.main(filtered_prepare_data,for_connect)
    cprint('Файл с результатами парсинга:','yellow',end=' '), print(pasing_verification_char.file_path_of_result_data[0])

    try:
        with open(pasing_verification_char.file_path_of_result_data[0], 'r', encoding='utf-8') as fp:
            for_print=json.load(fp)
    except Exception as e:
        cprint(f'Ошибка открытия файла с результатами парсинга: {e}','red')
    try:
        pasing_print.main(for_print)
    except Exception as e:
        cprint(f'Ошибка в модуле pasing_print: {e}','red')

r"""
    Обычно
    Логи в C:\Users\user1\Desktop\pasing_verification_RWAY\add_manage\logs
    Файл с результатами парсинга C:\Users\user1\Desktop\pasing_verification_RWAY\add_manage
    данные пути указаны в pasing_verification_char
"""

if __name__ == "__main__":
    direct_order()