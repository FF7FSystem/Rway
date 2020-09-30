import pasing_verification_char
import json
import time
from datetime import datetime
import parse_monitor_config as CONFIG


class Stak():
    def __init__(self):
        self.offers = []

    def add(self, val):
        if val not in self.offers:
            if len(self.offers) < 1000:
                self.offers.append(val)
            else:
                self.offers.pop(0)
                self.offers.append(val)

def refresh_monitor_data():
    """
    В бесконечном цикле осуществляется считывание  из базы SQL
        1)prepare_data  - Получение списка задач в задании с некоторыми предварительными хар-ками, или по списку задач получения предварительных хар-к
        2)filtered_prepare_data - Из предварительного списка задач отфильровываются задачи запущенные несколько раз,
    filtered_prepare_data это список словарей с данными по задачам (источникам), а именно, номер задачи,
    когда началась задача, когда окончиласт и сколько собрала предложений (еще несколько характеристик). Идея в том,
    что пока каждая задача парсится (это происходит  в течении нескольких дней) получая данные о количестве собранных
    предложений по каждой задачи с момента  запуска парсинга  с периодом, Например 5 минут, я могу отследить динамику
    получения предложений по каждой задачи из интернета и построить график.
    Для формирования списка сначенией хар "Вссего предложений" в разное время выше создан класс стека, это список,
    который содержит не более 999 значений. Стек цикличный - Новые значения  добавляются в конец, старые значения
    удаляются в начале.
    Все новые значения полученные по задаче  добавляются в стек. Если по данной задаче мы уже получали данные
    (они сохранены в файле file_save_name), тогда в стек, в первую очередь добавляются ранее полученные данные,
    потом добавляются свежие из запроса  к sql.

    data_cnt_offers  - словарь, Ключ - номер задачи (типа 0001-0003-0005), значение - список сласса Stak.
    data_cnt_offers - создан чтобы:
        1) отслеживать какие задачи мы уже обработали а какие нет
        2)  Внего добавляется номер задачи 9ключ), и Стек с данными
        3) В последующем  методом сопоставления в списке словарей filtered_prepare_data, Хар "всего предложений",
        которая представлена 1 числом, заменяется на список (СТЕК) и записывается в файл.

    CONFIG.TASK_NUM - Номер задачи
    FILE_SAVE_NAME - имя файла для сохранения данных
    TIME_TO_REFRESH - в секудах до следующего обновления
    """
    data_cnt_offers = {}    #Ключ, номер задачи, значение, список сласса Stak
    with open(CONFIG.FILE_SAVE_NAME, 'r', encoding='utf-8') as fp: #для проверки, продолжаем дописывать данные по задаче или начинаем новые
        try:
            was_in_data = json.load(fp)
        except json.decoder.JSONDecodeError as e:
            was_in_data = {}
            print(f'Неудалось открыть файл {CONFIG.FILE_SAVE_NAME} или он пустой, будет перезаписан, {e}')
    while True:
        prepare_data = pasing_verification_char.task_list_prepare(task_num=CONFIG.TASK_NUM, for_connect=CONFIG.FOR_CONNECT) #Получение данных по задачам из SQL
        if not prepare_data: raise ValueError(f"Задача {CONFIG.TASK_NUM} отсутствует в 1С")    #Если данных нет, кидаем исключение (подробнее в модуле)
        filtered_prepare_data = pasing_verification_char.excluding_task(prepare_data, False, False)    #Фильтруются данные, на случай если неоднократного запуска задач (подробнее в модуле)

        f_continue_task = True if was_in_data and CONFIG.TASK_NUM in was_in_data[0]['Код задачи'] else False #Флаг, продолжить файл или начать заново если в файле не текущая задача
        for item in filtered_prepare_data:
            if not item['Код задачи'] in data_cnt_offers and not f_continue_task:   #Если в словаре data_cnt_offers нет текущей задачи и раньше мы не собирали данные по этой задаче
                data_cnt_offers[item['Код задачи']] = Stak()                        #Создать ключ типа "0001-0003-0005", значение Стек
            elif not item['Код задачи'] in data_cnt_offers and f_continue_task:     #Если в словаре data_cnt_offers нет текущей задачи и ранее мы собирали данные по этой задаче
                for old_item in was_in_data:                                        #беребрать все задачи которые собирали раньше
                    if old_item['Код задачи']==item['Код задачи']:                  #Найти соответствующие задачи в новом парсинге
                        data_cnt_offers[item['Код задачи']] = Stak()                #Инициализировать Стек
                        for num in  old_item['Всего предложений']:                  #перебрать все значения "Всего предложений" которые мы собирали ранее
                            data_cnt_offers[item['Код задачи']].add(num)            # и добавить их в Стек
                else:
                    data_cnt_offers[item['Код задачи']] = Stak()

            data_cnt_offers[item['Код задачи']].add(item['Всего предложений'])      # В случаи когда в словаре data_cnt_offers есть  текущая  задача  просто дополняем Стек новымиданными

        for i in filtered_prepare_data: #В ранее полученном списке словарей содержащих данные по каждой задаче
            i['Всего предложений'] = data_cnt_offers[i['Код задачи']].offers # заменяем позицию "Всего предложений" на обновленный список Стек
            i['Статус выполнения'] = '4' if i['Статус выполнения'] == 'None' else i['Статус выполнения'] #Если задание создано, но парсинг не запущен, приведение к словарному значению (числу), а не значению None (и тот строкой)
        with open(CONFIG.FILE_SAVE_NAME, 'w',encoding='utf-8') as fp:
            json.dump(filtered_prepare_data, fp, ensure_ascii=False, indent=4)
        if CONFIG.PRINT_MONITOR_DATA_GO:
            print('проход', time.strftime("%H:%M:%S %d.%m.%Y", time.localtime()))
        time.sleep(CONFIG.TIME_TO_REFRESH) #пауза перед стледующим обращением в базу за новыми  данными парсинга

if __name__ == '__main__':
    refresh_monitor_data()
