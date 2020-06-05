from termcolor import cprint
import pandas as pd
import os
import statistics


def print_char_dict(char_dict):
    """
    Функция печатает в терминал данные из словаря в красивом формате
    :param char_dict: Словарь с результатами оценки заполняемости каждой задачи (источника)
    :return: ничего не возвращает
    """
    cprint(('{:_^40}'.format(char_dict['Псевдоним'])), 'magenta')
    #Заголовки headers будут выводится в болом цвете
    headers = ['Псевдоним', 'Код задачи', 'Всего предложений', 'Дата начала', 'Дата окончания', 'Статус выполнения']
    for key in headers:
        print('{:_<40}'.format(key), end=' ')
        cprint(char_dict[key], 'white')

    char_list = [[i, char_dict[i]] for i in char_dict if i not in headers]  #исключение из словаря заголовков отпечатаных выше
    char_list = sorted(char_list, key=lambda x: float(x[1]), reverse=True)  #сортировка словаря в список по убыванию процента заполняемости,
    # чтобы сначала печатались полностью заполненные характеристики в зеленом цвете а потом уж незаполненые в красном цвете
    for val in char_list:
        term= 'green' if float(val[1]) >= 99 else 'red'
        print('{:_<40}'.format(val[0]), end=' ')
        cprint(val[1], term)
    print()

def prepare_comparison_data():
    os.chdir(r'C:\Users\user1\AppData\Roaming\JetBrains\PyCharmCE2020.1\scratches\parse_resilt_items')
    files = sorted(os.listdir(), reverse=True)
    data = ''
    for file in files:
        file_name, extension = os.path.splitext(file)
        if extension == '.json':
            # print(file)
            inp = pd.read_json(file, encoding='utf8')
            if isinstance(data, pd.core.frame.DataFrame):
                data = data.append(inp, sort=True)
            else:
                data = pd.read_json(file, encoding='utf8')

    # sorted(pd.read_json(r'C:\Users\user1\AppData\Roaming\JetBrains\PyCharmCE2020.1\scratches\parse_resilt_items\async_items_(2020-04-22 17-46-42).json', encoding='utf8').columns))
    need_cols = ['Адрес', 'Всего предложений', 'Дата Сбора Информации', 'Дата начала', 'Дата окончания', 'ДатаРазмещения', 'Долгота По Источнику', 'Заголовок', 'Класс Объекта', 'Код задачи', 'Назначение Объекта Предложения', 'Общая Площадь Предложения', 'Описание', 'Подсегмент', 'Предмет Сделки', 'Продавец', 'Псевдоним', 'Размерность Площади', 'Размерность Стоимости', 'СамаяРанняяДатаРазмещения', 'Сегмент', 'Способ Реализации', 'СсылкаИсточника', 'Статус выполнения', 'Субъект', 'Телефон Продавца', 'Тип объекта недвижимости', 'ТипРынка', 'ТипСделки', 'Цена', 'Цена Предложения За 1 кв.м.', 'Широта По Источнику', 'Этаж', 'Этажность']
    for ser in data:
        if ser not in need_cols:
            data = data.drop(ser, axis=1)
    return data

def statist_data(data,source):

    data = data.query(f'Псевдоним == "{source}"')
    headers = ['Псевдоним', 'Код задачи', 'Дата начала', 'Дата окончания', 'Статус выполнения']
    for_stat = {}
    for elem in data:
        if elem not in headers:
            elem_data = [i for i in data[elem] if i > 0]
            if elem_data:
                val = f'{statistics.mean(elem_data):.2f}'
            else:
                val = 0
            for_stat[elem] = val
    return for_stat

def print_char_dict_comparison(char_dict,comparis_dict):
    """
    Функция печатает в терминал данные из словаря в красивом формате
    :param char_dict: Словарь с результатами оценки заполняемости каждой задачи (источника)
    :return: ничего не возвращает
    """
    cprint(('{:_^40}'.format(char_dict['Псевдоним'])), 'magenta')
    #Заголовки headers будут выводится в болом цвете
    headers = ['Псевдоним', 'Код задачи', 'Всего предложений', 'Дата начала', 'Дата окончания', 'Статус выполнения']
    for key in headers:
        print('{:_<40}'.format(key), end=' ')
        if key in comparis_dict:
            term = 'green' if char_dict[key] >= float(comparis_dict[key]) else 'red'
            cprint(char_dict[key], term,end=' ')
            cprint(f'({int(float(comparis_dict[key]))})', 'white')
        else:
            cprint(char_dict[key], 'white')
    char_list = [[i, char_dict[i]] for i in char_dict if i not in headers]  #исключение из словаря заголовков отпечатаных выше
    char_list = sorted(char_list, key=lambda x: float(x[1]), reverse=True)  #сортировка словаря в список по убыванию процента заполняемости,
    # чтобы сначала печатались полностью заполненные характеристики в зеленом цвете а потом уж незаполненые в красном цвете
    for val in char_list:
        term = 'green' if float(val[1])  >= float(comparis_dict[val[0]]) or float(val[1]) == 100 else 'red'
        cprint('{:_<40}'.format(val[0]), term, end=' ')
        cprint(f'{val[1]:>6}', term,end=' ')
        cprint('{:>7}'.format('('+str(float(comparis_dict[val[0]]))+')'), term, end=' ')
        difference=float(comparis_dict[val[0]])-float(val[1])
        tolerance=2
        if -tolerance>difference or difference>tolerance:
            if term == 'green':
                cprint(f'+{difference*(-1):.2f}','magenta')
            else:
                cprint(f'-{difference :.2f}', 'magenta')
        else:
            print()
    print()

def main(for_output):
    # печать в терминал функцией print_char_dict в нужном формате  всех полученых результатов  по задачам из результирующего сортированого списка
    comparison_data=prepare_comparison_data()
    for i in sorted(for_output, key=lambda x: x['Псевдоним']):
        comparison_val=statist_data(comparison_data,i['Псевдоним'])
        print_char_dict_comparison(i,comparison_val)  #В функцию передается словарь, так как результирующий список for_output содержит список словарей

if __name__ == "__main__":
    val=[{'Цена': '92.48', 'Цена Предложения За 1 кв.м.': '92.36', 'Размерность Стоимости': '100.00', 'Продавец': '100.00', 'Телефон Продавца': '100.00', 'Общая Площадь Предложения': '99.96', 'Размерность Площади': '99.96', 'Широта По Источнику': '44.65', 'Долгота По Источнику': '44.65', 'Заголовок': '100.00', 'Способ Реализации': '100.00', 'Этаж': '59.12', 'Этажность': '42.08', 'Тип объекта недвижимости': '0.01', 'Предмет Сделки': '11.61', 'Класс Объекта': '0.00', 'Назначение Объекта Предложения': '100.00', 'Дата Сбора Информации': '100.00', 'Адрес': '100.00', 'ДатаРазмещения': '100.00', 'Описание': '100.00', 'СамаяРанняяДатаРазмещения': '100.00', 'Сегмент': '100.00', 'Подсегмент': '100.00', 'СсылкаИсточника': '100.00', 'Субъект': '100.00', 'ТипРынка': '100.00', 'ТипСделки': '100.00', 'Код задачи': '0001-0615-0014', 'Псевдоним': 'PROPER', 'Дата начала': '2020-04-19 11:22:49', 'Дата окончания': '2020-04-19 19:23:09', 'Всего предложений': 21977, 'Статус выполнения': '2'}]
    main(val)