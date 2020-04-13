import pyodbc   #для связи с бд в синхронном виде
import aioodbc  #для связи с бд в асинхронном виде
import time
import datetime
import asyncio
from termcolor import cprint
import json


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
        term= 'green' if float(val[1]) >=99 else 'red'
        print('{:_<40}'.format(val[0]), end=' ')
        cprint(val[1], term)
    print()

def task_list_prepare(task_num,for_connect):
    """
    Данная задача  при получении  номера задания task_num формирует список словарей содержащих данные по каждой задаче для
    дальнейшей передачи в асанхронную функцию и получения оценочных данных

    :param task_num:    Номер ЗАДАНИЯ
    :param for_connect: Настройки для подключения к базе данных
    :return: Возвращает список словарей содержащих данные по каждой задаче
    """
    task_list_prepare=[]            #активация итогового списка
    conn = pyodbc.connect(for_connect)  #подключение к базе
    cursor = conn.cursor()              #подключение к базе
    #Запрос к базе для получения списка всех задач и некоторых характеристик (_IDRRef,ДатаНачала,ДатаОкончания,Код задачи,Псевдоним,Статус,Шаблон)
    cursor.execute(f'''
    SELECT _IDRRef,ДатаНачала,ДатаОкончания,_Fld198 as 'Код задачи' ,Псевдоним, jour.Статус,_shablon.Наименование as 'Шаблон'
    FROM [rway].[dbo].[_Task62] task
    left join [rway].[Документ].[Задание] doc on task._Fld192RRef = doc.Ссылка and task._Fld231RRef=0xAB47107B44B1669011E990CEFFC204B6
    left join [rway].[Справочник].[Источники] sourc on task._Fld197RRef = sourc.Ссылка
    left join [rway].[РегистрСведений].[ЖурналЗадач] jour on task._IDRRef = jour.Задача
	left join [rway].[Справочник].[ШаблоныИмпорта] _shablon on task._Fld214RRef = _shablon.НастройкаПСОД
    where doc.КодЗадания = '{task_num}'
         ''') #and jour.Статус = 2 добавить для получения данных только о выполненых задачах 

    list_task= cursor.fetchall()    #Результаты выполненого запроса, список, где хар-ки (_IDRRef,ДатаНачала,ДатаОкончания,Код задачи,Псевдоним,Статус,Шаблон) указаны позиционно
    for item in list_task:          #перебираем каждую строку, добавляем хар-ку "всего предложений",создаме словарь, чтобы обращаться к хар-кам не позицонно а по ключу
        # Для получения данных о количестве всех предложений по каждой задаче необходимо отредатировать хар _IDRRef
        # Преобразовать _IDRRef из байтового представления в целочисленное для подстановки в селект
        code_of_task='0x'+item[0].hex().upper()
        #запрос в базу для получения количества предложений по данной задаче
        cursor.execute(
            f'''Select COUNT(Предложение) from [rway].[РегистрСведений].[ПредложенияЗадач] where Задача= {code_of_task}''')
        quantity=cursor.fetchall()[0][0]    #количеств предложений по данной задаче
        psevdonim=item[4]                   #псевдоним задачи или название источника, например "АВИТО"

        sub = ''
        if psevdonim == 'MOVE':         # есть задачи у которых одинаковый псевдоним, это значит что при выводе данных будет не понятно что это за источник и почему его несколько
            if 'ГорМил' in item[6]:     #Например задачи MOVE имеют одинаковый псевдоним но шаблоны инструкции по которым осуществляется парсинг
                sub = '_ГОРМИЛ'         #у них разные по этому в этом участке проверяется, если псевдоним МУВ, то к псевдониму добавляется
            elif 'Регионы' in item[6]:  #имя региона по которому осуществляется парсинг и указанному в имени шаблона
                sub = '_РЕГИОНЫ'        #так создается 3 источника с разными псевдонимами
            elif 'Москва' in item[6]:
                sub = '_МОСКВА'
        #Запись в результирующий список словаря с характеристиками по каждой задаче   {'Код задачи','Псевдоним','Дата начала','Дата окончания','Всего предложений','Статус выполнения')
        #'Статус выполнения':str(item[5]) данная характеристика преобразуется в строку, так как база данных возвращает ответ типа declare(2),
        # но на языке питона это вызов функции declare, которой тут , конечно нет. В результате тв данную хар-ку записывается число (1,2 или 3)
        task_list_prepare.append({'Код задачи':item[3],'Псевдоним':psevdonim+sub,'Дата начала':item[1],'Дата окончания':item[2],'Всего предложений':quantity,'Статус выполнения':str(item[5])})
    cursor.close()  #Закрытие подключения к базе
    conn.close()    #Закрытие подключения к базе
    return(task_list_prepare)

async def task_full_info(prepare_data,for_connect):
    """
    Основная функция (корутина) получения оценочных данных по каждой задаче.
    :param prepare_data: Словарь с данными о задаче полученых ранее  {'Код задачи','Псевдоним','Дата начала','Дата окончания','Всего предложений','Статус выполнения')
    :param for_connect: Настройки для подключения
    :return:  Возвращает список словарей с полными оценочными данными о заполнямости задачи (источника)
    """
    task_code=prepare_data['Код задачи']
    print(f'задача {task_code} начата')
    task_beg=time.time()
    conn = await aioodbc.connect(dsn=for_connect)   #Подключение к базе данных через асинхронный коннектор
    cursor = await conn.cursor()                    #Подключение к базе данных через асинхронный курсор
    '''Запросы SQL возвращают 1 строку без названий колонок, по этому приходится создавать списки названий колонок
    col_name с последующим создаванием словарей (Ключ = Имя колонки, значение = значение), путем  сопоставления двух списков
    колонок и ключей позиционно, что неочень хорошо.
    '''
    col_name_char1=(
        'Цена','Цена Предложения За 1 кв.м.','Размерность Стоимости','Продавец','Телефон Продавца','Общая Площадь Предложения',
        'Размерность Площади','Широта По Источнику','Долгота По Источнику','Заголовок','Способ Реализации','Этаж','Этажность',
        'Тип объекта недвижимости','Предмет Сделки','Класс Объекта','Назначение Объекта Предложения','Дата Сбора Информации')
    col_name_char2 = (
        'Адрес','ДатаРазмещения','Описание','СамаяРанняяДатаРазмещения','Сегмент','Подсегмент','СсылкаИсточника','Субъект',
        'ТипРынка','ТипСделки')
    try:
        try:
            #Запрос из бызы характеристик указанных в списке col_name_char1. Самый долгий запрос который блокирует обычную функцию на несколько минут
            #Сначала запрос скрепляет несколько таблиц разной формы (select Объект, Наименование,Значение). Одна из таблиц горизонтальная и для одного предложения
            #содержит одну строку и много колонок, другая таблица вертикальная и для одного предложения содержит одну колонку с множеством  строк
            #Чтобы создать одну таблицу характеристик предложения с одной строкой и множеством колонок используется конструкция  case
            #при этом,содержание колонок не учитывается, а проставляется 1 если ячейка заполнена и 0 если пустая (тут не проверяется качество заполнения, но проверяется количество (полнота))
            # Получается большая таблица со всеми предложениями  (строка) и их характеристиками (столбцы ) для данной задачи (источника) где все характеристики заполнены 1 или 0
            # Следующим этапом в полученой  таблице  подситываются все заполненные поля

            await cursor.execute(f'''
            Select  count(case when Наименование = 'Цена' and Значение <> 0 then  1 end) 'Цена',
                    count(case when Наименование = 'Цена Предложения За 1 кв.м.' and Значение <> 0 then  1 end) 'Цена Предложения За 1 кв.м.',
                    count(case when Наименование = 'Размерность Стоимости' and Значение <> 0 then  1 end) 'Размерность Стоимости',
                    count(case when Наименование = 'Продавец' and Значение <> 0 then  1 end) 'Продавец',
                    count(case when Наименование = 'Телефон Продавца' and Значение <> 0 then  1 end) 'Телефон Продавца',
                    count(case when Наименование = 'Общая Площадь Предложения' and Значение <> 0 then  1 end) 'Общая Площадь Предложения',
                    count(case when Наименование = 'Размерность Площади' and Значение <> 0 then  1 end) 'Размерность Площади',
                    count(case when Наименование = 'Широта По Источнику' and Значение <> 0 then  1 end) 'Широта По Источнику',
                    count(case when Наименование = 'Долгота По Источнику' and Значение <> 0 then  1 end) 'Долгота По Источнику',
                    count(case when Наименование = 'Заголовок' and Значение <> 0 then  1 end) 'Заголовок',
                    count(case when Наименование = 'Способ Реализации' and Значение <> 0 then  1 end) 'Способ Реализации',
                    count(case when Наименование = 'Этаж' and Значение <> 0 then  1 end) 'Этаж',
					count(case when Наименование = 'Количество Этажей' and Значение <> 0 then  1 end) 'Этажность',
					count(case when Наименование = 'Тип объекта недвижимости' and Значение <> 0 then  1 end) 'Тип объекта недвижимости',
					count(case when Наименование = 'Предмет Сделки' and Значение <> 0 then  1 end) 'Предмет Сделки',
					count(case when Наименование = 'Класс Объекта' and Значение <> 0 then  1 end) 'Класс Объекта',
                    count(case when Наименование = 'Назначение Объекта Предложения' and Значение <> 0 then  1 end) 'Назначение Объекта Предложения',
                    count(case when Наименование = 'Дата Сбора Информации' and Значение <> 0 then  1 end) 'Дата Сбора Информации'
            From
            (SELECT Предложение FROM [rway].[dbo].[_Task62] task left join [rway].[РегистрСведений].[ПредложенияЗадач] predl on _IDRRef = predl.Задача where _Fld198 = '{task_code}') num_pred
            join
            (select Объект, Наименование,
                case
                    when Значение_Число > 0 then  1
                    when Значение_Строка <> '' then  1
                    when Значение_Дата is not null then  1
                    when Значение <> 0x00000000000000000000000000000000 then  1
                end Значение

                 from [rway].[РегистрСведений].[ЗначенияХарактеристик] Xap left join [rway].[ПВХ].[Характеристики] Xap_val on Xap.Характеристика = Xap_val.Ссылка) xap_tab_pred
            on num_pred.Предложение = xap_tab_pred.Объект
                        ''')
            result_sql_1= await cursor.fetchall()                   #Результаты запроса приходят в виде списка
            result_dict=dict(zip(col_name_char1, result_sql_1[0]))  #Создается словарь  где ключ это название хар-ки (col_name_char1), а значение это количество заполненыных предложений
        except Exception as e:
            cprint(f'Ошибка задачи {task_code} функциии task_full_info в запросе №1 = {e}', 'red')

        try:
            #Данный запрос похож на запрос выше, за исключеним того, что тут не нужно  соеденять таблицы 2-х форм
            # а результат аналогичный, но для  характеристик col_name_char2
            await cursor.execute(f'''
                SELECT	count(case when [АдресПредставление] <> '' then  1 end) 'Адрес',
                        count(case when [ДатаРазмещения] is not null then  1 end) 'ДатаРазмещения',
                        count(case when [Описание] <> '' then  1 end) 'Описание',
                        count(case when [СамаяРанняяДатаРазмещения] is not null then  1 end) 'СамаяРанняяДатаРазмещения',
                        count(case when [Сегмент] = 0x00000000000000000000000000000000 then null else  1 end) 'Сегмент',
						count(case when [Подсегмент] = 0x00000000000000000000000000000000 then null else  1 end) 'Подсегмент',
                        count(case when [СсылкаИсточника] <> '' then  1 end) 'СсылкаИсточника',
                        count(case when [Субъект] = 0x00000000000000000000000000000000 then null else  1 end) 'Субъект',
                        count(case when [ТипРынка] = 0x00000000000000000000000000000000 then null else  1 end) 'ТипРынка',
                        count(case when [ТипСделки] = 0x00000000000000000000000000000000 then null else  1 end) 'ТипСделки'
                FROM [rway].[dbo].[_Task62] task
                left join [rway].[Документ].[Задание] doc on task._Fld192RRef = doc.Ссылка and task._Fld231RRef=0xAB47107B44B1669011E990CEFFC204B6
                left join [rway].[Справочник].[Источники] sourc on task._Fld197RRef = sourc.Ссылка
                left join [rway].[РегистрСведений].[ЖурналЗадач] jour on task._IDRRef = jour.Задача
                left join [rway].[РегистрСведений].[ПредложенияЗадач] predl on _IDRRef = predl.Задача
                join [rway].[Справочник].[ПредложенияОбъектовНедвижимости] predl_val1 on Предложение = predl_val1.Ссылка
                where _Fld198 = '{task_code}' ''')

            result_sql_2 = await cursor.fetchall()                           #Результаты запроса приходят в виде списка
            result_dict.update(dict(zip(col_name_char2, result_sql_2[0])))   #Словарь  с характеристиками для задачи (созданный после первого запроса)  дополняется новыми характеристиками col_name_char2

        except Exception as e:
            cprint(f'Ошибка задачи {task_code} функциии task_full_info в запросе №2 = {e}', 'red')

        for key in result_dict:
            # Замена данных по характеристакам. Каждая характеристика представляется не как количество заполненных предложений
            # в источнике, а отношение всех общего количества к количеству заполненых (в процентах т.е. *100)
            #так для каждого источника получаются данные о процентном соотношении заполнености,
            # например  из 100 предложений по источнику АВИТО только у 90% предложений заполнена характеристика "ЦЕНА"
            if int(prepare_data['Всего предложений']) != 0:  # Дабы небыло деления на 0 в случае когда всего предложений  = 0
                val = result_dict[key] / int(prepare_data['Всего предложений']) * 100
            else:
                val = 0
            result_dict[key] = f'{val:.2f}'

    except Exception as e:
        cprint(f'Ошибка задачи {task_code} {e}', 'red')
    finally:
        result_dict.update(prepare_data) #Результирующий словарь с оценками заполняемости дополняется начальными данными {'Код задачи','Псевдоним','Дата начала','Дата окончания','Всего предложений','Статус выполнения')

        await cursor.close()    #закрытие соединения с базой данных
        await conn.close()      #закрытие соединения с базой данных
        cprint((f'задача {task_code} выполнена за  = {(time.time() - task_beg):.1f}' + ' сек'), 'cyan')
        return result_dict

async def gobaby(task_list_prepare_data,for_connect):
    '''
    Основная асинхронная  функиця создания и запуска других  асинхронных функций получения данных по задачам (в формате async).
    Пришлось создать эту задачу, а не осуществлять данный код в функции main потому, что в main присутствует
    вызов обычной (синхронной) функции task_list_prepare для предварительного получения списка задач. В таком случае main
    не может быть асинхронной а как следствие не может создать и запустить другие асинхронные функции.
    Либо необходимо единственныю синхронную функцию тоже делать условно асинхронной (что, как бы не нужно по сути,
    она должна выполнится первой и выполняется достаточно быстро)

    :param task_list_prepare_data:  Список всех задач парсинга (Иначе они называются задачи ИМПОРТА)
    :param for_connect: Настройки для подключения к базе данных
    :return:   Ничего не возвращает
    '''
    begin = time.time() #условное время начала запуска скрипта

    futures = [task_full_info(args, for_connect) for args in task_list_prepare_data] #созадниае списка фьючерсов (функций, которые будут выполнены в асинхронном режиме)
    #Это и есть те финкии, которые будут получать данные о  качестве заполняемости по каждой задаче
    #В качестве аргументов передается параметры одной задачи из списка задач и свойства подключения в базе данных

    done, _ = await asyncio.wait(futures)   #Запуск всех ФУТУР !!!
    for_output=[]                           #Список в котором будут собираться результаты, формат - СПИСОК СЛОВАРЕЙ
    for future in done:                     #Если футура выполнена
        try:
            for_output.append(future.result())  #В результирующий список добавляются результат выполнения футуры
        except Exception as e:
            cprint(('Ошибка вот такая', e),'red')
    cprint((f'Общее время выполнения всех задач  = {(time.time() - begin):.1f}' + ' сек'), 'green')

    #Запись всех результатов полученных выше в файл json  и указанием даты и времени в имени файла
    time_now = (datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S"))
    with open(r'C:\Users\user1\AppData\Roaming\JetBrains\PyCharmCE2020.1\scratches\parse_resilt_items\async_items_({}).json'.format(time_now), 'w', encoding='utf-8') as fp:
        json.dump(for_output, fp, ensure_ascii=False, indent=4)
        print(r'Файл записан в C:\Users\user1\.PyCharmCE2019.3\config\scratches\async_items_({}).json'.format(time_now))

    # печать в терминал функцией print_char_dict в нужном формате  всех полученых результатов  по задачам из результирующего сортированого списка
    for i in sorted(for_output, key=lambda x: x['Псевдоним']):
        print_char_dict(i)  #В функцию передается словарь, так как результирующий список for_output содержит список словарей

def main(task_num):
    """
    Скрипт предназначен для  получения оценочных данных результатов парсинга. конкретного ЗАДАНИЯ из базы данных RWAY.
    Задание включает в себя определенное количество задач парсинга. (задачи не парсинга = обрадотки фильтруются на стадии sql запросов)
    Каждая задача парсит конкретный источник данных, например сайт АВИТО и содержит определенное количество предложений.
    Каждое предложение содержит ряд ключевых характеристик, без которых предложение не имеет ценности.
    Результаты парсинга представлены как процентное соотношение заполнения ключевых  характеристик по отношению к общему количеству предложений по этой задаче(источнику).


    :param task_num: Номер задачи для проверки в формате 0001-0606
    :return:  Скрипт который ничего не возвращает в виде результата, но запускает функции выводящие результат в терминал.
     """

    for_connect = 'DRIVER={SQL Server};SERVER=10.199.13.60;DATABASE=rway;UID=vdorofeev;PWD=Sql01' #Настройки подключения к базе RWAY, будет использовано для подключения в синхронной и асинхронной функциях
    """ Нельзя передавать одно подключение в потоки (aioodbc.connect(dsn=for_connect) или pyodbc.connect(for_connect)), первый поток выполнится а другие скажут что соединение занято
        Нельзя передавать в потоки один курсор (conn.cursor()), первый поток выполнится а другие сойдут с ума т.к. курсор переместится в первом потоке
        Но можно передать простую строку настроек которая будет использована  при каждом подключении к серверу SQL
    """
    prepare_data = task_list_prepare(task_num,for_connect)      # полчение перечня всех задач входящих в задание (в синхронном формате, потому что быстро)
    loop = asyncio.get_event_loop()                             # запуск планировщика  асинхронных задач
    loop.run_until_complete(gobaby(prepare_data,for_connect))   #помещение в планировщик основной задачи
    loop.close()                                                #закрытия планировщика задач после выполнения
    print()

if __name__ == "__main__":
    prime_taks_name = '0001-0610'
    main(prime_taks_name)
