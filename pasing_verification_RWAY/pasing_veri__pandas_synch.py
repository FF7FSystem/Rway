import pandas as pd
from sqlalchemy import create_engine
from termcolor import cprint
import time
import json
import concurrent.futures

"""
Подробное описание есть в основнам рабочем файле pasing_verification_char, а здесь используется
пандас для подключение к базе данных и обработки даных, но использовать подключение пандас в асинхронном
режиме не удалось  (использовалась асинхронная библиотека asinc io pandas, не  модуль откузывается подключаться к  базе
подобным  коннектором ('mssql+pymssql://{username}:{password}@{server}/{database}'.format(**for_connect)))

Так что это рабочей код но работает в синхронном режиме и в среднм для обрадотки задачи тратится 40-50 минут 
"""



for_connect={'server' : '10.199.13.60', 'database' : 'rway', 'username' :'vdorofeev', 'password' :'Sql01'}
engine = create_engine('mssql+pymssql://{username}:{password}@{server}/{database}'.format(**for_connect))
conn = engine.connect()

def task(task_num):
    all_task = pd.read_sql(f'''
        SELECT _IDRRef,ДатаНачала,ДатаОкончания,_Fld198 as 'Код задачи' ,Псевдоним,jour.Статус FROM [rway].[dbo].[_Task62] task 
        left join [rway].[Документ].[Задание] doc on task._Fld192RRef = doc.Ссылка and task._Fld231RRef=0xAB47107B44B1669011E990CEFFC204B6
        left join [rway].[Справочник].[Источники] sourc on task._Fld197RRef = sourc.Ссылка 
        left join [rway].[РегистрСведений].[ЖурналЗадач] jour on task._IDRRef = jour.Задача  
        where doc.КодЗадания = '{task_num}'  ORDER BY Псевдоним 
         ''', conn) #and jour.Статус = 2
    cprint('Таблица  Задачи загружена', 'yellow')
    # cprint(all_task,'green')
    return (all_task)

def char_of_task(val_char):
    task_time_begin = time.time()
    link = val_char[0]
    nickname = val_char[1]
    tsk_num = val_char[2]
    task_status =val_char[3]

    # print(nickname,val_char)

    link='0x' + link.hex().upper() # Преобразование из байтового представления в целочисленное для подстановки в селект
    cnt_predl = pd.read_sql(f'''
        Select COUNT(Предложение) from [rway].[РегистрСведений].[ПредложенияЗадач] where Задача= {link}''', conn)
    cnt_predl=int(cnt_predl.loc[0][0])
    cprint(nickname,'red',end=' ')
    cprint(('Текущий статус = '+str(int(float(task_status)))), 'green')

    predl_char_dict = {'!Псевдоним':nickname,'!Всего предложений':cnt_predl,'!Код Задачи':tsk_num,'!Статус':float(task_status)}

    Char_1 = pd.read_sql(f'''
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
            count(case when Наименование = '7.Улица' and Значение <> 0 then  1 end) 'Улица',
            count(case when Наименование = '8.Дом' and Значение <> 0 then  1 end) 'Дом',
            count(case when Наименование = 'Назначение Объекта Предложения Список' and Значение <> 0 then  1 end) 'Назначение Объекта Предложения Список',
            count(case when Наименование = 'Дата Сбора Информации' and Значение <> 0 then  1 end) 'Дата Сбора Информации'
    From
    (SELECT Предложение FROM [rway].[dbo].[_Task62] task left join [rway].[РегистрСведений].[ПредложенияЗадач] predl on _IDRRef = predl.Задача where _Fld198 = '{tsk_num}') num_pred
    join 
    (select Объект, Наименование, 
        case
            when Значение_Число > 0 then  1
            when Значение_Строка <> '' then  1
            when Значение_Дата is not null then  1
        end Значение
    
         from [rway].[РегистрСведений].[ЗначенияХарактеристик] Xap left join [rway].[ПВХ].[Характеристики] Xap_val on Xap.Характеристика = Xap_val.Ссылка) xap_tab_pred
    on num_pred.Предложение = xap_tab_pred.Объект
                ''', conn)

    Char_2 = pd.read_sql(f'''
            SELECT	count(case when [АдресПредставление] <> '' then  1 end) 'Адрес',
                    count(case when [ДатаРазмещения] is not null then  1 end) 'ДатаРазмещения',
                    count(case when [НаселенныйПункт] <> '' then  1 end) 'НаселенныйПункт',
                    count(case when [Город] <> '' then  1 end) 'Город',
                    count(case when [Описание] <> '' then  1 end) 'Описание',
                    count(case when [СамаяРанняяДатаРазмещения] is not null then  1 end) 'СамаяРанняяДатаРазмещения',
                    count(case when [Сегмент] = 0x00000000000000000000000000000000 then null else  1 end) 'Сегмент',
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
            where _Fld198 = '{tsk_num}' ''', conn)

    Char_all=pd.concat([Char_1,Char_2],axis=1) #Сложение 2-х таблиц форматом(Строка+Строка)
    for i in (Char_all.columns):               # Извлечения данных из таблицы pandas в словарь где ключ=Хар-ка, Значение = Процент количества заполнености хар-ки к общему числу предложений.
        if int(Char_all[i][0]) !=0:            #Дабы небыло деления на 0 в случае когда всего предложений  = 0
            val = int(Char_all[i][0]) / cnt_predl * 100
        else:
            val=0
        predl_char_dict[i] = f'{val:.2f}'

    cprint((f'{(time.time() - task_time_begin):.1f}' + ' сек'), 'cyan')
    return(predl_char_dict)

def print_char_dict(char_dict):
    for key in sorted(char_dict):
        print('{:_<40}'.format(key), end=' ')
        if key in ('!Всего предложений','!Код Задачи','!Псевдоним','!Статус'):
            term='white'
        else:
            term = 'green' if float(char_dict[key]) >= 100 else 'red'
        cprint(char_dict[key],term)

    #cprint(("{0:.2f}".format(time.time() - task_time_begin)),'cyan')

def main(mission):
    cprint(('Время старта = '+(time.strftime("%H:%M:%S", time.localtime()))),'red')
    task_info=task(mission)
    #list_of_task_num=list(task_num)
    #rint(list_of_task_num)
    # all_info=char_of_task(task_num)
    result=[]
    list_of_task_num = [[i[1]['_IDRRef'],i[1]['Псевдоним'],i[1]['Код задачи'],i[1]['Статус']] for i in task_info.iterrows()]
    # for i in list_of_task_num:
    # print(list_of_task_num)

    for i in list_of_task_num:
        dict_avl=char_of_task(i)
        print_char_dict(dict_avl)
        result.append(dict_avl)
    print(result)

    with open(r'C:\Users\user1\Documents\items_cont\items.json', 'w', encoding='utf-8') as fp:
        json.dump(result, fp, ensure_ascii=False, indent=4)

    """
    with concurrent.futures.ThreadPoolExecutor() as tpe:
        list_of_value=tpe.map(char_of_task,list_of_task_num[:2])
    print(list(list_of_value))
    """

    # print(list(list_of_value))

    # print_char_dict(list_of_value)
if __name__ == "__main__":
    #t_x = time.time()

    main('0001-0608')  #передать номер в формате строки"
    #cprint(('Время Выполения функции МАИН', time.time() - t_x), 'blue')