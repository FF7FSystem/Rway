import pandas as pd
from pandas import ExcelWriter
from sqlalchemy import create_engine
from termcolor import cprint
import time
import re
import json

def benchmark(func):
    def wrapper(*args, **kwargs):
        t = time.time()
        res = func(*args, **kwargs)
        print(f'Функция {func.__name__} выполнилась за время {time.time() - t:.2f}s')
        return res
    return wrapper


''' Создание словаря соответствия Названий Представлений  и имен таблиц
    Сохранение данных в файл result.json
tab = pd.read_excel('base_struct.xlsx', sheet_name='Лист1')
tab_row = ['_' + i for i in tab[tab.columns[1]]]
tab_col_temp = [str(i) for i in tab[tab.columns[2]]]
tab_col = []
for i in tab_col_temp:
    pat = r'([\w]+)[.]?'
    result = re.findall(pat, str(i))
    result = ".".join(['[' + i + ']' for i in result])
    tab_col.append('[rway].' + result)
tab_dict = {tab_row[i]:tab_col[i] for i in range(len(tab_col))}

with open(r"result.json", 'w') as write_file: 
    json.dump(tab_dict, write_file, sort_keys=True, indent=4)
'''
with open('result.json') as json_data: # Загрузка словаря соотвествия названия таблиц и Пердставлений на сервере SQL
    tab_dict = json.load(json_data)

for_connect={'server' : '10.199.13.60', 'database' : 'rway', 'username' :'vdorofeev', 'password' :'Sql01'}
engine = create_engine('mssql+pymssql://{username}:{password}@{server}/{database}'.format(**for_connect))
conn = engine.connect()

#@benchmark
def name_table(tab_index_code):     # функция возвращает имя таблицы и имя поля содержащего строковое значение данных
    if isinstance( tab_index_code, int): #на случай ручного добавления  таблиц приходящие данные в целочисленном значении
        tab_index= tab_index_code
    else:
        tab_index = int.from_bytes((tab_index_code), byteorder='big')  # Перевод из байтов в инт значения
    if tab_index < 132:                             # Таблицы Reference имеют счет до 131 (_Reference131 - последняя)
        return '_Reference'+str(tab_index),'Наименование',tab_index
        #return '_Reference' + str(tab_index), '_Description', tab_index
    elif tab_index > 131 and tab_index < 174:       # Таблицы _Enum имеют счет c 133 173 (_Enum133 - первая)
        return '_Enum' + str(tab_index), 'Значение',tab_index
        #return '_Enum' + str(tab_index), '_EnumOrder', tab_index
    elif tab_index > 173 and tab_index < 175:       # Таблицы _Enum имеют счет c 133 173 (_Enum133 - первая)
        return '_Chrc' + str(tab_index), 'Наименование', tab_index
        #return '_Chrc' + str(tab_index), '_Description',tab_index   # Для добавления таблицы _Chrc174 (имена характеристик)

@benchmark
def create_globdict(*manual):
    cprint('Start update globdict', 'green')
    list_add=[]

    #@benchmark
    def addval_globdict(table_name, col_val, tab_num ):
        global globdict
        if tab_num not in list_add:
            list_add.append(tab_num)
            if col_val == 'Значение':
                t_x = time.time()
                #temp = pd.read_sql_table(table_name, conn)[['_IDRRef', col_val]]
                temp = pd.read_sql('SELECT Ссылка,{} FROM {}'.format(col_val,tab_dict[table_name]), conn)
                cprint(('Время выполнения строки 70', time.time() - t_x),'blue')
                #temp = temp.rename(columns={'_EnumOrder': '_Description'})
                temp = temp.rename(columns={'Значение': 'Наименование'})
            else:
                t_x = time.time()
                # temp = pd.read_sql_table(table_name, conn)[['_IDRRef', col_val]]
                temp = pd.read_sql('SELECT Ссылка,{} FROM {}'.format(col_val, tab_dict[table_name]), conn)
                cprint(('Время выполнения строки 77', time.time() - t_x), 'blue')
            globdict = pd.concat([globdict, temp])
            cprint(('В словарь добавлены  поля из :', tab_dict[table_name]), 'red')

    # Поиск числовой приставки к таблице ('_Reference###,_Enum###'), в которых берутся значения хар-к объектов
    if not manual:
        PVH=pd.read_sql('SELECT * FROM [rway].[ПВХ].[Характеристики] as tab WHERE tab.ЗначениеПоУмолчанию_RTRef <> 0',conn)[['ЗначениеПоУмолчанию_RTRef']]
        for i in range(len(PVH)):
            table_name,col_val,tab_num=name_table(PVH.iloc[i]['ЗначениеПоУмолчанию_RTRef'])
            addval_globdict(table_name,col_val,tab_num)
    #Ручное добавление таблиц (из списка по индексу приставки )в случаи их отсутствия в таблице ПВХ. Хараактеристики
    if manual:
        for i in manual:
            table_name, col_val, tab_num = name_table(i)
            addval_globdict(table_name, col_val, tab_num)


@benchmark
def znach_xap(obj):
    obj_str = '0x' + obj.hex().upper()
    t_x = time.time()
    rszh=pd.read_sql('SELECT Характеристика, Значение_Дата,Значение_Строка,Значение_Число,Значение,Значение_Тип FROM [rway].[РегистрСведений].[ЗначенияХарактеристик] as tb WHERE tb.Объект  = {}'.format(obj_str),conn)
    cprint(('Время выполнения строки 110', time.time() - t_x), 'blue')
    print(rszh)
    #result_df = pd.DataFrame(columns = ['Объект','Характеристика', 'Значение']) #Создание пустой результирующей таблицы
    #df.loc[df['shield'] > 6]
    #cprint(rszh[rszh['Значение_Тип'] == b'\x08'],'red')
    #cprint(rszh[rszh['Значение_Тип'] == b'\x01'], 'red')
    #cprint(rszh[rszh['Значение_Тип'] == b'\x03'], 'red')
    #cprint(rszh[rszh['Значение_Тип'] == b'\x04'], 'red')
    df_temp = pd.merge(rszh, globdict, how='inner', left_on='Характеристика', right_on='Ссылка')[['Значение_Дата',
    'Значение_Строка','Значение_Число','Значение','Значение_Тип','Наименование']].rename(columns={'Наименование': 'Характеристика'})
    df5=(df_temp[df_temp['Значение_Тип'] == b'\x05'][['Характеристика','Значение_Строка']].rename(columns={'Значение_Строка': 'Значение'}))
    df4=(df_temp[df_temp['Значение_Тип'] == b'\x04'][['Характеристика','Значение_Дата']].rename(columns={'Значение_Дата': 'Значение'}))
    df3 = (df_temp[df_temp['Значение_Тип'] == b'\x03'][['Характеристика', 'Значение_Число']].rename(columns={'Значение_Число': 'Значение'}))
    df1 = (df_temp[df_temp['Значение_Тип'] == b'\x01'][['Характеристика', 'Значение']])
    df8 = (df_temp[df_temp['Значение_Тип'] == b'\x08'][['Характеристика', 'Значение']])

    df8=pd.merge(df8, globdict, how='inner', left_on='Значение', right_on='Ссылка')[['Характеристика','Наименование']]
    #В случаи если в ячейке со значением 0 после  соедениния по этому полю данная строка удаляется (видимо отсутствует в словаре)
    result_df = pd.concat([df1,df3,df4,df5,df8])

    '''
    for i in range(len(rszh)):
        if rszh.iloc[i]['Значение'] != b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00':
            cod = (int.from_bytes((rszh.iloc[i]['Значение_Тип']), byteorder='little'))
            if  cod == 5:
                znach = rszh.iloc[i]['Значение_Строка']
            elif  cod == 4:
                znach = rszh.iloc[i]['Значение_Дата']
            elif  cod == 3:
                znach = rszh.iloc[i]['Значение_Число']
            elif  cod == 1:
                znach = 'Должно быть значение ДА/НЕТ'
            elif cod == 8:
                znach = globdict[globdict.Ссылка == rszh.iloc[i]['Значение']]['Наименование']
                try:
                    znach = list(znach)[0]
                except:
                    znach = 'Данные не найдены в словаре' #если вдруг конкретные таблицы не были добавлены в глобальный словарь
        else:
            znach=''

        #xap = search_value(table_name = '_Chrc174', search_str = rszh.iloc[i]['Характеристика'],searc_col= '_Description') #дл каждой характеристики выполняется функция поиска значений
        xap = globdict[globdict.Ссылка == rszh.iloc[i]['Характеристика']]['Наименование']
        try:
            xap = list(xap)[0]
        except:
            xap = 'Данные не найдены в словаре'  # если вдруг конкретные таблицы не были добавлены в глобальный словарь
        result_df.loc[i]={'Объект' : rszh.iloc[i]['Объект'], 'Характеристика': xap, 'Значение' : znach}
        '''
    cprint('Данные из таблицы Значения характеристик выгружены', 'magenta')
    cprint(result_df[['Характеристика','Значение']],'magenta')
    print(result_df.loc[9])
@benchmark
def pred_obe_nedv(lnk):
    result_df2 = pd.DataFrame(columns = ['Ссылка','Характеристика','Значение']) #Создание пустой результирующей таблицы

    lnk_str = '0x' + lnk.hex().upper()
    t_x = time.time()
    #spon=pd.read_sql('SELECT * FROM [rway].[Справочник].[ПредложенияОбъектовНедвижимости] as tb WHERE tb.Ссылка = {}'.format(lnk_str),conn) [['Ссылка','Код','Адрес','АдресAhunter','АдресЗначенияПолей',
     #   'АдресПредставление','АдресУровеньПривязки','АктуальнаяСсылкаИсточника','Город','ДатаПересмотраЭкспозиции', 'ДатаПроверкиАктуальности','ДатаРазмещения',
     #   'ИсточникИнформации','МастерПредложение','НаселенныйПункт','ОбъектНедвижимости' ,'Описание', 'Подсегмент','СамаяПоздняяДатаПроверкиАктуальности',
     #   'СамаяРанняяДатаРазмещения','Сегмент','СсылкаИсточника','Субъект','ТипРынка','ТипСделки']] #Можно так можно в селект забросить
#
    spon = pd.read_sql('''SELECT Ссылка, Код, Адрес, АдресAhunter,АдресЗначенияПолей,АдресПредставление, АдресУровеньПривязки,
                          АктуальнаяСсылкаИсточника,Город,ДатаПересмотраЭкспозиции,ДатаПроверкиАктуальности,ДатаРазмещения,
                          ИсточникИнформации,МастерПредложение,НаселенныйПункт,ОбъектНедвижимости,Описание,Подсегмент,
                          СамаяПоздняяДатаПроверкиАктуальности,СамаяРанняяДатаРазмещения,Сегмент,СсылкаИсточника,
                          Субъект,ТипРынка,ТипСделки 
                          FROM [rway].[Справочник].[ПредложенияОбъектовНедвижимости] as tb WHERE tb.Ссылка = {}'''.format(lnk_str), conn)  # Можно так можно в селект забросить

    cprint(('Время выполнения строки 149', time.time() - t_x), 'blue')
    link_temp=list(spon['Ссылка'])[0]
    col=spon.columns
    for id, item in enumerate(col):
        if item != 'Ссылка':
            Xap = item
            znach = globdict[globdict.Ссылка == list(spon[item])[0]]['Наименование']
            try:
                znach = list(znach)[0]
            except:
                znach = 'None'

            if znach == 'None':
                znach = list(spon[item])[0]
            elif znach == b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00':
            #elif znach == '0x00000000000000000000000000000000':
                znach=''

            result_df2.loc[id]={'Ссылка': link_temp , 'Характеристика': Xap, 'Значение' : znach}

    cprint('Данные из таблицы Преждложения Объектов Недвижимости выгружены', 'cyan')
    cprint(result_df2[['Характеристика','Значение']],'cyan')
    print(result_df2.loc[9])


#znach_xap(b'\x80\x00\x06f\xf3AC@E\x04c\xb2)\xe9J\xdb')
#pred_obe_nedv(b'\x80\x01\x17u\x15|E\x15@3\xb4\xe0/z\xb9\xc8')

globdict = pd.DataFrame()  # Создание глобального славаря
@benchmark
def main(task_num):
    # Для использования как модуль создание словаря вынести в функцияю, чтобы наполнить его 1 раз для всех задач.
    #create_globdict()   # Добавление в глобальный словарь данных из таблицы ПВХ
    #create_globdict(99,174)  # Добавление в глобальный словарь таблиц по числовому индексу (словарь обрабатывает входные данных как список)
    #cprint(globdict, 'blue')
    task_num = '\'' + task_num + '\''
    task_data = pd.read_sql('SELECT _IDRRef,_Date_Time,_Name, _Fld198 FROM [rway].[dbo].[_Task62] tb WHERE tb._Fld198 = {}'.format(task_num), conn)
    cprint('Таблица  Задачи загружена', 'green')

    task_cod = '0x' + task_data.loc[0]['_IDRRef'].hex().upper()
    pred_task=pd.read_sql('SELECT Задача,Предложение FROM [rway].[РегистрСведений].[ПредложенияЗадач] tb WHERE tb.Задача = {}'.format(task_cod),conn)
    cprint('Таблица Предложения задач загружена', 'green')

    for i in range(len(pred_task)):
        #if i ==1:
        #    break
        task_link=pred_task.loc[0]['Предложение']
        pred_obe_nedv(task_link)
        znach_xap(task_link)


if __name__ == "__main__":
    main('0001-0405-0001')  #передать номер в формате строки"

