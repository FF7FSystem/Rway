import pandas as pd
from sqlalchemy import create_engine
from termcolor import cprint
'''   Словарь соответствия Названий Представлений  и имен таблиц
tab = pd.read_excel('base_struct.xlsx', sheet_name='Лист1')
tab_col=[ i for i in tab[tab.columns[2]]]
tab_row=[ '_'+i for i in tab[tab.columns[1]]]
tab_dict= {tab_col[i]: tab_row[i] for i in range(len(tab_col))}
'''

for_connect={'server' : '10.199.13.60', 'database' : 'rway', 'username' :'vdorofeev', 'password' :'Sql01'}
engine = create_engine('mssql+pymssql://{username}:{password}@{server}/{database}'.format(**for_connect))
conn = engine.connect()

def name_table(tab_index_code):     # функция возвращает имя таблицы и имя поля содержащего строковое значение данных
    if isinstance( tab_index_code, int): #на случай ручного добавления  таблиц приходящие данные в целочисленном значении
        tab_index= tab_index_code
    else:
        tab_index = int.from_bytes((tab_index_code), byteorder='big')  # Перевод из байтов в инт значения
    if tab_index < 132:                             # Таблицы Reference имеют счет до 131 (_Reference131 - последняя)
        return '_Reference'+str(tab_index),'_Description',tab_index
    elif tab_index > 131:                            # Таблицы _Enum имеют счет c 133 (_Enum133 - первая)
        return '_Enum' + str(tab_index), '_EnumOrder',tab_index

def create_globdict(*manual):
    cprint('Start update globdict', 'green')
    list_add=[]
    def addval_globdict(table_name, col_val, tab_num ):
        global globdict
        if tab_num not in list_add:
            list_add.append(tab_num)
            if col_val == '_EnumOrder':
                temp = pd.read_sql_table(table_name, conn)[['_IDRRef', col_val]]
                temp = temp.rename(columns={'_EnumOrder': '_Description'})
            else:
                temp = pd.read_sql_table(table_name, conn)[['_IDRRef', col_val]]
            globdict = pd.concat([globdict, temp])
            cprint(('В словарь добавлены  поля из :', table_name), 'red')

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

def search_value(**dict):           #возвращает  строкое значение ячейки ,которая ищется в table_name по первичному ключу соответствующему search_str, в колонке searc_col
    table_name=dict['table_name']
    search_str = dict['search_str']
    searc_col = dict['searc_col']

    temp_zn = pd.read_sql_table(table_name, conn)                  # Считывание с сервера таблицы по имени
    temp_zn = temp_zn[temp_zn._IDRRef == search_str][[searc_col]]  # Находит в скаченой по имени таблице строку search_str и выбирает из этой таблице только 1 колонку searc_col ,
    return list(temp_zn[searc_col])[0]                              # из полученой таблици с 1 строкой и 1 колонкой выбирает значение (кривовато, но обратится по индексу не получается т.к. он имеет неизвестное значение из  исходной таблицы)

def znach_xap(obj):
    obj_str = '0x' + obj.hex().upper()
    rszh=pd.read_sql('SELECT Объект,Характеристика, Значение_Дата,Значение_Строка,Значение_Число,Значение_RTRef,Значение,Значение_Тип FROM [rway].[РегистрСведений].[ЗначенияХарактеристик] as tb WHERE tb.Объект  = {}'.format(obj_str),conn)
    result_df = pd.DataFrame(columns = ['Объект','Характеристика', 'Значение']) #Создание пустой результирующей таблицы
    for i in range(len(rszh)):
        znach='Поле не определено' #На случай попытки определения значения и характеристик Тип значения которых не равен 1,3,4,5,8
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
            znach = globdict[globdict._IDRRef == rszh.iloc[i]['Значение']]['_Description']
            try:
                znach = list(znach)[0]
            except:
                znach = 'Данные не найдены 8-(' #если вдруг конкретные таблицы не были добавлены в глобальный словарь

        if znach=='Поле не определено':
            cprint(('Поле не определено',cod),'red')
        #Чтобы добавить Характеристику в глобальный словарь необходимо доработать добавление таблиц типа _Chrc174
        xap = search_value(table_name = '_Chrc174', search_str = rszh.iloc[i]['Характеристика'],searc_col= '_Description') #дл каждой характеристики выполняется функция поиска значений
        result_df.loc[i]={'Объект' : rszh.iloc[i]['Объект'], 'Характеристика': xap, 'Значение' : znach}

    print(result_df)
    #print(result_df.loc[9])

def pred_obe_nedv(lnk):
    result_df2 = pd.DataFrame(columns = ['Ссылка','Характеристика','Значение']) #Создание пустой результирующей таблицы

    lnk_str = '0x' + lnk.hex().upper()
    spon=pd.read_sql('SELECT * FROM [rway].[Справочник].[ПредложенияОбъектовНедвижимости] as tb WHERE tb.Ссылка = {}'.format(lnk_str),conn) [['Ссылка','Код','Адрес','АдресAhunter','АдресЗначенияПолей',
        'АдресПредставление','АдресУровеньПривязки','АктуальнаяСсылкаИсточника','Город','ДатаПересмотраЭкспозиции', 'ДатаПроверкиАктуальности','ДатаРазмещения',
        'ИсточникИнформации','МастерПредложение','НаселенныйПункт','ОбъектНедвижимости' ,'Описание', 'Подсегмент','СамаяПоздняяДатаПроверкиАктуальности',
        'СамаяРанняяДатаРазмещения','Сегмент','СсылкаИсточника','Субъект','ТипРынка','ТипСделки']] #Можно так можно в селект забросить

    link_temp=list(spon['Ссылка'])[0]
    col=spon.columns
    for id, item in enumerate(col):
        if item != 'Ссылка':
            Xap = item
            znach = globdict[globdict._IDRRef == list(spon[item])[0]]['_Description']
            try:
                znach = list(znach)[0]
            except:
                znach = 0
            znach = znach if znach != 0 else  list(spon[item])[0]
            result_df2.loc[id]={'Ссылка': link_temp , 'Характеристика': Xap, 'Значение' : znach}

    print(result_df2)
    #print(result_df2.loc[9])


#znach_xap(b'\x80\x00\x06f\xf3AC@E\x04c\xb2)\xe9J\xdb')
#pred_obe_nedv(b'\x80\x01\x17u\x15|E\x15@3\xb4\xe0/z\xb9\xc8')

globdict = pd.DataFrame(columns=['_IDRRef', '_Description'])  # Создание глобального славаря
def main(task_num):
    create_globdict()   # Добавление в глобальный словарь данных из таблицы ПВХ
    create_globdict(99)  # Добавление в глобальный словарь таблиц по числовому индексу (словарь обрабатывает входные данных как список)
    cprint(globdict, 'blue')
    task_num = '\'' + task_num + '\''
    task_data = pd.read_sql('SELECT _IDRRef,_Date_Time,_Name, _Fld198 FROM [rway].[dbo].[_Task62] tb WHERE tb._Fld198 = {}'.format(task_num), conn)
    task_cod = '0x' + task_data.loc[0]['_IDRRef'].hex().upper()
    pred_task=pd.read_sql('SELECT Задача,Предложение FROM [rway].[РегистрСведений].[ПредложенияЗадач] tb WHERE tb.Задача = {}'.format(task_cod),conn)
    cprint('Количество задач в последней строке', 'green')
    cprint(pred_task,'green')

    for i in range(len(pred_task)):
        task_link=pred_task.loc[0]['Предложение']
        pred_obe_nedv(task_link)
        znach_xap(task_link)
        if i ==3:
            break

if __name__ == "__main__":
    main('0001-0405-0001')  #передать номер в формате строки"

