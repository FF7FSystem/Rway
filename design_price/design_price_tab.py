import pandas as pd
import numpy as np
# from config_design_price_tab import * #импорт tab_name_dict,target_sheet
import json

save_in_file,tab_name_dict,target_sheets = None,None,None

def tab_slices(tab_data):
    """
    В эксель файле ищет строки содержащие названия таблиц (содержат Средневзвешенные цены), записывает как номер
    строки начала таблицы. От начальной ячейки ищет первую полностью пустую строку. Найдя первую пустую строку,
    записывает ее номер как номер строки окончания таблицы.
    :param tab_data: контент эксель файла (Датафрейм)
    :return: список, содержащий список начала и конца таблицы, например  [[6, 22], [26, 42], [75, 90], [96, 113]]
    """
    result = []
    temp = []   #Временный список где первое число, это номер строки начала таблицы, второе число, номер строки последней в таблице
    tab_fl = False
    for num  in range(len(tab_data)):
        str_row = ''.join([str(i) for i in (tab_data.iloc[num])])
        for tab_name in tab_name_dict:
            if tab_name in  str_row and not temp: #Если название таблицы есть в строке и временный список пуст (не содержит номера первой строки таблицы)
                temp.append(num)
                tab_fl = True
                break
            elif str_row =='' and tab_fl and temp: #Если таблица окончилась т.е. посшли пустые строки и временный список не пуст (не содержит номер первой строки таблицы)
                temp.append(num)
                tab_fl = False
                result.append(temp)
                temp = []
                break
    return result

def drop_empty_col(tab):
    """
    Удаляет пустые колонки
    :param tab: таблица
    :return:  отредактированная таблица
    """
    drop_cols = tab.columns[tab.replace('', np.nan).isnull().all()]
    return tab.drop(drop_cols, axis=1)

def select_required_col(tab,tab_name):
    """
    выбирает "нужные" колонки из входящей таблицы:
     а именно  = нулевой столбец   + столбцы с ценой (Средневзвешенные цены, руб./сот.)
    первую строку удаляем (не берем), т.к. в ней название таблицы
    :param tab: таблица (Датафрейм)
    :param tab_name:    Название таблицы
    :return:    отредактированная таблица
    """
    try:
        lead_col=tab_name_dict[tab_name]['lead_col']
    except KeyError:
        print(f"Не найдено имя таблицы  (указано ниже) в словаре конфига (название  текущей таблицы отличается \
        от словарного на пробел в конце строки, точку  и т.д.... (удачи) \n '{tab_name}'" )

    col_num = [i for i in range(len(tab.columns)) if lead_col in list(tab.iloc[:, i])]
    col_num.insert(0, 0)
    return tab.iloc[1:, col_num]

def load_config():
    """
    Загрузить файл конфига. Данные из конфига записать в глобальниые переменнные, чтобы не протаскивать в функции
    :return: Ничего
    """
    try:
        with open('config_design_price_tab.json', encoding='utf8') as inp:  #Загрузка конфига
            conf = json.load(inp)
    except Exception as e:
        print('Файл конфига отсутствует/не открывается!')
        raise e
    #Переназначение глобальных переменных
    global save_in_file
    global target_sheets
    global tab_name_dict
    save_in_file = conf['save_in_file']
    target_sheets = conf['target_sheets']
    tab_name_dict = conf['tab_name_dict']

def rename_head(tab,tab_name):
    """
    Переименование заголовков  таблицы.
    Названия заголовков берется из нулевой строки, потом эта строка  и еще несолько удаляются
    согласно настройки таблицы в словаре tab_name_dict
    :param tab: таблица (Датафрейм)
    :param tab_name:    Название таблицы
    :return:    отредактированная таблица
    """
    head = [i for i in tab.iloc[0].to_list() if i]
    dict_head = dict(zip(tab.columns, head))
    return tab.rename(columns=dict_head)[tab_name_dict[tab_name]['row_for_del']:]

def drop_empty_row(tab):
    """
    Удаление пустых строк.
    В первом столбце указаны названия строк. Строки считаются пустыми если в столбцах начиная со второго нет данных.
    :param tab:  таблица (Датафрейм)
    :return: отредактированная таблица
    """
    not_del_rows = [i for i in range(len(tab)) if not tab.replace('', np.nan).iloc[i, 1:].isnull().all()]
    return tab.iloc[not_del_rows, :]

def prepare_tab(current_tab):
    """
    получает на вход срез из Эксель файла, содержащий таблицу, удаляет лишее
    :param current_tab: срез из Эксель файла
    :return: список: имя таблицы, таблица (ДатаФрейм)
    """
    tab_name = ''.join([str(i)for i in current_tab.iloc[0]]).strip()  # Имя таблицы полностью
    tab_name_no_date = [key for key in tab_name_dict if key in tab_name][0] #имя таблицы без даты (берется первое частичное совпадение названия с ключем словаря)
    new_tab = drop_empty_col(current_tab) # Удаление пустых столбцов
    new_tab = new_tab.replace(' – ', 0) #Заменить прочерки
    new_tab = new_tab.replace(' - ', 0) #Заменить прочерки
    new_tab = select_required_col(new_tab,tab_name_no_date)  #выбираются нужные колонки
    new_tab = rename_head(new_tab,tab_name_no_date)          #переименовываются столбцы
    new_tab = drop_empty_row(new_tab)       #удаляются строки не сожержыщие  значения в колонках от второй и т.д.
    return [tab_name,new_tab]

def bond_same_tab(tab_list):
    """
    Соединение таблиц с одинаковым названием
    :param tab_list: Список сожержащий название таблицы и Датафрейм
    :return:  Словарь, ключ, название таблицы, значение Датафрейм
    """
    result_tab_dict={}
    for name,current_tab in tab_list:
        if name not in result_tab_dict:
            result_tab_dict[name]=current_tab.reset_index(drop=True)
        else:
            try:
                result_tab_dict[name] = result_tab_dict[name].merge(current_tab,how='outer')#.fillna(0)
            except Exception as e:
                print(f'''При попытке "склеивания" таблиц (с одинаковым названием), не найдена колонка (название колонки) присутствующая в обеих таблицах,
ориентировочно вот этот столбец "{current_tab.columns[0]}" не равен "{result_tab_dict[name].columns[0]}"''')
                raise e
    return result_tab_dict

def load_excel_content(excel_file_path):
    """
    Загружает вкладки (указаны в конфиге target_sheets) эксель файла.
    В каждом датафрейме, полученном загрузкой вкладки, необходимо переименовать колонки на 0,1,2,3 и т.д.
    чтобы при формировании единого файла (методом append) один датафрейм был под другим без смещения (если в разных
    таблицах одинакове столцы, данные из них помещаются одни под другие, пополоняя столбце, если столбцы назваются
    поразному, то новый столбез добавляется в конец, расширяя таблицу)
    :param excel_file_path: путь до файла
    :return: возвращает контент вкладки эксель Датафрейм
    """
    result_table=pd.DataFrame()
    excel_data = pd.ExcelFile(excel_file_path)
    sheets = excel_data.sheet_names
    index_sheets =[sheets.index(sheet) for sheet in sheets if sheet in target_sheets] #список индексов закладок в эксель, соответствующие закладкам указанным в конфиге.
    for ind_sheet in index_sheets: #перебор закладок
        this_sheet_data = pd.read_excel(excel_data, ind_sheet, na_filter=False) #Загрузка вкладки в Datafreame
        this_sheet_data =  this_sheet_data.rename(columns={col:num for num,col in enumerate(this_sheet_data.columns)}) #Переименование колонок
        result_table = result_table.append(this_sheet_data,ignore_index=True).fillna(0) #добавление в результирующий DataFrame
    return result_table

def save_result_in_file(result_tab_dict):
    sheets_num=1
    writer = pd.ExcelWriter('OUTPUT_DESIGN_PRICE.xlsx', engine='xlsxwriter')
    for name,tab in result_tab_dict.items():
        tab.to_excel(writer, str(sheets_num))
        sheets_num+=1
    writer.save()

def main(excel_file_path):
    """
    Основная функция, которая загружает указанный эксель файл, ищет в ней отдельные таблицы. В каждой таблице обрезает
    лишние, форматирует и возвращает в виде словаря где ключ: название таблицы, значение таблица (Датафрейм)
    :param excel_file_path: путь к фалу Эксель, который необходимо  обрадотать
    :return: Словарь
    """
    load_config() #Загрузка конфига в глобальные переменные
    all_tab=load_excel_content(excel_file_path) #Загрузка контента эксель файла
    slices = tab_slices(all_tab)    #Поиск отдельных таблиц по ключевому слову. Возвращает список "срезов" для каждой таблицы в виде списка.
    print(slices)
    tab_list = [prepare_tab(all_tab.iloc[begin:end]) for begin,end in slices]   #Каждую таблицу форматирует и берет только данные о ценах
    result_tab_dict = bond_same_tab(tab_list)   #Таблицы с одинаковыми названиями склеиваются
    if save_in_file:                            #Для отладки (Запись результирующих таблиц в эксель файл)
        save_result_in_file(result_tab_dict)
    return result_tab_dict


if __name__ == '__main__':
    # main('Шаблон_ЗУ_Мск.xlsx')
    # main('Шаблон_ЗУ_МО.xlsx')
    # main('Шаблон_ЗУ_регионы.xlsx')
    # main('Шаблон_ЗУ_МО.xlsx')
    main(r'C:\Users\user1\Desktop\Шаблон_КН_МСК_типо_олд.xlsx')

'''
1  - не соединяются таблицы, причина, разные первые столюбцы, по которым соединяется
2 - нет названия таблицы в словаре конфигурации, причина, в названии таблицы пропущены пробелы, точки и т.д.  
'''