from collections import namedtuple
import pandas as pd
import numpy as np
import os
import os.path
from docx import Document
from docx.shared import Pt
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
import progressbar
import re
import config as CONFIG


def create_files_list():
    """
    Функция создает 2 пути к ворд и эксель файлам с учетом расположения скрипта, перибирает ворд и эксель файлы создавая
    список. Каждый элемент списка  это тоже список, содержащий
    1)путь к Эксель файлу
    2)соответствующий данному эксель файлу Ворд Шаблон
    3)Название файла (для сохранения результата)

    Для сопоставления ворд и эксель файлов и них должны быть одинаковые названия
    :return: Возвращает Список списков
    """
    file_list_result = []       #Результирующий список
    current_folder = os.getcwd()
    exec_folder = os.path.join(current_folder, CONFIG.EXCEL_FOLDER) #Пути содаржания Эксель файлов
    if CONFIG.USE_A_FOOTER:
        docx_folder = os.path.join(current_folder, CONFIG.DOCX_FOLDER_4_COL) if CONFIG.FOUR_COLUMNS  else os.path.join(current_folder, CONFIG.DOCX_FOLDER_NORMAL)  # Пути содаржания ворд файлов c колонтитулом и без
    else:
        docx_folder = os.path.join(current_folder, CONFIG.DOCX_FOLDER_NO_FOOTER_4_COL) if CONFIG.FOUR_COLUMNS else os.path.join(current_folder, CONFIG.DOCX_FOLDER_NO_FOOTER)  # Пути содаржания ворд файлов c колонтитулом и без

    xlsx_files = sorted(os.listdir(exec_folder))  # Список файлов в Эксель папке
    doc_files = sorted(os.listdir(docx_folder))  # Списов файлов в ворд папке

    for x_file in xlsx_files:  # Перибираем Эксель файлы
        x_file_name, x_extension = os.path.splitext(x_file)
        for d_file in doc_files:  # Перибираем Ворд файлы
            d_file_name, d_extension = os.path.splitext(d_file)
            if x_file_name == d_file_name:  # Если названия файлов совпадат (только названия без расширений), то полные пути к этим файлам и общее название добавляютися в итоговый список
                file_list_result.append(
                    [os.path.join(exec_folder, x_file), os.path.join(docx_folder, d_file), x_file_name])
    return file_list_result


def give_table_num_list(data):
    start_num = False
    if 'city' in data.filename_for_save and 'source' not in data.filename_for_save and CONFIG.CITY_TABLE_NUM_EDIT:
        start_num = CONFIG.CITY_TABLE_START_NUM
    elif 'city' in data.filename_for_save and 'source' in data.filename_for_save and CONFIG.CITYSOURCE_TABLE_NUM_EDIT:
        start_num = CONFIG.CITYSOURCE_TABLE_START_NUM
    elif 'sklad' in data.filename_for_save and 'source' not in data.filename_for_save and CONFIG.SKLAD_TABLE_NUM_EDIT:
        start_num = CONFIG.SKLAD_TABLE_START_NUM
    elif 'sklad' in data.filename_for_save and 'source' in data.filename_for_save and CONFIG.SKLADSOURDE_TABLE_NUM_EDIT:
        start_num = CONFIG.SKLADSOURDE_TABLE_START_NUM

    if start_num:
        num_paragraphs = [num for num, i in enumerate(data.word_data.paragraphs) if i.text.find('Таблица ') != -1]  # Номера строк с номером таблицы
        paragraphs_values = ['{:.1f}'.format(start_num + (i / 10)) for i in range(len(data.word_all_tables))]  # Номера строк с номером таблицы
    else:
        num_paragraphs = paragraphs_values =  [None for i in range(len(data.word_all_tables))]
    return num_paragraphs, paragraphs_values


def dividing_data_into_parts(excel_file_path, word_file_path, file_name_for_save):
    data = namedtuple('data',
                      ['excel_full_file',
                       'excel_sheet'
                       'word_data',
                       'word_all_tables'
                       'select_doc_table'
                       'doc_table_num',
                       'doc_table_value',
                       'num_paragraphs_for_edit_date',
                       'num_paragraphs_for_edit_item',
                       'filename_for_save',
                       'source'])

    data.word_data = Document(word_file_path)  # Загрузка текущего Ворд файла
    data.excel_full_file = pd.ExcelFile(excel_file_path)  # Загрузка текущего Эксель файла
    excel_sheets_list = data.excel_full_file.sheet_names
    data.word_all_tables = data.word_data.tables
    data.filename_for_save = file_name_for_save
    data.source = True if 'source' in excel_file_path else False
    doc_table_num, doc_table_value = give_table_num_list(data)
    data.row_not_for_replace = 3 if data.source else 2  # Количество строк в таблице не подлежащих редактированию (Заголовки)
    if data.source:
        num_paragraphs_for_edit_date = [num for num, i in enumerate(data.word_data.paragraphs) if i.text.find('предоставленных в дату предоставления') != -1]  # Номера строк с датами
        num_paragraphs_for_edit_item = [num for num, i in enumerate(data.word_data.paragraphs) if i.text.find('4.3.') != -1]  # Номера строк с пунктами
    else:
        num_paragraphs_for_edit_date = num_paragraphs_for_edit_item = [None]

    # #Этот блок нужен для отладки, по хорошему нужно удалить, но код дорадатывается, постоянно это писать не хочу.
    # print('данные ворд ', data.word_data)
    # print('данные эксель ', data.excel_full_file)
    # print('Вкладки эсель ',excel_sheets_list)
    # print('номер таблицы для замены ',doc_table_num)
    # print('номер таблицы на который заменить ',doc_table_value)
    # print('номер параграфа для замены даты ',num_paragraphs_for_edit_date)
    # print('номер параграфа для замены пункта ',num_paragraphs_for_edit_item)
    # print('имя файла ля соханения',data.filename_for_save)

    assert len(data.word_all_tables) == len(doc_table_num) == len(doc_table_value) == len(num_paragraphs_for_edit_date) == \
           len(num_paragraphs_for_edit_item) == len(excel_sheets_list), 'Количество таблиц в ворде, вкладок (таблиц)' \
            ' в экселе,  параграфов замены даты, парафов замены пунктов, номеров таблиц, и номеров для замены не совпадает'

    for num in range(len(data.word_all_tables)):
        data.select_doc_table = data.word_all_tables[num]
        data.excel_sheet = excel_sheets_list[num]
        data.doc_table_num = doc_table_num[num]
        data.doc_table_value = doc_table_value[num]
        data.num_paragraphs_for_edit_date = num_paragraphs_for_edit_date[num]
        data.num_paragraphs_for_edit_item = num_paragraphs_for_edit_item[num]
        yield data


def load_and_prepare_excel(excel_info, sheet, source):
    """
    Функция загружает таблицу из эксель файла excel_info расположенной на вкладке sheet
    Удаляет из эксель таблицы не нужные столбцы, форматирует эксель таблицу так, чтобы  ее можно было вставить в Ворд таблицу, это
    Для файлов sourc:
        Последняя строка состоит из 3-х ячеек, первая из которых это 9-ть объединенных ячеек ,
        чтобы последняя строка выглядела как |Итого| 0000| 0000| в датафрейме нужно сделать в последней строке
        вот так |не важно|не важно|не важно|не важно|не важно|не важно|не важно|не важно|Итого:|0000|0000|
    Для других файлов добавляет строку с суммой по каждому столбцу (в эксель таблице этих данных нет а в Ворд таблице они нужны)

    :param excel_info: Общий файл эксель
    :param sheet:   Название вкладки в эксель файле
    :param source: True/False Если мы работаем с файлом/данными "источников", например файл city_source.docx
    :return:  Возвращает эксалевскую таблицу в виде  pandas DateFrame и параметр equal_result (True/False),
     который показывает, нужно ли менять пункт указанные в тексте перед таблицей.
    """
    data = pd.read_excel(excel_info, sheet_name=sheet)  # Загрузка таблицы
    col_name_for_del = ['subject', 'city', 'Unnamed: 0']  # Столбцы, которые требуется удалить если они есть в таблице
    equal_result = False
    for col_name in col_name_for_del:  # Удаление ненужных столбцов
        if col_name in data.columns:
            data = data.drop([col_name], axis=1)
    if source:  # в описании функции указано
        data.iloc[-1, 9] = 'Итого:'
        equal_result = True if data.iloc[-1][-1] == data.iloc[-1][-2] else False  # если данные в двух последних ячейках последней строки совпадают то True иначе False
    else:  # в описании функции указано
        data.number = data.number.astype('int')  # Изменить тип столбца на целочисленный, потому что в эксель он указан строкой,  и поэтому неверно выравнивается в ячейке Ворд таблицы.
        total_row = data.sum()
        total_row['number'] = 'Итого:'
        data = data.append(total_row, ignore_index=True)

    return data, equal_result


def get_cell_format(cell):
    """
    Функция сохранение параметров форматированя текста и шрифта текущей ЯЧЕЙКИ
    :param cell:  Текущая ячейка
    :return:    Возвращает параметры выравнивания, размера шрифта, названия шрифта, жирный текст или нет
    """
    run = cell.paragraphs[0].runs
    alignment = 0 if cell.paragraphs[0].alignment == None else int(
        re.findall('\d', str(cell.paragraphs[0].alignment))[0])
    font = run[0].font
    font_size = font.size
    font_name = font.name
    font_bold = font.bold
    return alignment, font_size, font_name, font_bold


def set_cell_format(cell, font_ali, font_size, font_bold, font_name):
    """
    Фунция применяет к текущей ячекки параметры форматирования выравнивания, размера и тд.
    :param cell:    текущая ячейка
    :param font_ali:    выравнивание    по гооризонтали
    :param font_size:  размер текста
    :param font_bold:   Жирный/нежирный
    :param font_name:   Название шрифта
    :return:    ничего не возвращает так как изменяет параметры ячейки по ссылке cell
    """
    run = cell.paragraphs[0].runs
    cell.paragraphs[0].alignment = font_ali  # 0 for left, 1 for center, 2 right, 3 justify ....
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER  # Выравнивание текста по центру по ВЕРТИКАЛИ
    font = run[0].font
    font.size = font_size
    font.bold = font_bold
    font.name = font_name


def get_parag_format(parag):
    run = parag.runs
    font = run[0].font
    font_size = font.size
    font_name = font.name
    font_bold = font.bold
    return font_size, font_name, font_bold


def set_parag_format(parag, font_size_was, font_name_was, font_bold_was):
    run = parag.runs
    font = run[0].font
    font.size = font_size_was
    font.bold = font_bold_was
    font.name = font_name_was


def edit_date_and_item_in_parag(current_table,target=''):
    """
    Функция редактирует дату и пункт (если требуется) на который ссылается текст в тексте перед текущей таблицей
    :param parag: Строка, которую необходимо отредактировать
    :param new_data:    Новая дата для замены, берется из названия закладки с текущей таблицей в экселе
    :param edit_paragraph_item_bool:    True/False  требуется ли менять пункт
    :return:    ничего не возвращает так как редактирует  параграф (текст перед таблицей), по ссылке parag
    """
    if target == 'date':
        parag = current_table.word_data.paragraphs[current_table.num_paragraphs_for_edit_date]
        format_rule = get_parag_format(parag)   #Сохранение параметров форматированя текста и шрифта
        date_in_parag = re.findall(r'\d{2}\.\d{2}\.\d{4}', parag.text)[0]                   #Находим даты в тексте
        parag.text = parag.text.replace(date_in_parag, current_table.excel_sheet)           #Заменяем дату на новую
        set_parag_format(parag,*format_rule)    #Применить параметры форматированя текста и шрифта
    elif target == 'item':
        parag = current_table.word_data.paragraphs[current_table.num_paragraphs_for_edit_item]
        format_rule = get_parag_format(parag)
        parag.text = parag.text.replace('4.3.2', '4.3.3')                  # Заменяем дату на новую
        set_parag_format(parag, *format_rule)
    elif target == 'table_num':
        parag = current_table.word_data.paragraphs[current_table.doc_table_num]
        format_rule = get_parag_format(parag)
        replace_it = re.findall(r'[\d.]{3,5}', parag.text)[0]
        parag.text = parag.text.replace(replace_it, current_table.doc_table_value)  # Заменяем номер на новый
        set_parag_format(parag, *format_rule)


def edit_head_in_doc_table(doc_table, excel_table):
    """
    Функция редактирует даты в заголовке Ворд таблицы на те, что присутствуют в заголовках Экселесвкой таблицы.
    Функция редактирует пункты ('4.3.2', '4.3.3') в заголовке. Условие, в одну дату, если совпадают суммы по столбцам
    "Фактическое количество" и Статус актуальной ссылки", тогда пункт '4.3.3', иначе '4.3.2'
    В данной функции используется list.pop() для перемещения по спискам, чтобы постепенно выбирать элемент за элементом  а не делать  какие-то указатели на соответствующий индекс элемента.
    :param doc_table:       Воровская таблица
    :param excel_table:     Экселевская таблица
    :return:  Ничего не возвращает т.к. редактирует таблицу  по ссылке doc_table
    """
    pat_date = r'[\d.]{10}'  # патерн поиска даты для регулрки
    # из списка колонок выбирает колонки где есть даты (регуляркой), сортирует по убыванию, чтобы позже через list.pop() выбирать)
    date_from_excel_columns = [re.findall(pat_date, i)[0] for i in excel_table.columns if re.findall(pat_date, i)][::-1] #развернуть для pop()
    condition_list_edit_par =  []           #Список False/True для замены пункта в заголовке ил инет.
    condition_data = excel_table.iloc[-1][1:]       #берем последниюю строку таблицы с суммами по столбцам, исключая первую колонку, там написано "Итого"
    for item in np.array_split(condition_data, len(condition_data) / 2):    #Получившийся массив делим по два столбца (Фактические и актуальные) и перечисляем его
        for _ in range(2): condition_list_edit_par.append(True if np.unique(item).size == 1 else False) #Если значения двух элементов списка совпадают (т.е. совпадают фактические и актуальные)
        # то дважды добавляем True (Совпадают) или False (Несовпадают). Дважды, потому что в заголовке 10 столбцов с датами и пунктами, содениненых по два, но фактически их 10
    condition_list_edit_par = list(reversed(condition_list_edit_par)) # Разварачиваю список чтобы выбирать значения спомощью list.pop()

    doc_cols = len(doc_table.columns)  # Количество колонок в Вордовской таблице
    num_row_for_edit = 0  # первая строка, которую необходимо править, там где даты
    for col in range(doc_cols):  # перебираю все колонки
        cell = doc_table.rows[num_row_for_edit].cells[col]  # выбор конкретной ячейки
        date_in_cell = re.findall(pat_date, cell.text)  # Поиск в тексте ячейки даты
        if date_in_cell:  # Если дата есть
            font_ali_was, font_size_was, font_name_was, font_bold_was = get_cell_format(cell)  # Получить характеристики текста
            try:
                cell.text = cell.text.replace(date_in_cell[0],date_from_excel_columns.pop())  # Заменить дату ячейки на ту, что была в списке заголовков экселевской таблицы
            except IndexError:
                print('Число столбцов в эксель файлах, которые нужно вставить ворд не соответсвует числу столбков в шаблоне ворд,\n'
                      'или число таблиц в экселе не соответствует числу таблиц в шаблоне ворд. За эти настройки отвечают \n'
                      'параметры USE_A_FOOTER и FOUR_COLUMNS в файле конфига config.py, и необходимо проверить эксель файл с данными')
            if condition_list_edit_par.pop():   #нужно ли менять пункт в данном ячейке
                cell.text = cell.text.replace('4.3.2', '4.3.3')

            set_cell_format(cell, font_ali_was, font_size_was, font_bold_was,
                            font_name_was)  # Применить характеристики текста

    doc_table.rows[num_row_for_edit].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY  # Разрешить редактирование высоты ячейки
    doc_table.rows[num_row_for_edit].height = Pt(CONFIG.HEIGHT_OF_CELL_IN_HEADER)  # Установить высоту ячейки


def add_row_in_doc_table(current_table,excel_row):
    """
    Изначально в шаблоне таблицы Ворд 3 или 4 строки, это 2 строки заголовка (в том числе объединенные ячейки),
    1 строка номера колонки (в файлах источников), 1 строка с данными для образца форматирования ячеек и текста.
    Функция добавляет в таблицу ворд столько строк (без строк заголовков), чтобы количество соответствовало количеству строк из эксель файла
    из которого берутся  данные для таблицы.
    :param excel_row: Количество строк в эксель файле
    :param select_tables:  Текущая таблица ворд с которой мы работаем
    :param row_not_for_replace: Количество строк  уже присутствующих в таблице (строки заголовков и шаблона форматрования)
    :param source: True/False Если мы работаем с файлом/данными "источников", например файл city_source.docx, в таком случаи
    последнуу строку итогов нужно объекдинить, чтобы получилось 3 ячейки.
    :return: Ничего не возвращает, т.к. функция вносит изменения в таблицу по ссылке select_tables
    """
    doc_rows = len(current_table.select_doc_table.rows)                      #текущее кол-во строк в Вод таблице
    doc_cols = len(current_table.select_doc_table.columns)                   #текущее количество столбцов в таблице
    delta = excel_row - doc_rows + current_table.row_not_for_replace         #Колчество строк которое необходимо добавить, чтобы в итоговой таблице было столькоже строк с данными сколько в экселе
    for _ in range(delta):
        current_table.select_doc_table.add_row()                             #Добавляем строки
    if current_table.source:                                                 #Если тип файла source, в последней строке  объединяем ячейки, чтобы получилось 3 колонки
        for col in range(doc_cols):
            if col < doc_cols - 2:
                current_table.select_doc_table.rows[-1].cells[0].merge(current_table.select_doc_table.rows[-1].cells[col])


def save_result_docx(doc_data, filename):
    """
    Функция Сохраняет итоги соединения эксель файла и вордовского шаблона в папку RESULT_DOCX.
    Если такая папка не существовала рядом с нашим скриптом, то создаем ее.
    :param doc_data: Итоговые данные для сохранения
    :param filename: имя файла для сохранения
    :return: Нечего возвращать
    """
    current_folder = os.getcwd()        #Текущая папка
    result_folder = os.path.join(current_folder, 'RESULT_DOCX') #Куда хотим сохранить
    if not os.path.exists('RESULT_DOCX'):   #Если такой папки нет, создаем
        os.mkdir(result_folder)
    doc_data.save(os.path.join(result_folder, f'{filename}.docx'))  #Сохраняем
    print(f'* файл {filename} сохранен')


def join_exceldata_and_docxtable(current_tab):

    excel_data, edit_paragraph_item_bool = load_and_prepare_excel(current_tab.excel_full_file, current_tab.excel_sheet, current_tab.source) #Загруженная и подготовленая конкретная эксель таблица, и значение необходимости менять пункт параграфа
    excel_row, excel_col = excel_data.shape #количество строк и столбцов в эксель, чтобы перебирать, как всегда.
    items_row = current_tab.select_doc_table.rows[-1]  # Строка шаблона ВОРД  как образец параметров выравнивания, имени шрифта , размера шрифта при формировании таблицы
    if current_tab.num_paragraphs_for_edit_item and edit_paragraph_item_bool:
         edit_date_and_item_in_parag(current_tab,target ='item')
    if current_tab.num_paragraphs_for_edit_date:
        edit_date_and_item_in_parag(current_tab,target ='date') #Редактирование даты в строке перед таблицей
    if current_tab.doc_table_num and current_tab.doc_table_value:
        edit_date_and_item_in_parag(current_tab, target='table_num')  # Редактирование даты в строке перед таблицей
    if not current_tab.source:
        edit_head_in_doc_table(current_tab.select_doc_table,excel_data)

    add_row_in_doc_table(current_tab,excel_row) #добавление в шаблон Ворда нужное количество строк (подробнее в функции add_row_in_doc_table)
    rows = len(current_tab.select_doc_table.rows)  #Количество строк в ворд таблице после  добавления строк
    cols = len(current_tab.select_doc_table.columns)   #Количество колонок в  ворд таблице
    bar = progressbar.ProgressBar(max_value=rows)   #Для отображения динамики происходящего

    for row in range(rows): #Перебираем строки ворд таблицы
        bar.update(row)     #Для отображения динамики происходящего
        if row - current_tab.row_not_for_replace < excel_row:   #В ворд таблице строк больше чем в Эксель, так как в ворде заголовки считаются строками (0 в итоге в ворде на 2-3 строки больше в зависимости от файла
            for col in range(cols):                 #Перебираем колонки
                cell = current_tab.select_doc_table.rows[row].cells[col]   #Текущая ячейка
                items_cell = items_row.cells[col]               #Соответствующая ячейка из строки шаблона, специально оставленная для  определения необходимого форматирования всех новых строк в шаблоне
                if row > current_tab.row_not_for_replace - 1:  # Если строки не заголовок (ниже чем row_not_for_replace)
                    font_ali_was, font_size_was, font_name_was, font_bold_was = get_cell_format(items_cell)  # Получить характеристики текста
                    cell.text = str(excel_data.iloc[row - current_tab.row_not_for_replace][col]) #Заменить текст в ячеки соотвествующим текстом Эксель ячейки
                    if row == rows - 1: #Если строка последняя и файл типа source, то у него Должен быть жирный шриф
                        font_bold_was = True if current_tab.source else font_bold_was
                        font_ali_was = 0 if col < cols - 2 and current_tab.source else font_ali_was #А если .то первая ячейка последней строки, то у нее выравнивание по левому краю
                    set_cell_format(cell, font_ali_was, font_size_was, font_bold_was,font_name_was)  # Применить характеристики текста

                    current_tab.select_doc_table.rows[row].height_rule = None if current_tab.source else WD_ROW_HEIGHT_RULE.EXACTLY    # Разрешить редактирование высоты ячейки Если файл source
                    current_tab.select_doc_table.rows[row].height = Pt(CONFIG.HEIGHT_OF_CELL_IN_SOURCE) if current_tab.source else Pt(CONFIG.HEIGHT_OF_CELL_IN_NOT_SOURCE) # Установить высоту ячейки, в зависимости от параметра source
    bar.finish()    #Для отображения динамики происходящего


def manage():
    files_path_list = create_files_list() #Создать список путей к файлам шаблонов и данных
    for excel_file_path, word_file_path, file_name_for_save in files_path_list: #Перебираем этот список постепенно обрабатывая все файлы
        print('* {} + {} ---> {}.docx'.format(os.path.split(excel_file_path)[1], os.path.split(word_file_path)[1], file_name_for_save))
        table_list = dividing_data_into_parts(excel_file_path,word_file_path,file_name_for_save)
        for table_data in table_list:
            join_exceldata_and_docxtable(table_data)
            word_data = table_data.word_data
        save_result_docx(word_data, file_name_for_save)  # Сохраняем результаты редактиирования Ворд фала



if __name__ == '__main__':
    manage()