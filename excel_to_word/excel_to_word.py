import pandas as pd
import os
import os.path
from docx import Document
from docx.shared import Pt
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
import progressbar
import re


def add_row_in_doc_table(excel_row, select_tables, row_not_for_replace, source):
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
    doc_rows = len(select_tables.rows)                      #текущее кол-во строк в Вод таблице
    doc_cols = len(select_tables.columns)                   #текущее количество столбцов в таблице
    delta = excel_row - doc_rows + row_not_for_replace      #Колчество строк которое необходимо добавить, чтобы в итоговой таблице было столькоже строк с данными сколько в экселе
    for _ in range(delta):
        select_tables.add_row()                             #Добавляем строки
    if source:                                              #Если тип файла source, в последней строке  объединяем ячейки, чтобы получилось 3 колонки
        for col in range(doc_cols):
            if col < doc_cols - 2:
                select_tables.rows[-1].cells[0].merge(select_tables.rows[-1].cells[col])


def edit_date_and_item_in_parag(parag, new_data, edit_paragraph_item_bool):
    """
    Функция редактирует дату и пункт (если требуется) на который ссылается текст в тексте перед текущей таблицей
    :param parag: Строка, которую необходимо отредактировать
    :param new_data:    Новая дата для замены, берется из названия закладки с текущей таблицей в экселе
    :param edit_paragraph_item_bool:    True/False  требуется ли менять пункт
    :return:    ничего не возвращает так как редактирует  параграф (текст перед таблицей), по ссылке parag
    """
    #Сохранение параметров форматированя текста и шрифта
    run = parag.runs
    font = run[0].font
    font_size_was = font.size
    font_name_was = font.name
    font_bold_was = font.bold

    date_in_parag = re.findall(r'\d{2}\.\d{2}\.\d{4}', parag.text)[0]   #Находим даты в тексте
    parag.text = parag.text.replace(date_in_parag, new_data)            #Заменяем дату на новую
    if '4.3.2' in parag.text and edit_paragraph_item_bool:              #Если требуется то редактируем пункт на который ссылается текст
        parag.text = parag.text.replace('4.3.2', '4.3.3')
    elif '4.3.3' in parag.text and not edit_paragraph_item_bool:        #Если требуется то редактируем пункт на который ссылается текст
        parag.text = parag.text.replace('4.3.3', '4.3.2')

    # Применить к измененному тексту параметра  текста из шаблона
    run = parag.runs
    font = run[0].font
    font.size = font_size_was
    font.bold = font_bold_was
    font.name = font_name_was


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
    data = pd.read_excel(excel_info, sheet_name=sheet)  #Загрузка таблицы
    col_name_for_del = ['subject', 'city', 'Unnamed: 0'] #Столбцы, которые требуется удалить если они есть в таблице
    equal_result = False
    for col_name in col_name_for_del:                   #Удаление ненужных столбцов
        if col_name in data.columns:
            data = data.drop([col_name], axis=1)
    if source:  #в описании функции указано
        data.iloc[-1, 9] = 'Итого:'
        equal_result = True if data.iloc[-1][-1] == data.iloc[-1][-2] else False #если данные в двух последних ячейках последней строки совпадают то True иначе False
    else: #в описании функции указано
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
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER  #Выравнивание текста по центру по ВЕРТИКАЛИ
    font = run[0].font
    font.size = font_size
    font.bold = font_bold
    font.name = font_name


def edit_head_in_doc_table(doc_table, excel_table):
    """
    Функция редактирует даты в заголовке Ворд таблицы на те, что присутствуют в заголовках Экселесвкой таблицы.
    :param doc_table:       Воровская таблица
    :param excel_table:     Экселевская таблица
    :return:  Ничего не возвращает т.к. редактирует таблицу  по ссылке doc_table
    """
    pat_date = r'[\d.]{10}'  # патерн поиска даты для регулрки
    # из списка колонок выбирает колонки где есть даты (регуляркой), сортирует по убыванию, чтобы позже через list.pop() выбирать)
    date_from_excel_columns = sorted(
        [re.findall(pat_date, i)[0] for i in excel_table.columns if re.findall(pat_date, i)], reverse=True)
    doc_cols = len(doc_table.columns)  # Количество колонок в Вордовской таблице
    num_row_for_edit = 0  # первая строка, которую необходимо править
    for col in range(doc_cols):  # перебираю все колонки
        cell = doc_table.rows[num_row_for_edit].cells[col]  # выбор конкретной ячейки
        date_in_cell = re.findall(pat_date, cell.text)  # Поиск в тексте ячейки даты
        if date_in_cell:  # Если дата есть
            font_ali_was, font_size_was, font_name_was, font_bold_was = get_cell_format(
                cell)  # Получить характеристики текста
            cell.text = cell.text.replace(date_in_cell[0],
                                          date_from_excel_columns.pop())  # Заменить дату ячейки на ту, что была в списке заголовков экселевской таблицы
            set_cell_format(cell, font_ali_was, font_size_was, font_bold_was,
                            font_name_was)  # Применить характеристики текста

    doc_table.rows[num_row_for_edit].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY  # Разрешить редактирование высоты ячейки
    doc_table.rows[num_row_for_edit].height = Pt(26)  # Установить высоту ячейки


def join_exceldata_and_docxtable(excel_sheet, select_doc_tables, doc_parag, excel_full_info, word_data, source):
    """
    1) Загружает и подготавливает  текущие ворд и эксель таблицы для соединения.
    Эксель данные  подготавливаются в load_and_prepare_excel
    Водр данные подготавливаются с помощью добавления строк в add_row_in_doc_table, Так же если требуется, то
    редактируются строки перед таблицей в edit_date_and_item_in_parag и заголовки в ворд таблице с помощью edit_head_in_doc_table
    2)Соединяет таблицы с учетом форматирования текста и строк
    3) отображает прогресс бар в терминале по ходу процесса заполнения шаблонов таблиц Ворд данными из экселя

    :param excel_sheet: Имя вкладки эксель таблицы откуда брать данные
    :param select_doc_tables:  текущая ворд таблица для редактирования
    :param doc_parag: номер строки перед таблицей для редактирования
    :param excel_full_info: Весь эксель файл, чтобы по номеру вкладки получить таблицу
    :param word_data: Весь ворд файл, чтобы выбирать и редактировать строки
    :param source: True/False Если мы работаем с файлом/данными "источников", например файл city_source.docx
    :return: ничего не возвращаем т.к. работаем с сылкой word_data, и все меняется там без необходимсоти возврата (как редактировать список из функции)
    """
    row_not_for_replace = 3 if source else 2  # Количество строк в таблице не подлежащих редактированию (Заголовки)
    excel_data, edit_paragraph_item_bool = load_and_prepare_excel(excel_full_info, excel_sheet, source) #Загруженная и подготовленая конкретная эксель таблица, и значение необходимости менять пункт параграфа
    excel_row, excel_col = excel_data.shape #количество строк и столбцов в эксель, чтобы перебирать, как всегда.
    items_row = select_doc_tables.rows[-1]  # Строка шаблона ВОРД  как образец параметров выравнивания, имени шрифта , размера шрифта при формировании таблицы

    if source: #Если файл типа source, то выполянются соответствующие радактирования текста перед таблицами или заголовков таблицы
        edit_date_and_item_in_parag(word_data.paragraphs[doc_parag], excel_sheet, edit_paragraph_item_bool)
    else:
        edit_head_in_doc_table(select_doc_tables, excel_data)

    add_row_in_doc_table(excel_row, select_doc_tables, row_not_for_replace, source) #добавление в шаблон Ворда нужное количество строк (подробнее в функции add_row_in_doc_table)
    rows = len(select_doc_tables.rows)  #Количество строк в ворд таблице после  добавления строк
    cols = len(select_doc_tables.columns)   #Количество колонок в  ворд таблице
    bar = progressbar.ProgressBar(max_value=rows)   #Для отображения динамики происходящего

    for row in range(rows): #Перебираем строки ворд таблицы
        bar.update(row)     #Для отображения динамики происходящего
        if row - row_not_for_replace < excel_row:   #В ворд таблице строк больше чем в Эксель, так как в ворде заголовки считаются строками (0 в итоге в ворде на 2-3 строки больше в зависимости от файла
            for col in range(cols):                 #Перебираем колонки
                cell = select_doc_tables.rows[row].cells[col]   #Текущая ячейка
                items_cell = items_row.cells[col]               #Соответствующая ячейка из строки шаблона, специально оставленная для  определения необходимого форматирования всех новых строк в шаблоне
                if row > row_not_for_replace - 1:  # Если строки не заголовок (ниже чем row_not_for_replace)
                    font_ali_was, font_size_was, font_name_was, font_bold_was = get_cell_format(
                        items_cell)  # Получить характеристики текста
                    cell.text = str(excel_data.iloc[row - row_not_for_replace][col]) #Заменить текст в ячеки соотвествующим текстом Эксель ячейки
                    if row == rows - 1: #Если строка последняя и файл типа source, то у него Должен быть жирный шриф
                        font_bold_was = True if source else font_bold_was
                        font_ali_was = 0 if col < cols - 2 and source else font_ali_was #А если .то первая ячейка последней строки, то у нее выравнивание по левому краю
                    set_cell_format(cell, font_ali_was, font_size_was, font_bold_was,
                                    font_name_was)  # Применить характеристики текста

                    select_doc_tables.rows[row].height_rule = None if source else WD_ROW_HEIGHT_RULE.EXACTLY    # Разрешить редактирование высоты ячейки Если файл source
                    select_doc_tables.rows[row].height = Pt(20) if source else Pt(12)                           # Установить высоту ячейки, в зависимости от параметра source
    bar.finish()    #Для отображения динамики происходящего


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
    exec_folder, docx_folder = os.path.join(current_folder, 'XLSX'), os.path.join(current_folder, 'DOCX')   #Пути содаржания Эксель и ворд файлов
    xlsx_files = sorted(os.listdir(exec_folder))    #Список файлов в Эксель папке
    doc_files = sorted(os.listdir(docx_folder))     #Списов файлов в ворд папке

    for x_file in xlsx_files:           #Перибираем Эксель файлы
        x_file_name, x_extension = os.path.splitext(x_file)
        for d_file in doc_files:        #Перибираем Ворд файлы
            d_file_name, d_extension = os.path.splitext(d_file)
            if x_file_name == d_file_name:  #Если названия файлов совпадат (только названия без расширений), то полные пути к этим файлам и общее название добавляютися в итоговый список
                file_list_result.append(
                    [os.path.join(exec_folder, x_file), os.path.join(docx_folder, d_file), x_file_name])
    return file_list_result


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


def manage():
    """
    Функция управления.
    1) Создает список  содержащий пути к Эксель файлам, соответствующие им  Ворд файлы и названяи итоговых файлов для сохранения (подробнее в create_files_list)
    2) перебирает данный список загружая содержимое Эксель файла и соответствующего ему Ворд шаблона
    3) определяет в каких строках (не таблицах) Ворд шаблона необходимо внести корректировки
    4)Поскольк бывает ,ворд файл содержит 5 таблиц (шаблонов) и их нужно заполнить данными из 5 экселевских вкладок формируется список
    каждый элемент которого -  список из
         а)Название вкладки экселя из которой брать данные
         б)Таблица Ворд, в вордовском файле , в которую будут добавляется данные из Эксель
         в)Текст перед таблицей, который необходимо отредактировать
    Далее этот список перебирается постепенно наполняя ворд файл данными.
    содержимое итогового ворд фала сохраняется через функцию save_result_docx
    :return: Ничего не возвращает
    """
    files_path_list = create_files_list() #Создать список путей к файлам шаблонов и данных
    for excel_file_path, word_file_path, file_name_for_save in files_path_list: #Перебираем этот список постепенно обрабатывая все файлы
        print('* {} + {} ---> {}.docx'.format(os.path.split(excel_file_path)[1], os.path.split(word_file_path)[1],
                                              file_name_for_save))
        excel_all_info = pd.ExcelFile(excel_file_path) #Закгрузка текущего Эксель файла
        word_data = Document(word_file_path)           #Закгрузка текущего Ворд файла

        if 'source' in excel_file_path:                #Для текущего ворд файла находим строки перед таблицами в которые будут внесены изменения
            num_paragraphs_for_edit = [num for num, i in enumerate(word_data.paragraphs) if
                                       i.text.find('предоставленных в дату предоставления') != -1]
            source = True
        else:
            num_paragraphs_for_edit = [False, ]         #если строк для редатирования вне таблиц нет (и чтобы зип ниже работал)
            source = False
        list_of_sheets_tab_parag = list(zip(excel_all_info.sheet_names, word_data.tables, num_paragraphs_for_edit))
        for item in list_of_sheets_tab_parag: #Перебираем таблицы в Ворд файле и соответствующие им вклдаки в эксель и строки для замены
            join_exceldata_and_docxtable(*item, excel_all_info, word_data, source)
        save_result_docx(word_data, file_name_for_save) #Сохраняем результаты редактиирования Ворд фала


if __name__ == '__main__':
    manage()


'''
#Удаление строки
result_row = select_tables.rows[-1]
tbl = select_tables._tbl
tbl.remove(result_row._tr)
'''

#вариант копирования эксель таблиц в ворд
# def copy_win32():
#     from win32com import client
#     excel = client.Dispatch("Excel.Application")
#     word = client.Dispatch("Word.Application")
#     doc = word.Documents.Open(r"C:\Users\user1\Desktop\excel_word\demo.docx")
#     book = excel.Workbooks.Open(r"C:\Users\user1\Desktop\excel_word\city_source_2.xlsx")
#     sheet = book.Worksheets(1)
#     sheet.Range("B1:M69").Copy()  # Selected the table I need to copy
#     doc.Content.PasteExcelTable(False, False, True)
#     # sheet.Close()
#     # doc.Close()
