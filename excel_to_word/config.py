# Использовать ли сноски в файле шаблона ворд
USE_A_FOOTER = False
# Файл шаблона ворд должен сожержать 4 колонки
FOUR_COLUMNS = False

# Расположение файлов шаблонов ворд с разными хар-ками
DOCX_FOLDER_NORMAL = 'DOCX\DOCX_NORMAL'
DOCX_FOLDER_4_COL = 'DOCX\DOCX_4_COL'
DOCX_FOLDER_NO_FOOTER = 'DOCX\DOCX_NO_FOOTER'
DOCX_FOLDER_NO_FOOTER_4_COL = 'DOCX\DOCX_NO_FOOTER_4_COL'

# Расположение файлов  эксель из которых берутся данные
EXCEL_FOLDER = 'XLSX'

# Начальные номера таблиц  используемые при замене номеров
# если номера таблиц меняются, то Сноска в номере (для файлов источников) удалится (либо не меняй, либо делай сноску руками)
CITY_TABLE_START_NUM = 1.6          #номер таблицы в фомате число - точка - число
CITYSOURCE_TABLE_START_NUM = 1.22   #номер таблицы в фомате число - точка - число
SKLAD_TABLE_START_NUM = 2.6         #номер таблицы в фомате число - точка - число
SKLADSOURDE_TABLE_START_NUM = 2.22  #номер таблицы в фомате число - точка - число

# Условие, менять ли номера таблиц в ворд файле (еслиномера меняются, то сноски в названии таблиц удаляются)
CITY_TABLE_NUM_EDIT = True
CITYSOURCE_TABLE_NUM_EDIT = True
SKLAD_TABLE_NUM_EDIT = True
SKLADSOURDE_TABLE_NUM_EDIT = True

#Параметры высоты ячеек используемых в шаблонах ворд
HEIGHT_OF_CELL_IN_HEADER = 26
HEIGHT_OF_CELL_IN_SOURCE = 20
HEIGHT_OF_CELL_IN_NOT_SOURCE = 12