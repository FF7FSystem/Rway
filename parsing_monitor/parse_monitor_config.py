TASK_NUM='0001-0771' # Номер посделейднего задания
TIME_TO_REFRESH=300 #Периодичность получения 1С  количества спарсиных предложений по каждой задаче (В секундах)
FILE_SAVE_NAME = r'monitor_data.json' #Файл, где скапливается данные для монитора количества предложений по  задаче
FOR_CONNECT = 'DRIVER={SQL Server};SERVER=10.199.13.60;DATABASE=rway;UID=vdorofeev;PWD=VD12345#'
TIME_FIND_NEW_NUM_TASK = 1800 #Периодичность получения в 1С номера последней задачи парсинга (В секундах)


#выводить в консоль статистику по задачам полученную из базы данных (модуль parsing_manage)
PRINT_ALLTASK_STATISTIC = True
#выводить в консоль сообщения о начале и окончании получения данных из дазы 1С по каждой задаче (модуль parsing_manage)
PRINT_ALLTASK_PROCESS = False
#выводить ли в консоль дату прохода цикла по получению количества всех предложений в каждой задаче (модуль parse_monitor_data)
PRINT_MONITOR_DATA_GO = False

#Время обновление данных  на вкладках в браузере (в секундах).
REFRESH_TIME_BROWSE_TAB_1 = 300
REFRESH_TIME_BROWSE_TAB_2 = 300
REFRESH_TIME_BROWSE_TAB_3 = 300
REFRESH_TIME_BROWSE_TAB_4 = 300
