import os
from ftplib import FTP
import datetime

def ftp_connect(folder=''):
    host = '10.199.13.39'
    login = 'screenshot'
    password = 'SCREEN12345#'
    ftp = FTP(host)
    ftp.login(login,password)
    current_folder = folder if folder else ([i for  i in ftp.nlst() if i != 'OBJ_TEST_1'])[-1] # войти в указанную папку или последнюю из списка
    ftp.cwd(current_folder) #Зайти в папку
    return ftp.nlst(),ftp   #выгрузить список файлов и сам ftp, чтобы его применять в функции загрузки (иначе переменная объявлена только в данной функции)

def create_folders(sources):
    path = os.getcwd()
    #сохранение в одну общую папку
    folder_deep_1 =  str(datetime.date.today())
    '' if os.path.exists(folder_deep_1) else os.mkdir(folder_deep_1)

    # сохранение каждого источника в свою папку
    # for source in sources:
    #     path_source = os.path.join(path, folder_deep_1,source)
    #     '' if os.path.exists(path_source) else os.makedirs(path_source)
    # return folder_deep_1

def load_in_folder(source,files,ftp):
    source_files = sorted([i for i in files if i.find(source) != -1])
    source_files = source_files[:3] if len(source_files) > 3 else source_files
    for file in source_files:
        # path_for_save = os.path.join(os.getcwd(), str(datetime.date.today()), source, file) #сохранение в одну общую папку
        path_for_save = os.path.join(os.getcwd(), str(datetime.date.today()),file) #сохранение в одну общую папку
        handle = open(path_for_save, 'wb')
        ftp.retrbinary(f'RETR {file}', handle.write)

def load_one_source(source,files,ftp):
    # Скачать все файлы по конкретному источнику
    source_files = sorted([i for i in files if i.find(source) != -1])

    for file in source_files:
        path_for_save = os.path.join(os.getcwd(), str(datetime.date.today()), source) #сохранение в одну общую папку
        print(path_for_save)
        path_for_save_full = os.path.join(path_for_save,file)  # сохранение в одну общую папку
        '' if os.path.exists(path_for_save) else os.makedirs(path_for_save)

        handle = open(path_for_save_full, 'wb')
        ftp.retrbinary(f'RETR {file}', handle.write)

files,ftp = ftp_connect()
sources = sorted(set([i.split('&')[-1].replace('.jpeg','') for i in files]))
create_folders(sources)
for source in sources:
    load_in_folder(source, files,ftp)

load_one_source('IMLS',files,ftp)
