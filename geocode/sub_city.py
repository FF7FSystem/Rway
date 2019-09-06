import pandas as pd
import re
from termcolor import cprint
from geocode import geo_data,geo_subject
import codecs

def url_serarch(filename=r'', sheet=''):
    '''
    Выгрузка из Эксель файла 1 столбца содержащего ссылки на страцницы, выделение из сылок конкретной части (города)
    :param filename:  Путь к файлу + имя
    :param sheet:  Названия закладки
    :return: список найденных городов/областей или других фрагментов из урла
    '''
    result=[]   # Результирующий список
    tab = pd.read_excel(filename, sheet_name=sheet)[['Ссылка на источник информации']]    #Выгрузка столбца из эксель через пандас
    pat=r'youla\.ru/([\w-]+)/'
    for i in range(len(tab)):
        temp_res=re.findall(pat,tab.loc[i]['Ссылка на источник информации'])            #поиск в каждой строке по шаблону регулярки
        if temp_res:                                                                    # Если что-то найдено добавить в результирующий списко
            result.append(temp_res[0])
    cprint(('Всего получено городов',len(result)),'green')
    result=set(result)
    cprint(('Всего получено УНИКАЛЬНЫХ городов',len(result)),'green')
    return result

""" Выгрузка из текстового файла (данных 1С)  по которым находятся в шаблонах ПСОД города для сравнения городов,
    которые учтены в 1С и тех по котороым не находится субъект 
    c1_dict=[]                              #
    with open(r"1C.txt") as read_file:
        pat1=r'\"/([\w-]+)\"'
        pat2=r'Результат = \"(.+)\"'
        for i in read_file:
            c1_res=re.findall(pat1,i)
            if c1_res:
                state=re.findall(pat2,i)
                if state:
                    c1_dict.append([c1_res[0],state[0]])
    
    c1_dict=dict(c1_dict)
    print (c1_dict)
    c1_set={i for i in c1_dict}
raz=result-c1_set # Разница в уникальных значениях ссылок
cprint(len(raz),'red')
"""


def dict_1c():
    '''
    Выгрузка из текстового файла (данных словаря 1С) значений субъектов в формате 1С
    :return:
    '''
    fileObj = codecs.open("subject_1c.txt", "r", "utf_8_sig") #Перестал читать данные из файла (руский шрифт)
    sub_list= [ i.strip() for i in fileObj]                   #Пришлось применить import codecs
    fileObj.close()
    return sub_list

def search_sub(city,dict):
    '''
    Функция переберает список с городами, передает их в Яндекс АПИ описаное в файле geocode, яндекс возвращает субъект
    Из списка субъектов 1С ищется Субъект полученный в Яндексе, заполняется строка для 1С.
    Данные переносятся в 1С путем копипаста с результатов принта
    :param city: Список городов
    :param dict:   Словарь субъектов в формате 1С
    :return: ничего не возращает, итоговый вывод на экран
    '''
    nores=[]
    for i in sorted(city):
        # print(i,'---->',geo_subject(city=i))
        srch= geo_subject(city=i).split()[0]
        res=''
        for j in dict:
            if j.find(srch)>-1: #Если субъект из списка городов  которые вернул яндекс нашелся в списке субъектов 1С тогда
               res = ('ИначеЕсли СтрНайти(ССЫЛКА_ИСТОЧНИКА, "/{}") Тогда Результат = "{}";'.format(i,j))
        if res:
            print(res)
        else:
            nores.append('Значение: '+ srch+' в словаре 1С не найдено')

    print('_________не найденое________________')
    for i in nores:
        print(i)




if __name__ == "__main__":
   city_lst = url_serarch(filename=r'0001-0153-0032 2019-08-13.xlsx', sheet='SHEET')
   lexicon = dict_1c()
   search_sub(city_lst,lexicon)