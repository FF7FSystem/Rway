import pandas as pd
import re
from termcolor import cprint
from geocode import geo_data,geo_subject
import json

def url_serarch(filename=r'', sheet=''):
    '''
    Выгрузка из Эксель файла 1 столбца содержащего ссылки на страцницы, выделение из сылок конкретной части (города)
    :param filename:  Путь к файлу + имя
    :param sheet:  Названия закладки
    :return: список найденных городов/областей или других фрагментов из урла
    '''
    result=[]   # Результирующий список
    if filename.endswith('.xlsx'):
        tab = pd.read_excel(filename, sheet_name=sheet)[['Ссылка на источник информации']]    #Выгрузка столбца из эксель через пандас
        pat=r'youla\.ru/([\w-]+)/'
        for i in range(len(tab)):
            temp_res=re.findall(pat,tab.loc[i]['Ссылка на источник информации'])            #поиск в каждой строке по шаблону регулярки
            if temp_res:                                                                    # Если что-то найдено добавить в результирующий списко
                result.append(temp_res[0])
    
    elif filename.endswith('.txt'):
        with open(filename) as data:
            pat=r'[\w-]+(?=.n1)'                                                                 #Регулярка поиска городов и субъектов
            for i in data:
                temp_res=re.findall(pat,i)                                                       #поиск в каждой строке по шаблону регулярки
                if temp_res:                                                                    # Если что-то найдено добавить в результирующий списко
                    result.append(temp_res[0])

    cprint(('Всего получено городов',len(result)),'green')
    result=set(result)
    cprint(('Всего получено УНИКАЛЬНЫХ городов',len(result)),'green')
    return result

def dict_1c():
    '''
    Выгрузка из текстового файла (данных словаря 1С) значений субъектов в формате 1С
    :return:
    '''
    with open(r'subject.json', encoding='utf-8') as file:
        return json.load(file)

def search_sub(city,dict):
    '''
    Функция переберает список с городами полученных в фунции url_serarch, передает их в Яндекс АПИ (описаное в файле geocode),
     яндекс возвращает субъект. Далее в субъекте от яндекса ищется слово  по ключу словаря 1С, в результат подставляются значения ключа.
    Данные переносятся в 1С путем копипаста с результатов принта
    :param city: Список городов
    :param dict:   Словарь субъектов в формате 1С
    :return: ничего не возращает, итоговый вывод на экран
    '''
    nores=[]
    for i in sorted(city):
        srch= geo_subject(city=i).lower()   #Поиск субъекта через яндекс АПИ
        res=''
        for j in dict:                      #Перебор ключевых слов  по субъектам для определения значения из словаря 1С
            res_mach=re.findall(r'(?:^|\W){}(?:\W|$)'.format(j),srch)
            if res_mach:
                res = ('ИначеЕсли СтрНайти(ССЫЛКА_ИСТОЧНИКА, "/{}") Тогда Результат = "{}";'.format(i,dict[j]))
        if res:
            print(res)      #Если по ключевым словам нашлось значение
        else:
            nores.append('Значение: '+ srch+' в словаре 1С не найдено')
    print('_________не найденое_________')
    for i in nores:
        print(i)

if __name__ == "__main__":
   #city_lst = url_serarch(filename=r'0001-0153-0032 2019-08-13.xlsx', sheet='SHEET')
   city_lst = url_serarch(filename=r'imp.txt')
   lexicon = dict_1c()
   search_sub(city_lst,lexicon)