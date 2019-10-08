import requests
import json
from termcolor import cprint

def geo_data(city,keys=['25d93e53-6cba-4529-a36b-9d602d13a749']):
    '''
    Поиск геодыных (город, регион, страна, координаты) по строке (городу)
    :param city: строка по которой осуществляется поиск
    :param keys: Набор ключей для подключение к АПИ Яндекс
    :return:    Возвращает словарь с ключем: "сторока по которой осуществлялся поиск", Значением "весь набор полученых от яндекс"
    '''
    for key in keys:
        r = requests.get('https://geocode-maps.yandex.ru/1.x/?apikey={}&format=json&geocode={}'.format(key,city))
        if r.status_code == 200:
            try:
                data = json.loads(r.content)
            except Exception as e:
                cprint(e,'red')
            return  {city:data}
        else:
            print(r)

def geo_subject(city,keys=['25d93e53-6cba-4529-a36b-9d602d13a749']):
    '''
    Вспаомогательная функция возвращает из данных функции geo_data только значение субъекта (Например: "Алтакйский Край")
    :param city: строка по которой осуществляется поиск
    :param keys: Набор ключей для подключение к АПИ Яндекс
    :return: значение субъекта
    '''
    temp=geo_data(city,keys)
    try:
        result=temp[city]['response']['GeoObjectCollection']['featureMember'][0]['GeoObject']['metaDataProperty'][
            'GeocoderMetaData']['AddressDetails']['Country']['AdministrativeArea']['AdministrativeAreaName']
    except Exception as e:
        cprint((e,'Для региона: '+ city+' не найден субъект'),'green')
        result='Noresult'

    #Если не сделать это условие то яндекс возврящает "Тюменская обл", что не верно.
    if result =='Тюменская область'   and   city == 'hanty-mansiysk':
        result = 'Ханты-Мансийский Автономный округ - Югра АО'
    return(result)

    return(result)

if __name__ == "__main__":
    geo_data(city='maykop',)
