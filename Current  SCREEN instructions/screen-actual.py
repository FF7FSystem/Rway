import json
"""Данная скрипта берет 2 файла, проверка актуальности по сайтам и проверку аткуальности по скринам,
для каждого источника сравнивает  значения регулярок (и отрицательных и положительных), в случае если регулярки
для источников разные,  заменяет регулярки скринов  регулярками для сайтов и записывает новый файл "Новые скрены". 
"""


def readcontent(filename):  #открытие файлов
    with open (filename,encoding='utf-8') as inp:
        data=json.load(inp)
    return data

def data_replacement(screen,actual):    #Сравнение словарей файла актуализации и скривон
    info=[]
    for i in screen:
        if i in actual:
            if screen[i]['expr'] != actual[i]['expr']:
                screen[i]['expr'] = actual[i]['expr']
                info.append(i)
            if screen[i]['exprn']!= actual[i]['exprn']:
                screen[i]['exprn']=actual[i]['exprn']
                info.append(i)
    print('Актуализированы инструкции следующих сайтов:')
    print(sorted(set(info)))
    return screen   #Актуализированный словарь данных

def main():
    screen_data = readcontent(r'screen.json')   #Файл скрина со старыми данными
    actual_data = readcontent(r'actuality.json')    #Файл актуальных данных
    repl_data=data_replacement(screen_data['actuality_check']['regs'],actual_data['actuality_check']['regs'])
    screen_data['actuality_check']['regs']=repl_data
    with open(r'new_screen.json', 'w', encoding='utf-8') as out:        #актуализированный файл скринов
        json.dump(screen_data, out, indent=4, ensure_ascii=False)

if __name__ == '__main__':
    main()