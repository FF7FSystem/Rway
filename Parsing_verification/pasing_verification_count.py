import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime
from termcolor import cprint
import time
from openpyxl import load_workbook
import matplotlib as mpl
import matplotlib.pyplot as plt
import numpy

def task_list_prepare(task_num):
    result_dict={}
    # for_connect = {'server': '10.199.13.60:1433', 'database': 'rway', 'username': 'vdorofeev', 'password': 'Sql01'}
    for_connect = {'server': '10.199.13.57', 'database': 'rway', 'username': 'vdorofeev', 'password': 'VD12345#'}
    engine = create_engine('mssql+pymssql://{username}:{password}@{server}/{database}'.format(**for_connect))
    conn = engine.connect()

    all_task = pd.read_sql(f'''
    SELECT _IDRRef,_Fld198 as 'Код задачи' ,Псевдоним,jour.Статус,_shablon.Наименование as 'Шаблон'
	 FROM [rway].[dbo].[_Task62] task
    left join [rway].[Документ].[Задание] doc on task._Fld192RRef = doc.Ссылка and task._Fld231RRef=0xAB47107B44B1669011E990CEFFC204B6
    left join [rway].[Справочник].[Источники] sourc on task._Fld197RRef = sourc.Ссылка
    left join [rway].[РегистрСведений].[ЖурналЗадач] jour on task._IDRRef = jour.Задача
	left join [rway].[Справочник].[ШаблоныИмпорта] _shablon on task._Fld214RRef = _shablon.НастройкаПСОД
    where doc.КодЗадания = '{task_num}'
     ''', conn) #and jour.Статус = 2

    for i in all_task.iterrows():
        link = i[1]['_IDRRef']
        link = '0x' + link.hex().upper()  # Преобразование из байтового представления в целочисленное для подстановки в селект
        cnt_predl = pd.read_sql(f'''
        SELECT COUNT(Предложение) from [rway].[РегистрСведений].[ПредложенияЗадач] where Задача= {link}''',conn)
        cnt_predl = int(cnt_predl.loc[0][0])
        sub = ''
        if i[1]['Псевдоним'] == 'MOVE':
            if 'ГорМил' in i[1]['Шаблон']:
                sub = '_ГОРМИЛ'
            elif 'Регионы' in i[1]['Шаблон']:
                sub = '_РЕГИОНЫ'
            elif 'Москва' in i[1]['Шаблон']:
                sub = '_МОСКВА'
        name_task = i[1]['Псевдоним']+sub
        num_code = i[1]['Код задачи'][-4:]

        if name_task in result_dict:
            if  result_dict[name_task]['num_code'] < num_code:
                result_dict[name_task]['num_code'] = num_code
                result_dict[name_task]['count'] = cnt_predl
        else:
             result_dict[name_task] = {'num_code':num_code,'count':cnt_predl}
    return result_dict

def load_exel(file):
    xl = pd.read_excel(file, sheet_name='коммерция')
    replace_col_name = {}
    for i in xl.columns:
        if isinstance(i, datetime):
            replace_col_name[i] = i.strftime("%d.%m.%Y")
        elif isinstance(i, str) and i != 'Источники':
            pass
            # replace_col_name[i] = 'Комментарий'
    xl.rename(columns=replace_col_name, inplace=True)
    return xl

def plot_result(spis,psevdo):
    print('start plot_result',psevdo)
    dpi = 80
    fig = plt.figure(dpi=dpi, figsize=(512 / dpi, 384 / dpi))
    mpl.rcParams.update({'font.size': 10})
    # plt.xlabel('x')
    # plt.ylabel('F(x)')

    spis = [i for i in spis if not numpy.isnan(i[0])]
    # spis=[for i in spis if not isinstance(float,)]
    # cprint(spis, 'red')
    val1 = [i[0] for i in spis]
    val2 = [i[1] for i in spis]
    val3 = [i + 1 for i in range(len(spis))]

    # plt.plot(val3,val1)
    plt.bar(val3, val1)
    plt.title((psevdo,spis[-1:]))
    plt.axhline(numpy.mean(sorted(val1, reverse=True)[:int(len(val1) / 3)]), color='r') #среднее значение от трети самых большиъ результатов
    plt.axhline(numpy.median(val1),color='g')                                          # медиана от значений
    plt.xticks(val3, val2, rotation=45)

    # plt.legend(loc = 'upper right')
    plt.show()

def main(task_num):
    new_pars_data=task_list_prepare(task_num)
    for i in sorted(new_pars_data):
        print(i,new_pars_data[i]['count'])
    file = r'C:\Users\user1\Desktop\parse_kommerc.xlsx'
    excel_data=load_exel(file)

    date_now = (datetime.now().strftime("%d.%m.%Y"))
    for_outp = []
    for source in excel_data['Источники']:
        source=source.upper()
        if source in new_pars_data:
            for_outp.append(new_pars_data[source]['count'])
        else:
            for_outp.append(0)

    new_data_for_add=pd.DataFrame({date_now:for_outp})
    # cprint(excel_data,'red')
    # cprint(new_data_for_add, 'red')
    result_dp=excel_data.join(new_data_for_add)

    print(result_dp)

    col = list(result_dp.columns)
    for i in result_dp.iterrows():
        # print(i[1])
        val=i[1].tolist()
        res=[[val[id],item] for id, item in enumerate(col) if '2019' in item  or '2020' in item]
        plot_result(res,i[1][0])


    ## df1.to_excel(file, sheet_name='коммерция') Затирает файл
    ## Снять комент чтобы записать
    with pd.ExcelWriter(file, engine='openpyxl') as writer:
        writer.book = load_workbook(file)
        result_dp.to_excel(writer,index=False)


if __name__ == "__main__":
    # t_x = time.time()
    main('0001-0661')  # передать номер в формате строки"
    # cprint(('Время Выполения функции МАИН', time.time() - t_x), 'blue')