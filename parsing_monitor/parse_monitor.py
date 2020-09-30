#описание работы модулы в конце скрипта
import datetime
import json
import os
import re
import subprocess
import dash
import dash_core_components as dcc
import dash_html_components as html
import dash_table
import pandas as pd
import parse_monitor_config as CONFIG
import plotly.express as px
import plotly.graph_objects as go
import sys
from dash.dependencies import Input, Output ,State
from plotly.subplots import make_subplots
from sqlalchemy import create_engine  # для связи с бд и построяния графиков из модуля мониторинга

external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']
app = dash.Dash(__name__, external_stylesheets=external_stylesheets, suppress_callback_exceptions=True)

# Блок CSS настроек визуализации
tabs_styles = {
    'border-top-left-radius': '3px',
    'background-color': '#f9f9f9',
    'padding': '0px 14px',
    'border-bottom': '1px solid #d6d6d6'
}
tab_style = {
    'color':'#586069',
    'border-top-left-radius': '3px',
    'border-top-right-radius': '3px',
    'border-top': '3px solid transparent',
    'border-left': '0px',
    'border-right': '0px',
    'border-bottom': '0px',
    'background-color':' #fafbfc',
    'padding': '12px ',
    'font-family': "system-ui",
    'display': 'flex ',
    'align-items': 'center',
    'justify-content': 'center'


}
tab_selected_style = {
    'border-top-left-radius': '3px',
    'border-top-right-radius': '3px',
    'color': 'black',
    'box-shadow': "1px 1px 0px white",
    'border-left': '1px solid lightgrey',
    'border-right': '1px solid lightgrey',
    'border-top': '3px solid #e36209'
}
table_style = {'margin': '1%'}
thead_style = {'font-weight': 'bold', 'text-decoration': 'underline'}
td_1_style = {'width': '100%', 'vertical-align': 'top'}
td_2_style = {'width': '40%'}
button_style = {'margin-top': '25px', 'margin-bottom': '25px','font-weight': 'bold'}
span_style = {'margin-left': '25px', 'margin-right': '25px', }
div_no_info_tab_4_style = {'margin-left': '25px', 'margin-right': '25px', 'margin-top': '25px', 'fontWeight': 'bold'}
input_style  = {'margin-left': '25px', 'margin-right': '25px', }
message_div_style = {'margin-left': '25px','color':'tomato'}

STATUS_D = {1: ' Исполняется', 2: ' Выполнена', 3: ' Ошибка выполнения', 0: ' К исполнению', 4: ' Парсинг не запущен'} #Словарь перевода состояния парсинга задач
subprocess.Popen([sys.executable, r'parse_monitor_data.py'])    #Запуск скрипта получения по каждой задачи в задании количества предложений и записы в файл


def data_in_all_json_true(task_num):
    """
    Считывает все файлы формата json из папки parse_result и формирует
        1) Dateframe по текущему заданию task_num
        2) Dateframe по всем остальным заданиям
    :param task_num: текущий номер задания
    :return: выше написано
    """
    current_folder = os.getcwd()
    folder_path = current_folder if 'parse_result' in current_folder else f'{current_folder}\parse_result'
    files = sorted(os.listdir(folder_path), reverse=True)  # список файлов
    data = pd.DataFrame()
    for file in files:  #Перебирает все json файлы и сливает в один Dateframe
        file_name, extension = os.path.splitext(file)
        if extension == '.json':
            inp = pd.read_json(os.path.join(folder_path, file), encoding='utf8')
            data = data.append(inp, sort=True, ignore_index=True).fillna(0)

    data = data.loc[data['Дата начала'] != 0].rename(columns={'Дата начала': 'Дата_начала'}) #переименовываем столбец (только те строки где дата начала не  = 0)
    data = data.assign(Дата_начала=pd.to_datetime(data['Дата_начала'])) #Заменяем столбец с датой указанный строкой, столбцом с датой указаной в формате datetime
    data = data.astype({'Статус выполнения': 'int'}) #меняем тип столбца
    data['Статус выполнения'] = data.apply(lambda row: STATUS_D[row['Статус выполнения']], axis=1) #Заменяем  значения  статуса в столбце на словарные (числа меняем строками)

    data_current = data.loc[data['Код задачи'].str.contains(f"{task_num}")]  # Данные по текущей задаче
    data_rest = data[~data['Код задачи'].str.contains(f"{task_num}")]  # Данные по всем остальным задачам
    return data_current, data_rest


def prepare_data(statistic_data):
    """
    Получение статистических данных
    data_drop_excess - Усеченные данные по заданиям (для графиков)
    prepare_mean - значения по всем характеристикам каждой задачи (для таблицы)
    prepare_median - Медианы по всем характеристикам каждой задачи (для таблицы)
    :param statistic_data: Данные по всем заданиям ранее спарсенных
    :return:
    """
    data_drop_excess = statistic_data[['Псевдоним', 'Дата_начала', 'Всего предложений']]  # .rename(columns={'Дата начала':'Дата_начала'}).fillna(0)
    # data_drop_excess = data_drop_excess.query('(Дата_начала != 0) & (Псевдоним != 0)') #Исключение  полей где время нулевое (ломает преобразование времени)
    # data_drop_excess = data_drop_excess.assign(Дата_начала = pd.to_datetime(data_drop_excess['Дата_начала']).dt.strftime('%Y-%m-%d %H:%M:%S'))
    data_drop_excess = data_drop_excess.assign(Дата_начала=pd.to_datetime(data_drop_excess['Дата_начала']))
    prepare_mean = pd.pivot_table(data_drop_excess, columns=['Псевдоним'], values='Всего предложений', aggfunc='mean')
    prepare_median = pd.pivot_table(data_drop_excess, columns=['Псевдоним'], values='Всего предложений', aggfunc='median')
    return data_drop_excess, prepare_mean, prepare_median


def monitor_data(only_max_val=False):
    """
    Считывает данные подготовленные скриптом parse_monitor_data - это данные по каждой задаче например:
    Номер задачи, название, когда начилась, сколько спарсила СПИСОК. по этому списку позже формируется график.
    предусмотрено, чтобы функция возвращала только максимальное значение "всего предложений" по каждой задаче (не список)
    :param only_max_val: Флаг, возвращать в характеристики список значений или только максимальное значение
    :return: Dateframe
    """
    all_data = pd.read_json(CONFIG.FILE_SAVE_NAME, encoding='utf8')
    all_data = all_data.loc[all_data['Дата начала'] != 0].rename(columns={'Дата начала': 'Дата_начала'})  # Исключение  полей где время нулевое (ломает преобразование времени)
    all_data = all_data.assign(Дата_начала=pd.to_datetime(all_data['Дата_начала']))
    if only_max_val: #Если флаг, то берутся только максимальные занчения по хар-ке "Всего предложений"
        need_data = all_data[['Псевдоним', 'Дата_начала', 'Всего предложений']]
        need_data = need_data.assign(**{'Всего предложений': all_data['Всего предложений'].apply(lambda row: max(row))})
    else: #Иначе берется список занчений по хар-ке "Всего предложений"
        need_data = all_data[['Псевдоним', 'Дата_начала', 'Всего предложений', 'Статус выполнения']]
        need_data = need_data.astype({'Статус выполнения': 'int'})
        need_data['Статус выполнения'] = need_data.apply(lambda row: STATUS_D[row['Статус выполнения']], axis=1)
    return need_data.sort_values(by='Псевдоним', axis=0)


def prepare_table_data(current, rest, source):
    """
    Подготовка данных для таблицы.
    1)Берутся Статистические данные по всем характеристикам из текущего задания по конкретному источнику source (current_source), форматируются
    2)Подготавливаются средние (current_source_mean) и медианные (current_source_median) данные по всем характеристикам из
    данных по ранее спарсинным задачам  по конкретному источнику source(Dateframe 'rest')
    В data_for_table сливаются данные по текущему источнику, и разница текущих статистических данных с медианными и средними.
    В title_d не табличные данные которые необходимо удалить из таблицы но отразить на экране типа : Номер задачи, дата старта и т.д. По ним не может быть медианы и прочего
    :param current: DateFrame  Статистические данные по всем характеристикам по по всем задачам текущегой задания
    :param rest:    DateFrame по всем остальным заданиям
    :param source:  Текущий источник (задание)
    :return: DateFrame ТАблицы для отражения на странице и Словарь с нетаблицными данными
    """
    list_for_title = ["Код задачи", "Дата_начала", "Дата окончания", "Псевдоним","Статус выполнения"]  # Для удаления из итоговых таблицы (так как ломает сортировку и ненужны) + добавление в словарь  зоголовков таблицы
    current_source = current.assign(Дата_начала=pd.to_datetime(current['Дата_начала']).dt.strftime('%Y-%m-%d %H:%M:%S'))
    current_source = current_source.loc[current_source['Псевдоним'] == source]
    title_d = {i: current_source[i].values[0] for i in list_for_title}
    #Статистика по текущему заданию по конкретному источнику
    current_source = current_source.drop(list_for_title, axis=1).T.reset_index()
    current_source = current_source.rename(columns={current_source.columns[0]: 'Хар-ка', current_source.columns[1]: 'Значение'})
    #Среднее значение по характеристикам из ранее спарсинных заданий по конкретному источнику
    current_source_mean = rest.loc[rest['Псевдоним'] == source].drop(list_for_title, axis=1).mean().reset_index()
    current_source_mean = current_source_mean.rename(columns={current_source_mean.columns[0]: 'Хар-ка', current_source_mean.columns[1]: 'Среднее_значение'})
    #Медианное значение по характеристикам из ранее спарсинных заданий по конкретному источнику
    current_source_median = rest.loc[rest['Псевдоним'] == source].drop(list_for_title, axis=1).median().reset_index()
    current_source_median = current_source_median.rename(columns={current_source_median.columns[0]: 'Хар-ка', current_source_median.columns[1]: 'Медиана'})
    #Формирование итоговой таблицы путем склеивания данных по конкретному истонику, разницы медианы и средней, удаление ненужных столбцов
    data_for_table = current_source.merge(current_source_mean, left_on='Хар-ка', right_on="Хар-ка")
    data_for_table = data_for_table.merge(current_source_median, left_on='Хар-ка', right_on="Хар-ка")
    data_for_table['Разница_медиана'] = (data_for_table['Значение'].astype('float') - data_for_table['Медиана']).round(2)
    data_for_table['Разница_среднее'] = (data_for_table['Значение'].astype('float') - data_for_table['Среднее_значение']).round(2)
    data_for_table = data_for_table.drop(['Среднее_значение', 'Медиана'], axis=1).sort_values(['Значение', 'Хар-ка'],
                                                                                              ascending=False)
    return data_for_table, title_d


def create_table_page(this_task_data, past_task_data,past_task_data_cut):
    """
    Функция построения структуры страницы с таблицей и графиком.
    Верстка страницы в части расположения таблицы данных,
    графика и т.д. организована с помощь таблицы, в строки  и столюцы которой помещаются нужные HTML элементы.
    Из текущего задания this_task_data, извлекаются все источники (задачи)
    По каждому источнику формируется таблица, график, и еще несколько блоков и справочной информацией и добавляется в
    итоговый список контента данной странички content_list.
    Построение графика:
        по конкретному источнику (задаче) выбираются данные из прошлых парсингов (past_task_data) дополняются данными
        из текущего задания (this_task_data). По этим данным строится график зависимости "Дата парсинга - Всего предложений)
    Построение таблицы статистики (tables_data) происходит в функции prepare_table_data
    Далее График и таблица с данными вкладываются в структуру HTML (табличная верстка), к которой применяются разные
    стили отображения данных. В таблице статистики, текст раскрашиваются разным цветом в зависимости от величины
    значения, разницы с соседним столбцом и т.д. (style_data_conditiona)

    :param this_task_data: DateFrame  Статистические данные по всем характеристикам по по всем задачам текущегой задания
    :param past_task_data: DateFrame по всем остальным заданиям
    :param past_task_data_cut:  DateFrame по всем остальным заданиям с ограниченым количеством столбцов (для графика)
    :return: Список, каждый элемент которого это блоки HTML содержащие таблицы и графики для каждого источника текущего задания
    """
    content_list=[] #список контента данной странички
    sources = sorted(set(this_task_data['Псевдоним']))
    for source in sources: #перебираем источники
        fig_data_past_t = past_task_data_cut[past_task_data_cut['Псевдоним'] == source] #Данные по прошлым парсингам
        fig_data_cur_t = this_task_data.loc[this_task_data['Псевдоним'] == source].loc[:,['Псевдоним', 'Дата_начала', 'Всего предложений']] #Данные по текущему парсингу
        fig_data = fig_data_past_t.append(fig_data_cur_t).sort_values(by='Дата_начала') #Общие данные для графика
        fig = go.Figure(go.Scatter(y=fig_data['Всего предложений'], x=fig_data['Дата_начала']))
        fig.update_layout(xaxis=dict(
            tickmode='array',
            tickvals=fig_data['Дата_начала'],   #Подпись к графику по горизонтали
            ticktext=fig_data['Дата_начала'].dt.strftime('%Y-%m-%d')    #Подпись к графику по горизонтали
            )
        )
        tables_data, header = prepare_table_data(this_task_data, past_task_data, source) #Формирование таблицных данных и ...

        content_list.append(  # добавление таблиц по задачам на страничку
            html.Div([
                html.Table(style=table_style, children=[
                    html.Thead(style=thead_style, children=f'''{header['Псевдоним']}'''),
                    html.Tbody(children=[
                        html.Tr(children=[
                            html.Td(style=td_1_style, children=[
                                html.Div(
                                    children=f'''Задача №{header['Код задачи']} по источнику {header['Псевдоним']}'''),
                                html.Div(
                                    children=f'''Стартовала {header['Дата_начала']}, окончилась {header['Дата окончания']}'''),
                                html.Div(children=f'''Статус выполнения: {header['Статус выполнения']}'''),
                                html.Div(children=[dcc.Graph(id='example-graph', figure=fig)])
                            ]),
                            html.Td(style=td_2_style, children=[
                                html.Div([dash_table.DataTable(
                                    id='datatable-interactivity',
                                    columns=[{"name": i, "id": i} for i in tables_data.columns],
                                    data=tables_data.to_dict('records'),
                                    style_header={'fontWeight': 'bold', 'text-align': 'center'},
                                    style_cell_conditional=[
                                        {'if': {'column_id': 'Хар-ка'}, 'textAlign': 'left'},
                                        {'if': {'column_id': 'Разница_среднее'}, 'textAlign': 'center'},
                                        {'if': {'column_id': 'Разница_медиана'}, 'textAlign': 'center'},
                                        {'if': {'column_id': 'Значение'}, 'textAlign': 'center'}
                                    ],
                                    style_data_conditional=[
                                        {'if': {'row_index': 'odd'},
                                         'backgroundColor': 'rgb(248, 248, 248)'
                                         },
                                        {'if': {'filter_query': '{Разница_медиана} < -3', 'column_id': 'Значение'},
                                         'color': 'tomato',
                                         'fontWeight': 'bold'
                                         },
                                        {'if': {'filter_query': '{Разница_медиана} > -3', 'column_id': 'Значение'},
                                         'color': 'green',
                                         'fontWeight': 'bold'
                                         },
                                        {'if': {'filter_query': '{Разница_медиана} < -3',
                                                'column_id': 'Разница_медиана'},
                                         'color': 'tomato',
                                         'fontWeight': 'bold'
                                         },
                                        {'if': {'filter_query': '{Разница_медиана} > 3',
                                                'column_id': 'Разница_медиана'},
                                         'color': 'green',
                                         'fontWeight': 'bold'
                                         },
                                        {'if': {'filter_query': '{Разница_медиана} > -3 && {Разница_медиана} <3',
                                                'column_id': 'Разница_медиана'},
                                         'color': 'white'
                                         },

                                        {'if': {'filter_query': '{Разница_среднее} < -3',
                                                'column_id': 'Разница_среднее'},
                                         'color': 'tomato',
                                         'fontWeight': 'bold'
                                         },
                                        {'if': {'filter_query': '{Разница_среднее} > 3',
                                                'column_id': 'Разница_среднее'},
                                         'color': 'green',
                                         'fontWeight': 'bold'
                                         },
                                        {'if': {'filter_query': '{Разница_среднее} > -3 && {Разница_среднее} <3',
                                                'column_id': 'Разница_среднее'},
                                         'color': 'white'
                                         },

                                    ],
                                    style_as_list_view=True

                                )])
                            ])
                        ])
                    ])
                ])]))
    return content_list

this_task_data, past_task_data = None,None
past_task_data_cut, data_mean, data_median = None,None,None

def refresh_main_variables():
    """
    Функция обновляет глобальные переменные которые используются в функциях ниже.
    this_task_data - данные по текущей задаче
    past_task_data - данные по всем другим задачам
    past_task_data_cut - Есеченные данные по всем другим задачам (для графиков)
    data_mean   - Средние значения по всем характеристикам каждой задачи (для таблицы)
    data_median - Медианы по всем характеристикам каждой задачи (для таблицы)
    :return: ничего не возвращает
    """
    global this_task_data, past_task_data
    global past_task_data_cut, data_mean, data_median
    this_task_data, past_task_data = data_in_all_json_true(CONFIG.TASK_NUM)
    past_task_data_cut, data_mean, data_median = prepare_data(past_task_data)
refresh_main_variables()

#Описание Базовой странички и закладок на ней
app.layout = html.Div([
    dcc.Tabs(id='tabs',
             style=tabs_styles,
             value='tab-1', #Номер закладки, которая будет открыта по умолчанию
             children=[
                 dcc.Tab(label='Динамика парсинга', value='tab-1', style=tab_style, selected_style=tab_selected_style),
                 dcc.Tab(label='Общая статистика', value='tab-2', style=tab_style, selected_style=tab_selected_style),
                 dcc.Tab(label='Статистика по задачам', value='tab-3', style=tab_style,
                         selected_style=tab_selected_style),
                 dcc.Tab(label=f'Статистика по заданию {CONFIG.TASK_NUM}', value='tab-4', style=tab_style,
                         selected_style=tab_selected_style),
                 dcc.Tab(label='Произвольное задание', value='tab-5', style=tab_style,
                         selected_style=tab_selected_style),
             ]),
    html.Div(id='tabs-content')
])


@app.callback(Output('tabs-content', 'children'), [Input('tabs', 'value')])
def render_content(tab):
    """
    Декоратор и Функция для работы вкладок на базовой страничке
    для каждой вкладки описывается контент, и интервал обновления вкладки
    :param tab:
    :return:
    """
    if tab == 'tab-1':  #Если выбрана вкладка
        return html.Div(children=[      #Содержание вкладки
            dcc.Graph(id='example-graph'),  #график ( обращение к элементам по id)
            dcc.Interval(                   #параметры обновления
                id='interval-component',
                interval=CONFIG.REFRESH_TIME_BROWSE_TAB_1 * 100000,  # in milliseconds
                n_intervals=0
            )
        ]
        )
    elif tab == 'tab-2':    #Если выбрана вкладка
        return html.Div(    #Содержание вкладки
            children=[
                dcc.Graph(id='example-graph2'), #график ( обращение к элементам по id)
                dcc.Interval(
                    id='interval-component2',   #параметры обновления вкладки
                    interval=CONFIG.REFRESH_TIME_BROWSE_TAB_2 * 1000,  # in milliseconds
                    n_intervals=0
                )
            ]
        )
    elif tab == 'tab-3':    #Если выбрана вкладка
        return html.Div(    #Содержание вкладки
            children=[
                dcc.Graph(id='example-graph3'), #график ( обращение к элементам по id)
                dcc.Interval(                   #параметры обновления вкладки
                    id='interval-component3',
                    interval=CONFIG.REFRESH_TIME_BROWSE_TAB_3 * 1000,  # in milliseconds
                    n_intervals=0
                )
            ]
        )
    elif tab == 'tab-4':    #Если выбрана вкладка
        return html.Div(children=[  #Содержание вкладки
                html.Div( id='example-data'),   #Див блок( обращение к элементам по id)
                dcc.Interval(
                id='interval-component4',    #параметры обновления вкладки
                interval=CONFIG.REFRESH_TIME_BROWSE_TAB_4 * 1000,  # in milliseconds
                n_intervals=0
            )
        ])


    elif tab == 'tab-5':    #Если выбрана вкладка
        #Формирует 5-ю вкладку "Произвольное задание" на которой размещены поле для ввода, кнопка, подчеркивание
        # и основной Div блок, в котором отражается либо предупреждение, либо  таблицы-графики, продробнее в  output-state
        return html.Div(children=[
            dcc.Input(id='input-1-state', type='text', placeholder=f"формат {CONFIG.TASK_NUM}",style=input_style),
            html.Button(id='submit-button-state', n_clicks=0, children='Поиск',style=button_style),
            html.Hr(),
            html.Div(id='output-state')
                        ]

        )



@app.callback(Output('example-graph', 'figure'), [Input('interval-component', 'n_intervals')])
def update_graph_live(n):
    """
    Формирует 1 вкладку "Динамика парсинга". Берет DateFrame dinamic_data, это данные мониторинга характеристики
    "всего предложений" по текущему заданию взятого из конфига и строит график по каждому задаче(источнику).
    :param n: не используется, но должен быть
    :return: возвращает график с множеством подграфиков
    """
    dinamic_data = monitor_data()
    imes_cont = len(dinamic_data)
    col_max = 5 #Определить что максимум будет 5 графиков в строке
    row_max = (imes_cont // col_max) + 1 if imes_cont % col_max != 0 else imes_cont // col_max # определить сколько будет строк с кграфиками исмходя из количества источников
    row_col_list = [[j, i] for j in range(1, row_max + 1) for i in range(1, col_max + 1)]   #вспомогательный Список списков номеров строк и столбцов  для формирования таблицы графиков
    titles = [row['Псевдоним'] + row['Статус выполнения'] for index, row in dinamic_data.iterrows()]    #Формирования заголовков к графикам

    fig = make_subplots(rows=row_max, cols=col_max, subplot_titles=titles) #Создать матрицу графиков
    for index in range(len(dinamic_data)):  #перебрать все строки (источники) из входного dateframe и
        #построить график для каждого источника row_col_list[index][0],row_col_list[index][1] - позиция графика в матрице графиков
        fig.append_trace(go.Scatter(x=dinamic_data.iloc[index]["Всего предложений"], mode='lines+markers',name=dinamic_data.iloc[index]["Псевдоним"]), row_col_list[index][0],row_col_list[index][1])
    fig.update_layout(showlegend=False, height=row_max * 400, title_text="Скорость парсинга каждого источника") #Опции Легенду не ваыводить, Ширина строки  = кол-во строк *400, общий заголовок
    return fig


@app.callback(Output('example-graph2', 'figure'), [Input('interval-component2', 'n_intervals')])
def update_graph_live(n):
    """
    Формирует 2-ю вкладку "Общая статистика". Берет DateFrame dinamic_data по текущему заданию взятого из конфига
    (для каждого источника указано максимальноезначение сперсенных предложений only_max_val=True) и строит график
    по каждому задаче(источнику).  Чтобы графики с малым числом предложений не терялись на фоне "больших"
    задач вывод разбит на части.
    :param n: не используется, но должен быть.
    :return: возвращает график с множеством подграфиков
    """
    data_max_cnt = monitor_data(only_max_val=True).sort_values("Всего предложений") #Получение данных с флагом, сортировка

    col_max = 5 #определение количества столбцов с графиками
    col_max = col_max if len(data_max_cnt) // col_max > 1 else 1 #определение количества столбцов с графиками
    fig = make_subplots(rows=1, cols=col_max)   #построение списка графиков

    for i in range(col_max):# Может быть 1 или 5 столбцов
        part = int(len(data_max_cnt) / col_max)
        data_parse_part = data_max_cnt[part * i:part * (i + 1)] if i < col_max - 1 else data_max_cnt[part * i:] #Входные данные разбиваются на срезы (для каждого столбца)
        data_median_filter = data_median[list(data_parse_part['Псевдоним'])].T.reset_index()
        data_mean_filter = data_mean[list(data_parse_part['Псевдоним'])].T.reset_index()
        fig.add_trace(go.Bar(x=data_parse_part["Псевдоним"], y=data_parse_part["Всего предложений"], name=''), 1, i + 1)    # Для каждого среза строится гистограмии типа Источник/максимольное значние предложений
        fig.add_trace(go.Scatter(x=data_median_filter["Псевдоним"], y=data_median_filter["Всего предложений"], name='Медиана'), 1,i + 1)    #На этом же граифке строится медианное значение хар "всего предложений" из справсенных данных  предидущих заданий
        fig.add_trace(go.Scatter(x=data_mean_filter["Псевдоним"], y=data_mean_filter["Всего предложений"],name='Среднее значение'), 1, i + 1)  #На этом же граифке строится средние значение хар "всего предложений" из справсенных данных  предидущих заданий
    fig.update_layout(barmode='stack', showlegend=False, height=600,
                      title_text=f"Результаты парсинга по каждому источнику на данный момент (разбит на {col_max} графиков, чтобы источники с малым кол-вом не  терялись на общем фоне)")  #Опции Легенду не ваыводить, Ширина строки  = кол-во строк *400, общий заголовок
    return fig


@app.callback(Output('example-graph3', 'figure'), [Input('interval-component3', 'n_intervals')])
def update_graph_live(n):
    """
    Формирует 3-ю вкладку "Статистика по задачам".
    Данные монитора данных DateFrame dinamic_data, в котором указаны только максимальные значения предложений
    для каждого источника (одним числом), объединям с past_task_data_cut.
    past_task_data_cut это Dateframe содержащий значение хар-к "всего задач" для каждого источника  полученные
    в предидущих парсингах.
    По полученным данным строим графики которые показывают динамику хар-ки "всего предложений" для разных парсингов.
    :param n: не используется, но должен быть.
    :return: возвращает график с множеством подграфиков
    """
    monitor_data_values = monitor_data(only_max_val=True)
    data_update = past_task_data_cut.append(monitor_data_values)  # В прошлый результаты парсинга добавляются новые по текущей задаче

    data_update = data_update.loc[data_update['Псевдоним'].isin(monitor_data_values['Псевдоним'])]  # Выбрать данные по задачам (источникам) которые есть в мониторе (на случай когда в мониторе т.е. на текущий момент парсится не все источники и по ним статистику не выводить)
    all_source = sorted(set(list(data_update['Псевдоним'])))    #Все источники в алфавитном порядке
    imes_cont = len(set(list(data_update['Псевдоним'])))    #Количество источников
    col_max = 5 #количество колонок
    row_max = (imes_cont // col_max) + 1 if imes_cont % col_max != 0 else imes_cont // col_max #определяем количество строк
    row_col_list = [[j, i] for j in range(1, row_max + 1) for i in range(1, col_max + 1)] #вспомогательный Список списков номеров строк и столбцов  для формирования таблицы графиков

    fig = make_subplots(rows=row_max, cols=col_max, subplot_titles=all_source) #Создать матрицу графиков
    for num, source in enumerate(all_source):
        data_source = data_update.loc[data_update['Псевдоним'] == source].sort_values("Дата_начала") #Из общего data_update, бертся данные по конкретному ичтонику. Данные сортируеются по дате
        fig.append_trace(go.Scatter(x=data_source["Дата_начала"], y=data_source["Всего предложений"], mode='lines'),
                         row_col_list[num][0], row_col_list[num][1])
    height_row = 400 if row_max < 2 else 200
    fig.update_layout(showlegend=False, height=row_max * height_row, title_text='Характеристика "Всего предложений" по каждому источнику  за период.')
    return fig


@app.callback(Output('example-data', 'children'), [Input('interval-component4', 'n_intervals')])
def dead_space(n):
    """
    Формирует 4-ю вкладку "Статистика по заданию {номер}".
    my_big_div_family = [] это список, в который будет добавляться контент странички по блокам. каждый блок это html
    блок содержащий  таблицу и график по конкретному источнику.
    Для построения даннных используются глобальные переменные:
    this_task_data, past_task_data,past_task_data_cut - данные которые считываются с диска. Для их обновления вызывается
    функция refresh_main_variables. Если данных по текущему заданию нет т.е. this_task_data - пустая, то на экране
    формируется уведомление. При нажатии на кнопку запускается подпроцесс (файл parsing_manage.py  -
    получение данных из sql, сохранение на диск) В последующем при обновлении вкладки считываются свежие данные
    переменной this_task_data и контент вкладки отображается.

    :param n: не используется, но должен быть.
    :return: контент
    """
    refresh_main_variables()    #Обновление переменных (считывание с диска), которые содержат статистические данные.
    my_big_div_family = []      #Список, который будет содеражать  HTML блоки  контента вкладки

    if not this_task_data.empty:    #если по текущей задаче (из конфига) ЕСТЬ ДАННЫЕ  (ранее запрашивалась статистика из SQL и сохранена на диске)
        my_big_div_family.append(  # добавление кнопки  и пояснения на страничку
            html.Div([
                html.Span(children=f'Загрузить статистику по задачам  из базы SQL', style=span_style),
                html.Button('load...', id='btn-nclicks-1', n_clicks=0, style=button_style),
                html.Hr()])),

        my_big_div_family.extend(create_table_page(this_task_data, past_task_data,past_task_data_cut)) # добавление таблицы и графика на страницу
    else:    #если по текущей задаче (из конфига) НЕТ  ДАННЫХ
        my_big_div_family.append(   # добавление кнопки  и пояснения на страничку
            html.Div(children=[
                html.Div(children=f'''По задаче {CONFIG.TASK_NUM} ниразу не выгружались статистические данные.
                Для  выгрузки нажмите кнопку ниже. Выгрузка осущестлвяется около 5 минут.''',style=div_no_info_tab_4_style),
                html.Span(children=f'Загрузить статистику по задачам  из базы SQL', style=span_style),
                html.Button('load...', id='btn-nclicks-1', n_clicks=0, style=button_style),
                html.Hr()]))


    return my_big_div_family


@app.callback(Output('btn-nclicks-1', 'children'), [Input('btn-nclicks-1', 'n_clicks')])
def displayClick(btn1):
    """
    Фунция следит за нажатием кнопки "Загрузить" на вкладке  "Статистика по заданию №№№№-№№№№"
    Если Кнопку нажать, то запустится подпроцесс parsing_manage.py, и по текущему заданию будет отправлен запрос в SQL.
    Полученные данные из SQL Сохранятся на диск ( и будут считаны при обновлении вкладки "Статистика по заданию №№№№-№№№№")
    :param btn1:
    :return:
    """
    changed_id = [p['prop_id'] for p in dash.callback_context.triggered][0]
    if 'btn-nclicks-1' in changed_id:
        button_msg = 'Загружается...'
        subprocess.Popen([sys.executable, r'parsing_manage.py'])
    else:
        button_msg = 'Загрузить'
    return html.Div(button_msg)


@app.callback(Output('output-state', 'children'), [Input('submit-button-state', 'n_clicks')], [State('input-1-state', 'value')])
def update_output(n_clicks, input1):
    """
    Проверяется поле ввода, которое возвращает input1 (номер задания).
    содержимое input1 проверяется на правильность заполенния
    Если в input1 номер текущего задания, тогда обновляются основные переменные для формирования таблиц и графиков
    (search_task_data, other_task_data, other_task_data_cut, _data_mean, _data_median), и строится контент
    странички с таблицой и графиками по текущему заданию с помощью функции  create_table_page.
    Если в input1 номер не равный номеру текущего задания, то  основные переменные для формирования таблиц и графиков
    (search_task_data, other_task_data, other_task_data_cut, _data_mean, _data_median) формируются по новому номеру
    задания указанного в input1.
    По указанным переменным формируется контент странички (графики, таблицы) с помощью функции create_table_page
    :param n_clicks:
    :param input1:
    :return:
    """
    if input1:
        is_task = re.findall('[\d]{4}-[\d]{4}$', input1)
        try:
            is_task=is_task[0]
        except:
            return html.Div(children='Предупреждение: Номер задачи неверного формата (верный 0000-0000).',style=message_div_style)
        if is_task:
            if input1 != CONFIG.TASK_NUM: #в input1 номер не равный номеру текущего задания
                search_task_data, other_task_data = data_in_all_json_true(input1)
                other_task_data_cut, _data_mean, _data_median = prepare_data(other_task_data)
                if not search_task_data.empty:
                    return create_table_page(search_task_data, other_task_data,other_task_data_cut)
                else:
                    return html.Div(children='Предупреждение: Нет такого номера задачи среди ранее спарсиных!',style=message_div_style)
            elif input1 == CONFIG.TASK_NUM: #в input1 номер текущего задания
                return create_table_page(this_task_data, past_task_data, past_task_data_cut)
        else:
            return html.Div(children='Предупреждение: Что-то пошло не так...',style=message_div_style)

    else:
        return html.Div(children='Предупреждение: Номер задачи не указан.',style=message_div_style)


if __name__ == '__main__':
    #Запуск сервера приложения
    app.run_server(debug=True,host='0.0.0.0',port='5000')
    """
    Агоритм:
    
    """

