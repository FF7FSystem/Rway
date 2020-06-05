import pyodbc
import time
import concurrent.futures
import threading
from termcolor import cprint


def task_execute(tsk_num):

    task_beg=time.time()
    conn = pyodbc.connect('DRIVER={SQL Server};SERVER=10.199.13.60;DATABASE=rway;UID=vdorofeev;PWD=Sql01')
    cursor = conn.cursor()
    lock = threading.RLock()
    with lock:
        print(f'задача {tsk_num} начата')
    try:
        cursor.execute(f'''
        Select  count(case when Наименование = 'Цена' and Значение <> 0 then  1 end) From    (SELECT Предложение FROM [rway].[dbo].[_Task62] task left join [rway].[РегистрСведений].[ПредложенияЗадач] predl on _IDRRef = predl.Задача where
        _Fld198 = '{tsk_num}') num_pred join (select Объект, Наименование, case when Значение_Число > 0 then  1 when Значение_Строка <> '' then  1  when Значение_Дата is not null then  1 end Значение from [rway].[РегистрСведений].[ЗначенияХарактеристик] Xap left join [rway].[ПВХ].[Характеристики] Xap_val on Xap.Характеристика = Xap_val.Ссылка) xap_tab_pred
        on num_pred.Предложение = xap_tab_pred.Объект ''')
        print(f'задача {tsk_num} окончена с временеем = ', time.time() - task_beg)
        # cprint(cursor.fetchall(),'green')
        return cursor.fetchall()
    except Exception as e:
        cprint (f'Ошибка задачи {tsk_num} {e}','red')
    # with lock:

def main():
    """
    Данный код создает поток для каждого вызова СИНХРОННОЙ функции делая общее выполнение асинхранным
    """

    # conn = pyodbc.connect('DRIVER={SQL Server};SERVER=10.199.13.60;;DATABASE=rway;UID=vdorofeev;PWD=Sql01')
    # Нельзя передавать одно подключение в потоки, первый поток выполнится а другие скажут что соединение занято
    # cursor = conn.cursor()
    # Нельзя передавать в потоки один курсор, первый поток выполнится а другие сойдут с ума т.к. курсор переместится в первом потоке


    tasks_spis = ['0001-0601-0001', '0001-0601-0002', '0001-0601-0003', '0001-0601-0004', '0001-0601-0005']
    begin=time.time()

    with concurrent.futures.ThreadPoolExecutor() as tpe:
        res=tpe.map(task_execute,tasks_spis)
        # f1 = tpe.submit(task_execute,tasks_spis[1])
        # f2 = tpe.submit(task_execute,tasks_spis[2])
        # f3 = tpe.submit(task_execute,tasks_spis[3])
        # print(f1.result())
        # print(f2.result())
        # print(f3.result())

    # print(list(res))
    print(list(res))
    print(time.time()-begin)

if __name__ == "__main__":
    main()  # передать номер в формате строки"