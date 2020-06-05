import pyodbc
import time
from concurrent.futures import ThreadPoolExecutor
import asyncio
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

async def main(tasks_spis):
    # conn = pyodbc.connect('DRIVER={SQL Server};SERVER=10.199.13.60;;DATABASE=rway;UID=vdorofeev;PWD=Sql01')
    # Нельзя передавать одно подключение в потоки, первый поток выполнится а другие скажут что соединение занято
    # cursor = conn.cursor()
    # Нельзя передавать в потоки один курсор, первый поток выполнится а другие сойдут с ума т.к. курсор переместится в первом потоке

    begin=time.time()
    futures = [loop.run_in_executor(executor,task_execute, args) for args in tasks_spis]
    await asyncio.gather(*futures)

    # print(list(res))
    print(time.time()-begin)

if __name__ == "__main__":
    tasks_spis = ['0001-0601-0001', '0001-0601-0002', '0001-0601-0003', '0001-0601-0004', '0001-0601-0005']

    executor = ThreadPoolExecutor()
    loop = asyncio.get_event_loop()
    loop.run_until_complete(main(tasks_spis))
