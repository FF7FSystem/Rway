#Запрос к базе для получения списка всех задач из ЗАДАНИЯ (типа 0001-0646') и некоторых характеристики к ним (_IDRRef,ДатаНачала,ДатаОкончания,Код задачи,Псевдоним,Статус,Шаблон)
ALL_TASKS = """
        SELECT _IDRRef,ДатаНачала,ДатаОкончания,_Fld198 as 'Код задачи' ,Псевдоним, jour.Статус,_shablon.Наименование as 'Шаблон'
        FROM [rway].[dbo].[_Task62] task
        left join [rway].[Документ].[Задание] doc on task._Fld192RRef = doc.Ссылка and task._Fld231RRef=0xAB47107B44B1669011E990CEFFC204B6
        left join [rway].[Справочник].[Источники] sourc on task._Fld197RRef = sourc.Ссылка
        left join [rway].[РегистрСведений].[ЖурналЗадач] jour on task._IDRRef = jour.Задача
        left join [rway].[Справочник].[ШаблоныИмпорта] _shablon on task._Fld214RRef = _shablon.НастройкаПСОД
        where doc.КодЗадания = '{}'
        """ # and jour.Статус = 2 добавить для получения данных только о выполненых задачах

#Запрос к базе для получения предварительных характеристик по конкретным задачам  (типа ['0001-0646-0001','0001-0646-0002','0001-0646-0003'])
#предварительные характеристики (_IDRRef,ДатаНачала,ДатаОкончания,Код задачи,Псевдоним,Статус,Шаблон)
SOME_TASKS = """
        SELECT _IDRRef,ДатаНачала,ДатаОкончания,_Fld198 as 'Код задачи' ,Псевдоним, jour.Статус,_shablon.Наименование as 'Шаблон'
        FROM [rway].[dbo].[_Task62] task
        left join [rway].[Документ].[Задание] doc on task._Fld192RRef = doc.Ссылка and task._Fld231RRef=0xAB47107B44B1669011E990CEFFC204B6
        left join [rway].[Справочник].[Источники] sourc on task._Fld197RRef = sourc.Ссылка
        left join [rway].[РегистрСведений].[ЖурналЗадач] jour on task._IDRRef = jour.Задача
        left join [rway].[Справочник].[ШаблоныИмпорта] _shablon on task._Fld214RRef = _shablon.НастройкаПСОД
        where _Fld198 in {}
        """
#запрос в базу для получения количества предложений по данной задаче
CNT_OFFERS_PER_TASK = 'Select COUNT(Предложение) from [rway].[РегистрСведений].[ПредложенияЗадач] where Задача= {}'

#Первый запрос характеристик из базы данных по конкретной задачи
MAIN_QUERY_CHAR_1 = """
            Select  count(case when Наименование = 'Цена' and Значение <> 0 then  1 end) 'Цена',
                    count(case when Наименование = 'Цена Предложения За 1 кв.м.' and Значение <> 0 then  1 end) 'Цена Предложения За 1 кв.м.',
                    count(case when Наименование = 'Размерность Стоимости' and Значение <> 0 then  1 end) 'Размерность Стоимости',
                    count(case when Наименование = 'Продавец' and Значение <> 0 then  1 end) 'Продавец',
                    count(case when Наименование = 'Телефон Продавца' and Значение <> 0 then  1 end) 'Телефон Продавца',
                    count(case when Наименование = 'Общая Площадь Предложения' and Значение <> 0 then  1 end) 'Общая Площадь Предложения',
                    count(case when Наименование = 'Размерность Площади' and Значение <> 0 then  1 end) 'Размерность Площади',
                    count(case when Наименование = 'Широта По Источнику' and Значение <> 0 then  1 end) 'Широта По Источнику',
                    count(case when Наименование = 'Долгота По Источнику' and Значение <> 0 then  1 end) 'Долгота По Источнику',
                    count(case when Наименование = 'Заголовок' and Значение <> 0 then  1 end) 'Заголовок',
                    count(case when Наименование = 'Способ Реализации' and Значение <> 0 then  1 end) 'Способ Реализации',
                    count(case when Наименование = 'Этаж' and Значение <> 0 then  1 end) 'Этаж',
					count(case when Наименование = 'Количество Этажей' and Значение <> 0 then  1 end) 'Этажность',
					count(case when Наименование = 'Тип объекта недвижимости' and Значение <> 0 then  1 end) 'Тип объекта недвижимости',
					count(case when Наименование = 'Предмет Сделки' and Значение <> 0 then  1 end) 'Предмет Сделки',
					count(case when Наименование = 'Класс Объекта' and Значение <> 0 then  1 end) 'Класс Объекта',
                    count(case when Наименование = 'Назначение Объекта Предложения' and Значение <> 0 then  1 end) 'Назначение Объекта Предложения',
                    count(case when Наименование = 'Дата Сбора Информации' and Значение <> 0 then  1 end) 'Дата Сбора Информации'
            From
            (SELECT Предложение FROM [rway].[dbo].[_Task62] task left join [rway].[РегистрСведений].[ПредложенияЗадач] predl on _IDRRef = predl.Задача where _Fld198 = '{}') num_pred
            join
            (select Объект, Наименование,
                case
                    when Значение_Число > 0 then  1
                    when Значение_Строка <> '' then  1
                    when Значение_Дата is not null then  1
                    when Значение <> 0x00000000000000000000000000000000 then  1
                end Значение

                 from [rway].[РегистрСведений].[ЗначенияХарактеристик] Xap left join [rway].[ПВХ].[Характеристики] Xap_val on Xap.Характеристика = Xap_val.Ссылка) xap_tab_pred
            on num_pred.Предложение = xap_tab_pred.Объект
                        """

#Второй запрос характеристик из базы данных по конкретной задачи
MAIN_QUERY_CHAR_2 = """
                SELECT	count(case when [АдресПредставление] <> '' then  1 end) 'Адрес',
                        count(case when [ДатаРазмещения] is not null then  1 end) 'ДатаРазмещения',
                        count(case when [Описание] <> '' then  1 end) 'Описание',
                        count(case when [СамаяРанняяДатаРазмещения] is not null then  1 end) 'СамаяРанняяДатаРазмещения',
                        count(case when [Сегмент] = 0x00000000000000000000000000000000 then null else  1 end) 'Сегмент',
						count(case when [Подсегмент] = 0x00000000000000000000000000000000 then null else  1 end) 'Подсегмент',
                        count(case when [СсылкаИсточника] <> '' then  1 end) 'СсылкаИсточника',
                        count(case when [Субъект] = 0x00000000000000000000000000000000 then null else  1 end) 'Субъект',
                        count(case when [ТипРынка] = 0x00000000000000000000000000000000 then null else  1 end) 'ТипРынка',
                        count(case when [ТипСделки] = 0x00000000000000000000000000000000 then null else  1 end) 'ТипСделки'
                FROM [rway].[dbo].[_Task62] task
                left join [rway].[РегистрСведений].[ПредложенияЗадач] predl on _IDRRef = predl.Задача
                join [rway].[Справочник].[ПредложенияОбъектовНедвижимости] predl_val1 on Предложение = predl_val1.Ссылка
                where _Fld198 = '{}' """

#Запрос почти полность повторяет ALL_TASKS, за тем исключением, что в нем из запроса удалены несколько ненужных полей (ДатаНачала,ДатаОкончания).
#Используется для запроса количества предложений по каждой задачи (вот этот запрос CNT_OFFERS_PER_TASK),
# Ну а почему бы и нет.

ALL_TASK_CNT_FOR_PLOT = """
    SELECT _IDRRef,_Fld198 as 'Код задачи' ,Псевдоним,jour.Статус,_shablon.Наименование as 'Шаблон'
	 FROM [rway].[dbo].[_Task62] task
    left join [rway].[Документ].[Задание] doc on task._Fld192RRef = doc.Ссылка and task._Fld231RRef=0xAB47107B44B1669011E990CEFFC204B6
    left join [rway].[Справочник].[Источники] sourc on task._Fld197RRef = sourc.Ссылка
    left join [rway].[РегистрСведений].[ЖурналЗадач] jour on task._IDRRef = jour.Задача
	left join [rway].[Справочник].[ШаблоныИмпорта] _shablon on task._Fld214RRef = _shablon.НастройкаПСОД
    where doc.КодЗадания = '{}'
     """

# Из текущих заданий парсинга выбирает последнюю созданную коммерческий парсинг (не парсинг важных людей), возвращает номер задачи в формате  0001-0756
NUM_LAST_PARS_TASK = """SELECT TOP 1 task.КодЗадания 
                    FROM [rway].[Документ].[Задание] as task 
                    Where task.Наименование LIKE'%коммерческая вторичная%' AND task.Наименование NOT LIKE '%люди%' 
                    ORDER BY 'КодЗадания' DESC
                    """