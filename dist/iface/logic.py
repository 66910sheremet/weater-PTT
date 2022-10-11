import pandas as pd
import pprint
from statistics import mean
import openpyxl


pd.options.mode.chained_assignment = None


class Processing:
    """Класс Processing является реализацией программы. Содержит в себе методы для дальнейшей обработки
    датасета с температурой. Для обработки может использоваться .xls или .xlsx файл, скачанный, например, с
    сайта https://rp5.ru/. Чтобы программа могла работать с файлом необходимо удалить шапку перед таблицами
    на листе с данными температур, так, чтобы лист начинался с названий столбцов. Столбец с датами обрабатываемого
    периода необходимо назвать "data", столбец с температурами необходимо назвать "T"."""
    def __init__(self, t_mean_day=0, average_monthly_temperature=0, ds_duration_heating_period=0, data_heat_period=[],
                 real_start_heating_date=0, real_end_heating_date=0,
                 duration_heating_period=0, min_temp_day_of_heat_temp=0,
                 min_temp_five_day_of_heat_temp=0, average_temperature_heating_period=0,
                 gsop=0):
        self.t_mean_day = t_mean_day
        self.average_monthly_temperature = average_monthly_temperature
        self.ds_duration_heating_period = ds_duration_heating_period
        self.data_heat_period = data_heat_period
        self.real_start_heating_date = real_start_heating_date
        self.real_end_heating_date = real_end_heating_date
        self.duration_heating_period = duration_heating_period
        self.min_temp_day_of_heat_temp = min_temp_day_of_heat_temp
        self.min_temp_five_day_of_heat_temp = min_temp_five_day_of_heat_temp
        self.average_temperature_heating_period = average_temperature_heating_period
        self.gsop = gsop

    def preliminary_processing(self):
        """Метод preliminary_processing служит для предварительной обработки файлов типа .xls или .xlsx.
        Необходимо ввести полную ссылку до файла, например r"C:\\Users\\Eugene\\Downloads\\data1.xls".
         В качестве выходных данных реализуется таблица из двух колонок первая из которых дата,
          а вторая средняя температура за день. Также выводится начальная дата исследуемого периода,
          конечная дата исследуемого периода, количество дней между начальной и конечной датами,
          фактическое количество дней с данными по температуре, количество дней с пропущенными датами,
          список дат с отсутствующими данными по температуре с виде списка."""
        # Предлагается ввести ссылку на интересующий файл
        link_input = input("Введите ссылку на интересующий файл:")
        link = link_input.replace("\\", "/")
        # Открытие файла на чтение
        data = pd.read_excel(link)
        # Сохранение двух колонок для дальнейшей обраобтки
        test2005_2022 = data[["data", "T"]]
        # Приведение формата дат от "гггг:мм:дд чч:мм" к "гггг:мм:дд"
        test2005_2022["data"] = pd.to_datetime(test2005_2022["data"], format="%d.%m.%Y %H:%M").dt.date
        # Подсчет средней температуры за день и сортировка по дате
        self.t_mean_day = test2005_2022.groupby("data").agg({"T": "mean"})
        # Начальная дата исследиемого периода
        start_chain = self.t_mean_day.index[0]
        # Конечная дата исследуемого периода
        end_chain = self.t_mean_day.index[-1]
        # Вывод на печать начальной даты исследуемого периода с описанием
        start_chain_with_desc = f"Начальная дата исследуемого периода: {self.t_mean_day.index[0]}"
        # Вывод на печать конечной даты исследуемого периода с описанием
        end_chain_with_desc = f"Конечная дата исследуемого периода: {self.t_mean_day.index[-1]}"
        # Нахождение количества дней между начальной и конечной датами
        days_with_temp = (end_chain - start_chain)
        # Количество дней между начальной и конечной датами с описанием
        days_with_temp_with_desc = f"Количество дней между начальной и конечной датами: {days_with_temp.days} дней"
        # Нахождение фактического числа дней с данными по температуре
        total_cols = len(self.t_mean_day)
        # Фактическое число дней с данными по температуре с описанием
        total_cols_with_desc = f"Фактическое количество дней с данными по температуре: {total_cols} дней"
        # Нахождение количества пропущенных дней (дней без данных по температуре)
        numbers_of_missing_days = (end_chain - start_chain).days - len(self.t_mean_day)
        # Количество пропущенных дней с описанием
        numbers_of_missing_days_with_desc = f"Количество пропущенных дней временной последовательности " \
                                        f"{numbers_of_missing_days}"
        # Получение листа со значениями температур
        # t = t_mean_day["T"].values.tolist()
        self.t_mean_day["0"] = self.t_mean_day.index
        list_of_dates = pd.DataFrame(self.t_mean_day["0"]).reset_index()
        list_of_dates = list_of_dates.drop(columns="data")
        list_of_dates["0"] = pd.to_datetime(list_of_dates["0"])
        fact_list_of_dates = pd.date_range(start=start_chain, end=end_chain)
        missing_dates = fact_list_of_dates.difference(list_of_dates["0"])
        missing_dates = pd.DatetimeIndex(missing_dates).sort_values()
        list_of_missing_dates = list(missing_dates.astype(str).tolist())
        lenth_of_list_of_missing_dates = len(list_of_missing_dates)

        # Вывод на печать дат со среднедневной температурой в формате дата-температура
        print(self.t_mean_day["T"])
        # Вывод на печать начальной даты исследуемого периода
        print(start_chain_with_desc)
        # Вывод на печать конечной даты исследуемого периода
        print(end_chain_with_desc)
        # Вывод на печать количества дней между начальной и конечной датами
        print(days_with_temp_with_desc)
        # Вывод на печать фактического количества дней с данными по температуре
        print(total_cols_with_desc)
        # Вывод на печать количества дней с пропущенными датами
        print(f"Количество дней с пропущенными данными: {lenth_of_list_of_missing_dates}")
        # Вывод списка дат с отсутствующими данными по температуре
        pprint.pprint(f"Список дат с отсутствующими данными по температуре: {list_of_missing_dates}")

    def save_dataset_mean_temp(self):
        """Метод для сохранения в формате .xls датасета с усредненной температурой.
        Файлы сохраняются в папку с программой"""
        # Ввод названия файла
        name_of_set_mean_temp = input("Введите название файла для сохранения:")
        # Сохранение файла
        self.t_mean_day["T"].to_excel(f"{name_of_set_mean_temp}.xlsx")
        # Вывод на печать сообщения об успешном сохранении файла
        print("Файл сохранен!")

    def get_average_monthly_temperature(self):
        """Метод для получения датасета со среднемесячными температуры"""
        # Получение датафрейма с данными по средней температуре за сутки
        t_mean_day_for_month = pd.DataFrame(self.t_mean_day["T"])
        # Приведение датафрефма к формату datetime
        t_mean_day_for_month.index = pd.to_datetime(t_mean_day_for_month.index)
        # Получение датафрейма с среднемесяными температурами
        self.average_monthly_temperature = t_mean_day_for_month.resample("M").mean()
        # Вывод на печать датафрейма со среднемесяными температурами
        print(self.average_monthly_temperature)

    def save_dataset_average_monthly_temperature(self):
        """Метод для сохранения в формате .xls датасета со среднемесяной температурой.
        Файлы сохраняются в папку с программой"""
        # Получение датафрейма с данными со среднемесяной температурой
        self.average_monthly_temperature = pd.DataFrame(self.average_monthly_temperature)
        # Ввод названия файла
        name_of_average_monthly_temperature = input("Введите название файла для сохранения:")
        # Сохранение файла
        self.average_monthly_temperature.to_excel(f"{name_of_average_monthly_temperature}.xlsx")
        # Вывод на печать сообщения об успешном сохранении файла
        print("Файл сохранен!")

    def heating_period_treatment(self):
        """Метод для обработки конкретного отопительного периода. Для работы программы необходимо ввести
        начальную и конечную даты обработки в формате гггг-мм-дд. Для средней полосы целесообразно
        использование диапазна начиная с 1 сентября исследуемого года, по 1 июня следующего года. Метод
        выводит на печать три колонки:
        1. Дата
        2. Средняя температура за день
        3. Средняя температура за 5 дней посчитанные за предыдущие 5 дат, включая строчку и исследуемой.
        В соответствии с законодательством РФ выводятся даты отопительного периода.
        Выводится дата начала отопительного периода согласно законодательству;
        дата окончания отопительного периода согласно законодательству;
        продолжительность отопительного периода в днях;
        минимальная температура наиболее холодных суток отопительного периода;
        градусосутки отопительного периода
        """
        # Получаем датасет со среднедневными температурами в формате DataFrame
        day_date_plus_temp = pd.DataFrame(self.t_mean_day["T"])
        # Ввод начальной даты для обработки отопительного периода в формате "гггг-мм-дд"
        prepare_start_heating_date = pd.to_datetime(input(
            "Введите начальную дату обработки отопительного периода в формате гггг-мм-дд:"))
        # Ввод конечной даты для обработки отопительного периода в формате "гггг-мм-дд"
        prepare_end_heating_date = pd.to_datetime(input(
            "Введите конечную дату обработки отопительного периода в формате гггг-мм-дд:"))
        # исследуемый диапазон измерений отопительного периода (по умолчанию с 1 сентября по 1 июня)
        interesting_heating_period = day_date_plus_temp.loc[prepare_start_heating_date:prepare_end_heating_date]
        # Сброс индекса датафрейма
        interesting_heating_period.reset_index(inplace=True)
        # Получение листа со среднедневнимы температурами
        list_temp = interesting_heating_period["T"].values.tolist()
        # Алгоритм для расчета наиболее холодной пятидневки
        list_five_temp = []
        while list_temp:
            list_five_temp.append(list_temp[:5])
            del list_temp[:1]
        mean_five_temp = []
        for i in list_five_temp:
            mean_five_temp.append(round(mean(i), 3))
        mean_five_temp.insert(0, 0)
        mean_five_temp.insert(0, 0)
        mean_five_temp.insert(0, 0)
        mean_five_temp.insert(0, 0)
        mean_five_temp = mean_five_temp[:-4]

        # создание колонки для температур, усредненных за пять предыдущих дней
        interesting_heating_period["average_five_day_temperature"] = mean_five_temp
        # Получение датасета со температурой пятидневки меньше 8 градусов цельсия
        ds_real_start_heating_date = interesting_heating_period.loc[(
                interesting_heating_period.average_five_day_temperature < 8)]
        # Получение даты реального начала отопительного периода (температура пятидневки ниже 8 градусов)
        self.real_start_heating_date = ds_real_start_heating_date.iloc[4, 0]
        # Дата начала отопительного сезона с описанием
        real_start_heating_date_with_desc = f"Дата начала отопительного периода согласно законодательству:" \
                                            f"{self.real_start_heating_date}"
        # дата начала анализа для нахождения конца отопительного периода (конец от. периода - 3 месяца)
        help_end_heating_date = pd.to_datetime(prepare_end_heating_date - pd.DateOffset(months=3))
        # Перевод столбца data в формат datatime
        interesting_heating_period["data"] = pd.to_datetime(interesting_heating_period["data"])
        # Установка столба data как индекса
        interesting_heating_period = interesting_heating_period.set_index("data")
        # Диапазон дат от даты начала анализа конца отопительного периода до последней даты исследований
        ds_real_end_heating_date = interesting_heating_period.loc[help_end_heating_date:prepare_end_heating_date]
        # Нахождение дат с температурой пятидневки выше 8 градусов цельсия
        add_ds_real_end_heating_date = ds_real_end_heating_date.loc[(
                ds_real_end_heating_date.average_five_day_temperature > 8)].reset_index()
        # Нахождения даты конца отопительного периода
        self.real_end_heating_date = pd.to_datetime(add_ds_real_end_heating_date.loc[0, 'data'])
        # Удобный вывод даты конца отопительного периода
        self.real_end_heating_date = self.real_end_heating_date.date()
        # Дата конца отопительного периода с описанием
        real_end_heating_date_with_desc = f"Дата окончания отопительного периода согласно законодательству:" \
                                          f"{self.real_end_heating_date}"
        # Расчет продолжительности отопительного сезона
        self.duration_heating_period = self.real_end_heating_date - self.real_start_heating_date
        # Продолжительность отопительного периода с описанием
        duration_heating_period_with_desc = f"Продолжительность отопительного периода {self.duration_heating_period} дней"
        # датасет данных с температурами от начала до конца вычисленного отопительного периода
        self.ds_duration_heating_period = interesting_heating_period.loc[self.real_start_heating_date:self.real_end_heating_date]
        # Нахождение средней температуры отопительного периода
        self.average_temperature_heating_period = self.ds_duration_heating_period["T"].mean()
        # Округление средней температуры отопительного периода до двух знаков после запятой
        self.average_temperature_heating_period = round(self.average_temperature_heating_period, 2)
        # Средняя температура отопительного периода с описанием
        average_temperature_heating_period_with_desc = f"Средняя температура отпительного периода равна " \
                                                       f"{self.average_temperature_heating_period} градусов"
        # Вычисление градусосуток реального отопительного периода
        self.gsop = (18 - self.average_temperature_heating_period) * self.duration_heating_period.days
        # Округление до целых градусосуток реального отопительного периода
        self.gsop = round(self.gsop, 0)
        # Градусосутки отопительного периода с описанием
        gsop_with_desc = f"Реальные градусосутки отопительного периода равны {self.gsop} "
        # Нахождение минимальной температуры наиболее холодных суток отопительного периода
        self.min_temp_day_of_heat_temp = self.ds_duration_heating_period['T'].min()
        # Минимальная температура наиболее холодных суток с описанием
        min_temp_day_of_heat_temp_with_desc = f"Минимальная температура наиболее холодных суток " \
                                              f"отопительного периода: {self.min_temp_day_of_heat_temp}"
        # Нахождение температуры наиболее холодной пятидневки
        self.min_temp_five_day_of_heat_temp = self.ds_duration_heating_period['average_five_day_temperature'].min()
        # Температура наиболее холодной пятидневки с описанием
        min_temp_five_day_of_heat_temp_with_desc = f"Минимальная температура наиболее холодной пятидневки " \
                                                   f"отопительного периода: {self.min_temp_five_day_of_heat_temp}"
        # Наиболее холодные сутки в формате дата + температура
        str_with_min_temp_day_of_heat_temp = self.ds_duration_heating_period[self.ds_duration_heating_period['T'] ==
                                                                             self.ds_duration_heating_period['T'].min()]
        # Наиболее холодная пятидневка в формате дата + температура
        str_with_min_temp_five_day_of_heat_temp = self.ds_duration_heating_period[
                                                  self.ds_duration_heating_period.average_five_day_temperature ==
                                                  self.ds_duration_heating_period.average_five_day_temperature.min()]
        # Вывод датасета рассчитанного отопительного периода в формате дата, т-ра суток, температура пятидневки
        print(self.ds_duration_heating_period)
        # Вывод даты начала посчитанного отопительного периода с описанием
        print(real_start_heating_date_with_desc)
        # Вывод даты конца посчитанного отопительного периода с описанием
        print(real_end_heating_date_with_desc)
        # Вывод продолжительности посчитанного отопительного периода в днях с описанием
        print(duration_heating_period_with_desc)
        # Вывод минимальной температуры наиболее холодных суток для посчитанного отопительного периода
        print(min_temp_day_of_heat_temp_with_desc)
        # Вывод в формате дата - МИНИМАЛЬНАЯ ТЕМПЕРАТУРА СУТОК - пятидневка
        print(str_with_min_temp_day_of_heat_temp)
        # Вывод минимальной температуры пятидневки отопительного периода
        print(min_temp_five_day_of_heat_temp_with_desc)
        # Вывод в формате дата - минимальная температрура суток - МИНИМАЛЬНАЯ ТЕМПЕРАТУРА ПЯТИДНЕВКИ
        print(str_with_min_temp_five_day_of_heat_temp)
        # Вывод средней температуры за отопительный сезон
        print(average_temperature_heating_period_with_desc)
        # Вывод градусосуток отопительного периода
        print(gsop_with_desc)

    def save_ds_duration_heating_period(self):
        """Метод для сохранения в формате .xls датасета отопительного периода.
        Файлы сохраняются в папку с программой"""
        # Получение датафрейма с отопительным периодом в формате DataFrame
        self.ds_duration_heating_period = pd.DataFrame(self.ds_duration_heating_period)
        # Ввод название файла
        name_of_ds_duration_heating_period = input("Введите название файла для сохранения:")
        # Сохранение файла
        self.ds_duration_heating_period.to_excel(f"{name_of_ds_duration_heating_period}.xlsx")
        # Вывод на печать сообщения об успешном сохранении файла
        print("Файл сохранен!")

    def save_data_about_heating_period(self):
        """Метод для сохранения основных данных по расчетному отопительному периоду.
        Сохранение происходит в виде списка, элементами которого является кортеж с данными
        1. Дата начала отопительного сезона
        2. Дата окончания отопительного сезона
        3. Продолжительность отопительного сезона
        4. Минимальная температура отопительного сезона
        5. Температура наиболее холодной пятидневки отопительного сезона
        6. Средняя температура отопительного сезона
        7. Градусосутки отопительного сезона"""
        one_data_heat_period = (self.real_start_heating_date.strftime('%m/%d/%Y'),
                                self.real_end_heating_date.strftime('%m/%d/%Y'),
                                self.duration_heating_period.days, self.min_temp_day_of_heat_temp,
                                self.min_temp_five_day_of_heat_temp, self.average_temperature_heating_period,
                                self.gsop)
        self.data_heat_period.append(one_data_heat_period)
        pprint.pprint(self.data_heat_period)

