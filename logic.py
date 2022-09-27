import pandas as pd
import pprint
from statistics import mean

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
        Необходимо ввести полную ссылку до файла, например "C:\Users\Eugene\Downloads\data1.xls".
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

        # t = t_mean_day["T"].values.tolist() лист со значениями температуры (пока не нужен)
        self.t_mean_day["0"] = self.t_mean_day.index
        list_of_dates = pd.DataFrame(self.t_mean_day["0"]).reset_index()
        list_of_dates = list_of_dates.drop(columns="data")
        list_of_dates["0"] = pd.to_datetime(list_of_dates["0"])
        fact_list_of_dates = pd.date_range(start=start_chain, end=end_chain)
        # fact_list_of_dates1 = fact_list_of_dates1.loc[~fact_list_of_dates1["0"].isin(list_of_dates)]
        # start = time.time()
        missing_dates = fact_list_of_dates.difference(list_of_dates["0"])
        missing_dates = pd.DatetimeIndex(missing_dates).sort_values()
        # delete an already unnecessary column with dates
        list_of_missing_dates = list(missing_dates.astype(str).tolist())
        lenth_of_list_of_missing_dates = len(list_of_missing_dates)
        # print(fact_list_of_dates)
        # checking that all values in an index are the same
        # list_of_dates = set(list_of_dates)
        # print(T_meanday.head())
        print(self.t_mean_day["T"])

        # print(T_meanday.head())
        print(start_chain_with_desc)
        print(end_chain_with_desc)
        print(days_with_temp_with_desc)
        print(total_cols_with_desc)
        #print(numbers_of_missing_days_with_desc)

        # test2005_2022.sort_index(inplace=True)
        # print(test2005_2022.index[-1])
        # print(test2005_2022.head())
        print(f"Количество дней с пропущенными данными: {lenth_of_list_of_missing_dates}")
        pprint.pprint(f"Список дат с отсутствующими данными по температуре: {list_of_missing_dates}")
        # self.T_meanday.head().plot(kind = "bar", subplots = False, sharex = False, figsize=(40, 20))
        # plt.show()

    def save_dataset_mean_temp(self):
        name_of_set_mean_temp = input("Введите название файла для сохранения:")
        self.t_mean_day["T"].to_excel(f"{name_of_set_mean_temp}.xlsx")
        print("Файл сохранен!")

    def get_average_monthly_temperature(self):
        t_mean_day_for_month = pd.DataFrame(self.t_mean_day["T"])
        t_mean_day_for_month.index = pd.to_datetime(t_mean_day_for_month.index)
        self.average_monthly_temperature = t_mean_day_for_month.resample("M").mean()
        print(self.average_monthly_temperature)

    def save_dataset_average_monthly_temperature(self):
        self.average_monthly_temperature = pd.DataFrame(self.average_monthly_temperature)
        name_of_average_monthly_temperature = input("Введите название файла для сохранения:")
        self.average_monthly_temperature.to_excel(f"{name_of_average_monthly_temperature}.xlsx")
        print("Файл сохранен!")

    def heating_period_treatment(self):
        day_date_plus_temp = pd.DataFrame(self.t_mean_day["T"])
        prepare_start_heating_date = pd.to_datetime(input(
            "Введите начальную дату обработки отопительного периода в формате гггг-мм-дд:"))
        prepare_end_heating_date = pd.to_datetime(input(
            "Введите конечную дату обработки отопительного периода в формате гггг-мм-дд:"))
        # исследуемый диапазон измерений отопительного периода (по умолчанию с 1 сентября по 1 июня)
        interesting_heating_period = day_date_plus_temp.loc[prepare_start_heating_date:prepare_end_heating_date]
        interesting_heating_period.reset_index(inplace=True)
        list_temp = interesting_heating_period["T"].values.tolist()
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

        interesting_heating_period["average_five_day_temperature"] = mean_five_temp

        ds_real_start_heating_date = interesting_heating_period.loc[(
                interesting_heating_period.average_five_day_temperature < 8)]
        self.real_start_heating_date = ds_real_start_heating_date.iloc[4, 0]
        real_start_heating_date_with_desc = f"Дата начала отопительного периода согласно законодательству:" \
                                            f"{self.real_start_heating_date}"

        help_end_heating_date = pd.to_datetime(prepare_end_heating_date - pd.DateOffset(months=3))
        interesting_heating_period["data"] = pd.to_datetime(interesting_heating_period["data"])
        interesting_heating_period = interesting_heating_period.set_index("data")
        ds_real_end_heating_date = interesting_heating_period.loc[help_end_heating_date:prepare_end_heating_date]
        add_ds_real_end_heating_date = ds_real_end_heating_date.loc[(
                ds_real_end_heating_date.average_five_day_temperature > 8)].reset_index()
        self.real_end_heating_date = pd.to_datetime(add_ds_real_end_heating_date.loc[0, 'data'])
        self.real_end_heating_date = self.real_end_heating_date.date()

        real_end_heating_date_with_desc = f"Дата окончания отопительного периода согласно законодательству:" \
                                          f"{self.real_end_heating_date}"
        self.duration_heating_period = self.real_end_heating_date - self.real_start_heating_date
        duration_heating_period_with_desc = f"Продолжительность отопительного периода {self.duration_heating_period} дней"
        self.ds_duration_heating_period = interesting_heating_period.loc[self.real_start_heating_date:self.real_end_heating_date]

        self.average_temperature_heating_period = self.ds_duration_heating_period["T"].mean()
        self.average_temperature_heating_period = round(self.average_temperature_heating_period, 2)
        average_temperature_heating_period_with_desc = f"Средняя температура отпительного периода равна " \
                                                       f"{self.average_temperature_heating_period} градусов"
        self.gsop = (18 - self.average_temperature_heating_period) * self.duration_heating_period.days
        self.gsop = round(self.gsop, 0)
        gsop_with_desc = f"Реальные градусосутки отопительного периода равны {self.gsop} "
        self.min_temp_day_of_heat_temp = self.ds_duration_heating_period['T'].min()
        min_temp_day_of_heat_temp_with_desc = f"Минимальная температура наиболее холодных суток " \
                                              f"отопительного периода: {self.min_temp_day_of_heat_temp}"
        self.min_temp_five_day_of_heat_temp = self.ds_duration_heating_period['average_five_day_temperature'].min()
        min_temp_five_day_of_heat_temp_with_desc = f"Минимальная температура наиболее холодной пятидневки " \
                                                   f"отопительного периода: {self.min_temp_five_day_of_heat_temp}"
        str_with_min_temp_day_of_heat_temp = self.ds_duration_heating_period[self.ds_duration_heating_period['T'] ==
                                                                             self.ds_duration_heating_period['T'].min()]
        str_with_min_temp_five_day_of_heat_temp = self.ds_duration_heating_period[
                                                  self.ds_duration_heating_period.average_five_day_temperature ==
                                                  self.ds_duration_heating_period.average_five_day_temperature.min()]

        # print(interesting_heating_period)
        # print(interesting_heating_period.info())
        # print(ds_real_start_heating_date)
        # print(type(self.real_start_heating_date))
        # print(ds_real_end_heating_date)
        # print(help_end_heating_date)
        # print(prepare_end_heating_date)
        # print(add_ds_real_end_heating_date)
        # print(self.real_end_heating_date)
        # print(type(real_end_heating_date))

        print(self.ds_duration_heating_period)
        print(real_start_heating_date_with_desc)
        print(real_end_heating_date_with_desc)
        print(duration_heating_period_with_desc)
        print(min_temp_day_of_heat_temp_with_desc)
        print(str_with_min_temp_day_of_heat_temp)
        print(min_temp_five_day_of_heat_temp_with_desc)
        print(str_with_min_temp_five_day_of_heat_temp)
        print(average_temperature_heating_period_with_desc)
        print(gsop_with_desc)

        # print(list_temp)
        # print(list_five_temp)
        # print(mean_five_temp)
        # print(day_date_plus_temp.info())
        # print(str(start_heating_date))
        # print(str(end_heating_date))

    def save_ds_duration_heating_period(self):
        self.ds_duration_heating_period = pd.DataFrame(self.ds_duration_heating_period)
        name_of_ds_duration_heating_period = input("Введите название файла для сохранения:")
        self.ds_duration_heating_period.to_excel(f"{name_of_ds_duration_heating_period}.xlsx")
        print("Файл сохранен!")

    def save_data_about_heating_period(self):
        one_data_heat_period = (self.real_start_heating_date.strftime('%m/%d/%Y'),
                                self.real_end_heating_date.strftime('%m/%d/%Y'),
                                self.duration_heating_period.days, self.min_temp_day_of_heat_temp,
                                self.min_temp_five_day_of_heat_temp, self.average_temperature_heating_period,
                                self.gsop)
        self.data_heat_period.append(one_data_heat_period)
        pprint.pprint(self.data_heat_period)
        #print(type(self.real_start_heating_date))


