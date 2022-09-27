from logic import Processing

work = Processing()
work.preliminary_processing()

while True:
    print("Для сохрания датасета с усредненной температурой, введите 1")
    print("Для вывода списка со среднемесячными температурами введите 2")
    print("Для обработки конкретного отопительного периода нижмите 3")
    choice = input("Enter your choice:")
    if choice == "1":
        work.save_dataset_mean_temp()
    elif choice == "2":
        work.get_average_monthly_temperature()
        print("Сохранить датасет со среднемесячными температурами: Y/N")
        choice_save_ma = input("Enter your choice:")
        if choice_save_ma == "Y":
            work.save_dataset_average_monthly_temperature()
        elif choice_save_ma == "N":
            print("New operations")
    elif choice == "3":
        work.heating_period_treatment()
        print("Для сохранения датасета реального отопительного периода, нажмите 1")
        print("Для сохранения данных отопительного периода, нажмите 2")
        ch_ds_hp = input("Enter your choice:")
        if ch_ds_hp == '1':
            work.save_ds_duration_heating_period()
        elif ch_ds_hp == '2':
            work.save_data_about_heating_period()


