import pandas as pd
from tkinter import *
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText 

ViewOrder_0 = 0
ViewOrder_1 = 0
def clicked():
    text.delete(1.0, END)
    #label['text'] = ''    
    df0 = pd.read_excel('ЦА_ТО_Сведения_загрузки_данных.xlsx', sheet_name='TDSheet')     #  Учреждение, КоличествоНеЗагруженных, Ошибка, ВидФайла
    df_ppk = df0[df0['Ошибка']. str.contains('Начисление    Персональный повышающий коэффициент', na= False)]
    print('Персональный повышающий коэффициент\n', df_ppk ['Ошибка'])

    values = df_ppk['Ошибка'].tolist()
    print(len(df0))
    df0 = df0[df0.Ошибка.isin (values) == False]  # удаление(списком)строк содержащих PPK
    print(len(df0))


    df_Otmena = df0[df0['Ошибка']. str.contains('Приказ об отмене не загружен. Не найдены сведения о приказе, который требуется отменить', na= False)]   # Приказы об отмене
#print('222 \n', df_)
    df_Otmena = df_Otmena['Ошибка'].tolist()
    print(len(df_Otmena), '  - Приказы об отмене (всего)')
    print(len(df0))
    for i2 in range(len(df_Otmena)): 
        poz1 = df_Otmena[i2].index(' идентификатор: ')
        poz2 = df_Otmena[i2].index('Приказ об отмене не загружен. Не найдены сведения о приказе, который требуется отменить')
        ID = df_Otmena[i2][poz1 + 16:poz2 - 3:]
    #print(ID)
        df0 = df0[df0.ИдентификаторОбъекта != ID]  # удаление (построчно) строк содержащих ID
    print(len(df0))


    ViewOrder_1 = df0[df0['Ошибка'].str.contains('Приказ о назначении начислений не загружен.', na= False)]
    ViewOrder_1 = ViewOrder_1[['НомерОбласти','Учреждение','ИдентификаторУчреждения','ЦБ','КодСВР','GUIDЗапроса','ВидФайла','ИдентификаторОбъекта','Ошибка']]
    print(len(ViewOrder_1), '  - Ожидаем_ ViewOrder_1')
    ViewOrder_1.to_excel("Ожидаем_ ViewOrder_1_.xlsx", index=False)

    ViewOrder_0 = df0[df0['Ошибка'].str.contains('Приказ об изменении начислений не загружен.', na= False)]
    ViewOrder_0  = ViewOrder_0 [['НомерОбласти','Учреждение','ИдентификаторУчреждения','ЦБ','КодСВР','GUIDЗапроса','ВидФайла','ИдентификаторОбъекта','Ошибка']]
    print(len(ViewOrder_0), '  - Ожидаем_ ViewOrder_0')
    ViewOrder_0.to_excel("Ожидаем_ ViewOrder_0_.xlsx", index=False)

    s ='Ожидаем_ ViewOrder_0_.xlsx! - ' + str(len(ViewOrder_0)) + ' строк.\n''Ожидаем_ ViewOrder_1_.xlsx! - ' + str(len(ViewOrder_1)) + ' строк'
    text.insert(3.0, s)


root = Tk()
root.title('Формирование файлов в Техподдержку ЕИСУКС')
root.geometry('450x300')

frame = ttk.Frame(borderwidth=1, relief=SOLID, padding=[8, 10])
frame.pack(anchor=NW, fill=X, padx=5, pady=5)
Button(frame, text="СТАРТ", command=clicked).pack(side=LEFT)
Button(frame, text="Закрыть", command=root.destroy).pack(side=LEFT)
label = Label(text='END')
#label = Label(text='Ожидаем_ ViewOrder_0_.xlsx! - ' + str(len(ViewOrder_0)) + ' строк')
#label1 = Label(text='Ожидаем_ ViewOrder__.xlsx! - ' + str(len(ViewOrder_1)) + ' строк')
#label.pack()
#label1.pack()
text = ScrolledText(width=50, height=15, wrap="word") #  вертикальная прокрутка тхт-окна
text.insert(3.0, 'Для запуска программы нажмите СТАРТ.\n\n \n  \
    Программа работает с Отчётом =ЦА_ТО_Сведения_загрузки_данных=.\n\
    Отчёт должен находиться в одной папке с приложением.')
text.pack()
root.mainloop()
