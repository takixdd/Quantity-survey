import pandas as pd
import customtkinter
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import seaborn as sns
import configparser
import os
import sys


def on_closing():
    root.destroy()
    os._exit(1)
    sys.exit()

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("blue")
root = customtkinter.CTk()
root.title("Wykresy Plots")
root.geometry("1900x1000")
upper_frame = customtkinter.CTkFrame(root)
upper_frame.pack(fill='both', expand=True)

if not os.path.exists(r"wykresy_plot_config.ini"):
    config = configparser.ConfigParser()

    config.add_section('rbh')
    config.set('rbh', 'first', '103')

    config.add_section('rok')
    config.set('rok', 'start', '2023')
    config.set('rok', 'end', '2023')

    config.add_section('miesiac')
    config.set('miesiac', 'first', '1')
    config.set('miesiac', 'second', '12')

    config.add_section('material')
    config.set('material', '1', 'Lina stalowa fi 20')
    config.set('material', '2', 'Taśma izolacyjna')
    config.set('material', '3', 'Lina stalowa fi 16')
    config.set('material', '4', 'Acetylen')
    config.set('material', '5', 'Tlen')

    with open(r"wykresy_plot_config.ini", 'w') as configfile:
        config.write(configfile)
try:
    config_obj = configparser.ConfigParser()
    config_obj.read(r"wykresy_plot_config.ini")
    rbh = config_obj["rbh"]
    lata = config_obj["rok"]
    miesiace = config_obj["miesiac"]
    lista_materialow = config_obj["material"]

    cena_rbh = rbh['first']
    y1 = lata['start']
    y2 = lata['end']
    m1 = miesiace['first']
    m2 = miesiace['second']
    material = lista_materialow['1']

except KeyError:
    if os.path.exists(r"wykresy_plot_config.ini"):
        os.remove(r"wykresy_plot_config.ini")

path = r"dfplot.xlsx"
df = pd.read_excel(path, sheet_name=None)

df_final = pd.concat(df.values(), ignore_index=True)

# Atumatycznie dodawanie bazy do box year
try:
    df = pd.read_excel(r'dfall.xlsx')
    df['Data'] = pd.to_datetime(df['Data'], format='%d-%m-%Y', errors='coerce', utc=False)
    df['rok'] = pd.DatetimeIndex(df['Data']).year
    df1 = df['rok'].unique()
    df2 = [x for x in df1 if str(x) != 'nan']
    df2 = sorted(df2)
    df2 = [int(x) for x in df2]
    df2.reverse()

except FileNotFoundError:
    print("Nie odnaleziono bazy danych. Załaduj nowy obmiar, a następnie zrestartuj program.")

year_list = []

try:
    for i in df2:
        year_list.append(str(i))
except NameError:
    print("Brak bazy danych obmiaru.")

# Wyswietlanie listy materiałów z pliku config ini
materialy_lista = []
dodaj_nowe_range = []

for section in config_obj.sections():
    for range, item in config_obj.items(section):
        if section == 'material':
            materialy_lista.append(item)
            dodaj_nowe_range.append(range)


df_final['Wartosc pracy'] = df_final['Rbh'] * int(cena_rbh)
df_materialy = df_final.copy()
df_materialy = df_materialy.loc[(df_materialy['Data'] >= f'2013-01-01') & (df_materialy['Data'] <= f'{y2}-12-31')]

def reload_barplot_y1(box_rok_y1):
    global year_list
    global y1
    if box_rok_y1 in year_list:
        lata['start'] = box_rok_y1
        with open(r"wykresy_plot_config.ini", 'w') as configfile:
            config_obj.write(configfile)
        y1 = lata['start']
    return y1

def reload_barplot_y2(box_rok_y2):
    global year_list
    global y2
    if box_rok_y2 in year_list:
        lata['end'] = box_rok_y2
        with open(r"wykresy_plot_config.ini", 'w') as configfile:
            config_obj.write(configfile)
        y2 = lata['end']
    return y2

def reload_normalplot(material_box):
    global materialy_lista
    global material
    if material_box in materialy_lista:
        material = material_box

def update_normalplot():
    global material
    if entry_material.index("end") == 0:
        normal_plot()
    else:
        material = entry_material.get()
        clear_entry()
        normal_plot()


def clear_entry():
    entry_material.delete(0, 'end')


def update_bar_pie_plots():
    bar_pie_plots()

def dodaj_do_config():
    dodaj_root = customtkinter.CTk()
    dodaj_root.title("Wykresy Plots")
    dodaj_root.geometry("400x450")

    textbox = customtkinter.CTkTextbox(dodaj_root, width=400, height=450, corner_radius=0, fg_color=("#2b2b2b"))
    textbox.place(relx=0.01, rely=0.01)

    for section in config_obj.sections():
        for range, item in config_obj.items(section):
            if section == 'material':
                textbox.insert("0.1", f"{range} - {item}\n")

    def clear_entry_dodaj():
        dodaj_nowy.delete(0, 'end')
        usun_stary.delete(0, 'end')

    def add():
        materialy_lista = []
        for section in config_obj.sections():
            for range, item in config_obj.items(section):
                if section == 'material':
                    materialy_lista.append(item)
        new_materaial = dodaj_nowy.get()
        range_number = str(len(materialy_lista) + 1)
        config_obj.set('material', range_number, str(new_materaial))
        with open(r"wykresy_plot_config.ini", 'w') as configfile:
            config_obj.write(configfile)
        textbox.insert("0.1", f"{range_number} - {new_materaial}\n")
        material_box.configure(values=materialy_lista)
        clear_entry_dodaj()

    def delete():
        remove_material = usun_stary.get()
        config_obj.remove_option('material', remove_material)
        with open(r"wykresy_plot_config.ini", 'w') as configfile:
            config_obj.write(configfile)
        textbox.delete("0.0", "end")
        materialy_lista = []
        for section in config_obj.sections():
            for range, item in config_obj.items(section):
                if section == 'material':
                    materialy_lista.append(item)
                    textbox.insert("0.1", f"{range} - {item}\n")
        material_box.configure(values=materialy_lista)
        clear_entry_dodaj()

    dodaj_nowy_label = customtkinter.CTkLabel(master=dodaj_root, text="Dodaj nową pozycję", fg_color=("#2b2b2b"))
    dodaj_nowy_label.place(relx=0.45, rely=0.01)
    dodaj_nowy = customtkinter.CTkEntry(master=dodaj_root, width=120, height=28, border_width=1, justify="center")
    dodaj_nowy.place(relx=0.45, rely=0.06)
    dodaj_nowy_button = customtkinter.CTkButton(dodaj_root, text='Dodaj', width=55, height=30, border_width=1, command=add)
    dodaj_nowy_button.place(relx=0.53, rely=0.13)

    usun_stary_label = customtkinter.CTkLabel(master=dodaj_root, text="Usuń pozycję nr", fg_color=("#2b2b2b"))
    usun_stary_label.place(relx=0.5, rely=0.21)
    usun_stary = customtkinter.CTkEntry(master=dodaj_root, width=120, height=28, border_width=1, justify="center")
    usun_stary.place(relx=0.45, rely=0.26)
    usun_stary_button = customtkinter.CTkButton(dodaj_root, text='Usuń', width=55, height=30, border_width=1, command=delete)
    usun_stary_button.place(relx=0.53, rely=0.33)

    dodaj_root.mainloop()

def bar_pie_plots():
    global df_final
    df_plots = df_final.copy()
    df_plots = df_plots.loc[(df_plots['Data'] >= f'{y1}-{m1}-01') & (df_plots['Data'] <= f'{y2}-{m2}-31')]

    # Nowy df dla bar plot data i wartosc
    df_plots_wartosc = df_plots[["Data", "Wartosc"]]
    df_plots_wartosc = df_plots_wartosc[df_plots_wartosc['Wartosc'].notna()]
    df_plots_wartosc = df_plots_wartosc.sort_values(by='Data')
    df_plots_wartosc = df_plots_wartosc.groupby(df_plots_wartosc['Data'].dt.month)['Wartosc'].sum().round(0)

    # Nowy df dla pie plot data i wartosc
    df_nr_praca = df_plots[["Data", "Nr Suwnicy", "Wartosc pracy"]]
    df_nr_praca = df_nr_praca[df_nr_praca['Wartosc pracy'].notna()]
    df_nr_praca = df_nr_praca.groupby(df_nr_praca['Nr Suwnicy'])['Wartosc pracy'].sum()
    df_nr_praca = df_nr_praca.sort_values(ascending=False)

    # Dane do texbox
    df_textbox = df_plots[["Data", "Opis prac i wykaz materiału", "Ilosc", "Wartosc"]]
    df_textbox = df_textbox[df_textbox['Wartosc'].notna()]
    pattern = r'[-][ ]([0-9]+.{1,}|[0-9])+ (?:szt|kpl|op|kg|mb|l|m|m3|km)'
    df_textbox['Opis prac i wykaz materiału'] = df_textbox['Opis prac i wykaz materiału'].replace(regex=pattern, value='').str.join("")
    df_textbox['Opis prac i wykaz materiału'] = df_textbox['Opis prac i wykaz materiału'].replace(regex=" - ", value='').str.join("")
    df_textbox['Opis prac i wykaz materiału'] = df_textbox['Opis prac i wykaz materiału'].replace(regex="- ", value='').str.join("")
    df_textbox['Material lower'] = df_textbox['Opis prac i wykaz materiału'].str.lower()

    df_textbox['Opis prac i wykaz materiału'] = df_textbox['Opis prac i wykaz materiału'].str.strip()
    df_textbox['Opis prac i wykaz materiału'] = df_textbox['Opis prac i wykaz materiału'].str.title()

    df_textbox['Material lower'] = df_textbox['Opis prac i wykaz materiału'].replace(regex=" ", value='').str.join("")
    df_textbox['Material lower'] = df_textbox['Material lower'].str.lower()

    nazwa_dict = {}
    for i, j in zip(df_textbox['Opis prac i wykaz materiału'], df_textbox['Material lower']):
        nazwa_dict[i] = j

    df_textbox2 = df_textbox.groupby("Material lower")['Wartosc'].agg(['sum', 'count']).round(decimals=0)
    df_textbox2 = df_textbox2.sort_values(by=['sum'], ascending=True)
    df_textbox2['Material'] = df_textbox2.index
    df_textbox2['sum'] = df_textbox2['sum'].astype('Int64')

    # Znalezienie blednej nazwy i zamiana na prawidlowa
    true_name = []
    key_list = list(nazwa_dict.keys())
    value_list = list(nazwa_dict.values())

    for i in df_textbox2['Material']:
        if i in nazwa_dict.values():
            a = value_list.index(i)
            true_name.append(key_list[a])

    df_textbox2['Material'] = true_name

    def tk_textbox(*args):
        materialy_text_box.yview(*args)
        materialy_text_box2.yview(*args)
        materialy_text_box3.yview(*args)



    materialy_text_box = customtkinter.CTkTextbox(root, width=350, height=400, corner_radius=0, fg_color=("#2b2b2b"))
    materialy_text_box.place(relx=0.39, rely=0.01)
    materialy_text_box2 = customtkinter.CTkTextbox(root, width=60, height=400, corner_radius=0, fg_color=("#2b2b2b"))
    materialy_text_box2.place(relx=0.564, rely=0.01)
    materialy_text_box3 = customtkinter.CTkTextbox(root, width=100, height=400, corner_radius=0, fg_color=("#2b2b2b"))
    materialy_text_box3.place(relx=0.58, rely=0.01)
    ctk_textbox_scrollbar = customtkinter.CTkScrollbar(root, height=400, command=tk_textbox)
    ctk_textbox_scrollbar.place(relx=0.625, rely=0.01)

    for material, ilosc, wartosc in zip(df_textbox2['Material'], df_textbox2['count'], df_textbox2['sum']):
        materialy_text_box.insert("0.1", f"{material}\n\n")
        materialy_text_box2.insert("0.1", f"{ilosc}\n\n")
        materialy_text_box3.insert("0.1", f"{wartosc}zł\n\n")

    materialy_text_box.insert("0.0", "Nazwa materiału\n")
    materialy_text_box2.insert("0.0", "Ilość\n")
    materialy_text_box3.insert("0.0", "Suma [zł]\n")

    materialy_text_box.configure(yscrollcommand=ctk_textbox_scrollbar.set, state="disabled")
    materialy_text_box2.configure(yscrollcommand=ctk_textbox_scrollbar.set, state="disabled")
    materialy_text_box3.configure(yscrollcommand=ctk_textbox_scrollbar.set, state="disabled")

    def bar_plot(xbar, ybar):
        fig, ax = plt.subplots(figsize=(5.7, 3.8))
        ax.set_axisbelow(True)
        ax.grid(axis='y', linestyle='--', linewidth=0.5, zorder=0)

        ax.set_facecolor((0.16862, 0.16862, 0.16862))
        fig.patch.set_facecolor((0.16862, 0.16862, 0.16862))
        ax.tick_params(axis='x', colors='white')
        ax.tick_params(axis='y', colors='white')

        plt.bar(xbar, ybar, color='lightblue')

        font = {'family': 'serif', 'color': 'white', 'size': 12}
        plt.xlabel("Miesiąc", fontdict=font, color='white')
        plt.ylabel("Suma kosztu materiałów [zł]", fontdict=font, color='white')
        plt.title(f"Koszt materiałów w {y1}-{y2}", color='white')

        plt.xticks(xbar)
        fig.tight_layout()

        ax.bar_label(ax.containers[0], label_type='edge', color='white', rotation=0, fontsize=7, padding=3)
        ax.margins(y=0.1)
        plt.savefig("wykres_slupkowy.png")

        canvas = FigureCanvasTkAgg(fig, master=root)
        canvas.draw()
        canvas.get_tk_widget().place(relx=0.01, rely=0.03)

    def pie_plot(index_bar, wartosci_bar):
        global y
        df_pie_plot = pd.DataFrame(data={'Suwnica': index_bar, 'Wartosc pracy': wartosci_bar})

        wartosc_main = df_pie_plot[:12].copy()

        new_row = pd.DataFrame(data={'Suwnica': ['RESZTA'], 'Wartosc pracy': [df_pie_plot['Wartosc pracy'][12:].sum()]})

        wartosc_main = pd.concat([wartosc_main, new_row])

        fig, ax = plt.subplots(figsize=(5, 5))

        explode = [0.06, 0.06, 0.06, 0.06, 0.06, 0.06, 0.06, 0.06, 0.06, 0.06, 0.06, 0.06, 0.06]

        patches, texts, pcts = ax.pie(wartosc_main['Wartosc pracy'], labels=wartosc_main['Suwnica'], autopct=lambda x: f'{x:.1f}%\n({(x/100)*sum(df_nr_praca):.0f} zł)', startangle=90, textprops={'fontsize':9},
            colors=sns.color_palette('Set2'), labeldistance=0.24, explode=explode, pctdistance=1.15, rotatelabels=True)

        for i, patch in enumerate(patches):
            texts[i].set_color(patch.get_facecolor())
        plt.setp(pcts, color='aliceblue')
        plt.setp(texts, color='aliceblue')

        ax.set_facecolor((0.16862, 0.16862, 0.16862))
        fig.patch.set_facecolor((0.16862, 0.16862, 0.16862))

        font = {'family': 'serif', 'color':(0.16862, 0.16862, 0.16862), 'size': 14}
        plt.title(f"Zysk za rbh w {y1}-{y2} - {df_nr_praca.sum().round(2)} zł",
                  color='white')
        hole = plt.Circle((0, 0), 0.05, facecolor=(0.16862, 0.16862, 0.16862))
        plt.gcf().gca().add_artist(hole)
        plt.savefig("wykres_kolowy.png")

        canvas2 = FigureCanvasTkAgg(fig, master=root)
        canvas2.draw()
        canvas2.get_tk_widget().place(relx=0.33, rely=0.41)

    pie_plot(df_nr_praca.index, df_nr_praca)
    bar_plot(df_plots_wartosc.index, df_plots_wartosc.values)

def normal_plot():
    print(material)
    # Nowy df oraz przedział czasowy dla normal plot
    df_material_cena = df_materialy[['Data', 'Opis prac i wykaz materiału', 'Cena']].copy()
    df_material_cena['Opis prac i wykaz materiału'] = df_material_cena['Opis prac i wykaz materiału'].replace(regex=[r'[-][ ]([0-9]+.{1,}|[0-9])+ (?:szt|kpl|op|kg|mb|l|m|m3)'], value='').str.join("")
    df_material_cena['Opis prac i wykaz materiału'] = df_material_cena['Opis prac i wykaz materiału'].replace(regex=[r'[ ][-][ ]'], value='').str.join("")
    df_material_cena['Opis prac i wykaz materiału'] = df_material_cena['Opis prac i wykaz materiału'].replace(regex=[r'[-][ ]'], value='').str.join("")
    df_material_cena = df_material_cena[df_material_cena['Cena'].notna()]
    df_material_cena['Opis prac i wykaz materiału'] = df_material_cena['Opis prac i wykaz materiału'].str.strip()

    df_material_cena = df_material_cena[df_material_cena['Opis prac i wykaz materiału'].str.contains(f"(?i){material}", na=False)]
    df_material_cena = df_material_cena.sort_values(by='Data')
    df_material_cena['Data'] = df_material_cena['Data'].dt.date
    # df_material_cena = df_material_cena.drop_duplicates()

    # Usuwanie błędnych danych, sortowanie unikalnych cen z datami ich wystąpoenia
    final_cena = []
    final_data = []

    final_cena.append(df_material_cena['Cena'].iloc[4])
    final_data.append(df_material_cena['Data'].iloc[4])

    check_cena = (final_cena[0])
    count_cena = 0
    df_material_cena = df_material_cena[df_material_cena['Cena'] < df_material_cena['Cena'].mean() * 2]
    for data, cena in zip(df_material_cena['Data'], df_material_cena['Cena']):
        if cena == check_cena:
            count_cena += 1
            if count_cena > 1:
                if final_cena[-1] == cena:
                    continue
                else:
                    final_cena.append(check_cena)
                    final_data.append(data)
            else:
                continue
        else:
            count_cena = 0
            check_cena = cena

    x = final_data
    y = final_cena
    fig, ax = plt.subplots(figsize=(6, 4))
    ax.set_axisbelow(True)
    ax.grid(axis='both', linestyle='--', linewidth=0.5, zorder=0)

    ax.set_facecolor((0.16862, 0.16862, 0.16862))
    fig.patch.set_facecolor((0.16862, 0.16862, 0.16862))

    font = {'family': 'serif', 'color': 'white', 'size': 12}
    plt.xlabel("Data", fontdict=font, color='white')
    plt.ylabel("Cena materiału [zł]", fontdict=font, color='white')
    plt.title(f"Odnotowane ceny materiału: {material}", color='white')

    ax.tick_params(colors='white', which='both')

    plt.plot(x, y, marker='o', color='slateblue')
    for x, y in zip(x, y):
        label = "{:.2f}".format(y)
        plt.annotate(label,(x, y), textcoords="offset points", xytext=(0, 10), ha='center', color='white')

    canvas3 = FigureCanvasTkAgg(fig, master=root)
    canvas3.draw()
    canvas3.get_tk_widget().place(relx=0.63, rely=0.01)
    plt.savefig("material_wykres.png")

normal_plot()
bar_pie_plots()


change_year_label = customtkinter.CTkLabel(root, text=("Wybierz Rok dla wykresów: \n koszt materiałów oraz zysk z rbh \n \n --"), fg_color=("#dbdbdb", "#2b2b2b"))
change_year_label.place(relx=0.15, rely=0.51)
box_rok_y1 = customtkinter.CTkComboBox(root, border_width=1, values=year_list, width=100, justify='center', command=reload_barplot_y1)
box_rok_y1.place(relx=0.144, rely=0.55)
box_rok_y2 = customtkinter.CTkComboBox(root, border_width=1, values=year_list, width=100, justify='center', command=reload_barplot_y2)
box_rok_y2.place(relx=0.205, rely=0.55)
zatwierdz_year = customtkinter.CTkButton(root, text='Zatwierdź/Przeładuj wykresy', width=50, height=32, border_width=1, command=update_bar_pie_plots)
zatwierdz_year.place(relx=0.155, rely=0.587)

change_material = customtkinter.CTkLabel(root, text=("Wybierz lub wpisz materiał do wykresu:"), fg_color=("#dbdbdb", "#2b2b2b"))
change_material.place(relx=0.79, rely=0.54)
material_box = customtkinter.CTkComboBox(root, border_width=1, values=materialy_lista, width=150, justify='center', command=reload_normalplot)
material_box.place(relx=0.77, rely=0.57)
entry_material = customtkinter.CTkEntry(master=root, width=120, height=28, border_width=1, justify="center")
entry_material.place(relx=0.85, rely=0.57)
zatwierdz_material = customtkinter.CTkButton(root, text='Zatwierdź/Przeładuj wykres', width=55, height=30, border_width=1, command=update_normalplot)
zatwierdz_material.place(relx=0.797, rely=0.606)

dodaj_nowe_label = customtkinter.CTkLabel(root, text=("Dodaj nowe dane do bazy danych:"), fg_color=("#dbdbdb", "#2b2b2b"))
dodaj_nowe_label.place(relx=0.78, rely=0.65)
dodaj_button = customtkinter.CTkButton(root, text='Dodaj', width=55, height=30, border_width=1, command=dodaj_do_config)
dodaj_button.place(relx=0.885, rely=0.65)

root.mainloop()

