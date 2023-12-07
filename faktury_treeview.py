import customtkinter
from tkinter import ttk
import pandas as pd
import re


def faktury_treeview():
    faktury_window = customtkinter.CTkToplevel()
    faktury_window.title("Faktury Zestawienie materiałów")
    faktury_window.geometry("1455x613")
    faktury_window.attributes('-topmost', 1)
    faktury_window.resizable(0, 0)

    left_frame_faktury = customtkinter.CTkFrame(faktury_window, corner_radius=22, width=215, height=200, border_width=1)
    left_frame_faktury.grid(row=0, column=0, columnspan=1, padx=1, pady=1, sticky="ew")

    right_frame_faktury = customtkinter.CTkFrame(faktury_window, corner_radius=22, width=1170, height=610, border_width=1)
    right_frame_faktury.grid(row=0, column=3, padx=1, pady=1, sticky="nw")

    faktury_tree = ttk.Treeview(right_frame_faktury, height=28, show='headings', columns=('Firma', 'Data', 'Nazwa Materiału', 'Ilość', 'Cena [zł]', 'Nr faktury'))
    treescroll_faktury = customtkinter.CTkScrollbar(right_frame_faktury, command=faktury_tree.yview, fg_color=("#dbdbdb", "#2b2b2b"))
    faktury_tree.configure(yscrollcommand=treescroll_faktury.set, selectmode="extended")
    treescroll_faktury.grid(row=0, column=2, sticky="ne", pady=15, ipady=185)

    def search_faktury():
        global search_word
        search_word = search_entry_faktury.get()
        temp = list(search_word)
        temp.clear()
        clear_entry_faktury()
        tree()

    def clear_entry_faktury():
        search_entry_faktury.delete(0, 'end')

    def clear_faktury_tree():
        for i in faktury_tree.get_children():
            faktury_tree.delete(i)

    def drukowanie_faktury():
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        subprocess.call('cmd /c "Faktury materiały.xlsx"', startupinfo=si)

    df = pd.read_excel(r'Faktury materiały.xlsx')
    dff = df['Firma'].unique()
    list_firmy = []
    list_firmy.insert(0, 'Wszystkie')
    firma_nr = 'Wszystkie'

    def select_firma(box_firma):
        global firma_nr
        if box_firma in list_firmy:
            firma_nr = box_firma

    try:
        for i in dff:
            list_firmy.append(str(i))

    except NameError:
        print("Brak bazy faktur.")

    def dodaj_pozycje():
        dodaj_window = customtkinter.CTkToplevel(faktury_window)
        dodaj_window.title("Dodaj/Usuń pozycję")
        dodaj_window.geometry("640x250")
        dodaj_window.transient(faktury_window)
        dodaj_window.resizable(0, 0)

        dodaj_pozycje_frame = customtkinter.CTkFrame(dodaj_window, corner_radius=22,
                                                    width=100,
                                                    height=200,
                                                    border_width=1)
        dodaj_pozycje_frame.grid(row=0, column=0, columnspan=1, padx=1, pady=1, sticky="nw")

        def dodaj():
            df = pd.read_excel(r'Faktury materiały.xlsx')

            data_value = data_entry.get()
            pattern = r'[0-9]{4}-[0-9]{2}-[0-9]{2}'
            match = re.findall(pattern, data_value)
            if match:
                firma_value = firma_entry.get()
                material_value = material_entry.get()
                ilosc_value = ilosc_entry.get()
                cena_value = cena_entry.get()
                faktura_value = faktura_entry.get()

                d = [firma_value, data_value, material_value, ilosc_value, cena_value, faktura_value]
                df.loc[len(df)] = d

                df.to_excel('Faktury materiały.xlsx', index=False)
                # Filtrowanie, sortowanie tak jak w przypadku dodawania danych z pdf/images
                df = pd.read_excel('Faktury materiały.xlsx')
                df['Data'] = df['Data'].dt.date

                df['Cena'] = df['Cena'].replace(regex=[r','], value='.')
                df['Cena'] = df['Cena'].replace(regex=[r' '], value='')
                df['Cena'] = pd.to_numeric(df['Cena'], errors='coerce')
                df['Cena'] = df['Cena'].round(decimals=2)

                df['Ilość'] = df['Ilość'].astype(str)
                df['Ilość'] = df['Ilość'].replace(regex=[r','], value='.').str.join("")
                df['Ilość'] = df['Ilość'].replace(regex=[r' '], value='')
                df['Ilość'] = pd.to_numeric(df['Ilość'], errors='coerce')
                df['Ilość'] = df['Ilość'].round(decimals=2)

                df['Nazwa Materiału'] = df['Nazwa Materiału'].str.strip()
                df['Nr faktury'] = df['Nr faktury'].str.strip()
                df = df.sort_values(by='Data')

                df = df.drop_duplicates()
                writer = pd.ExcelWriter('Faktury materiały.xlsx')
                df.to_excel(writer, sheet_name='Sheet1', index=False, na_rep='NaN')

                writer.sheets['Sheet1'].set_column(0, 0, 7)
                writer.sheets['Sheet1'].set_column(1, 1, 10)
                writer.sheets['Sheet1'].set_column(2, 2, 60)
                writer.sheets['Sheet1'].set_column(3, 3, 5)
                writer.sheets['Sheet1'].set_column(4, 4, 7)
                writer.sheets['Sheet1'].set_column(5, 5, 14)

                writer._save()

                clear_entrys()
                tree()
                date_error = customtkinter.CTkLabel(dodaj_pozycje_frame, text=" "*57, fg_color=("#dbdbdb", "#2b2b2b"))
                date_error.grid(row=0, column=0, columnspan=1, padx=60, pady=62, sticky="nw")
            else:
                date_error = customtkinter.CTkLabel(dodaj_pozycje_frame, text="Wprowadź odpowieni format\n rok-miesiąc-dzień", text_color=("FF0000", "#FF0000"))
                date_error.grid(row=0, column=0, columnspan=1, padx=60, pady=62, sticky="nw")

        def usun():
            df = pd.read_excel(r'Faktury materiały.xlsx')

            delete_row_data = data_entry_usun.get()
            df = df.loc[df['Data'] != delete_row_data]

            delete_row_faktura = faktura_entry_usun.get()
            df = df.loc[df['Nr faktury'] != delete_row_faktura]

            df.to_excel('Faktury materiały.xlsx', index=False)

            clear_entrys()
            tree()

        def clear_entrys():
            firma_entry.delete(0, 'end')
            data_entry.delete(0, 'end')
            material_entry.delete(0, 'end')
            ilosc_entry.delete(0, 'end')
            cena_entry.delete(0, 'end')
            faktura_entry.delete(0, 'end')
            data_entry_usun.delete(0, 'end')
            faktura_entry_usun.delete(0, 'end')

        firma_label = customtkinter.CTkLabel(dodaj_pozycje_frame, text=("Firma" + "    " * 5 + "Data" + "    " * 8 +
                                                                        "Materiał" + "    " * 7 + "Ilość" + "    " * 4 + "Cena" +
                                                                        "    " * 5 + "Nr Faktury"), fg_color=("#dbdbdb", "#2b2b2b"))
        firma_label.grid(row=0, column=0, columnspan=1, padx=22, pady=1, sticky="nw")

        firma_entry = customtkinter.CTkEntry(master=dodaj_pozycje_frame, width=70, height=30, border_width=1, justify='center')
        firma_entry.grid(row=0, column=0, columnspan=1, padx=5, pady=28, sticky="nw")

        data_entry = customtkinter.CTkEntry(master=dodaj_pozycje_frame, width=90, height=30, border_width=1, justify='center')
        data_entry.grid(row=0, column=0, columnspan=1, padx=80, pady=28, sticky="nw")

        material_entry = customtkinter.CTkEntry(master=dodaj_pozycje_frame, width=170, height=30, border_width=1, justify='center')
        material_entry.grid(row=0, column=0, columnspan=1, padx=175, pady=28, sticky="nw")

        ilosc_entry = customtkinter.CTkEntry(master=dodaj_pozycje_frame, width=70, height=30, border_width=1, justify='center')
        ilosc_entry.grid(row=0, column=0, columnspan=1, padx=350, pady=28, sticky="nw")

        cena_entry = customtkinter.CTkEntry(master=dodaj_pozycje_frame, width=70, height=30, border_width=1, justify='center')
        cena_entry.grid(row=0, column=0, columnspan=1, padx=425, pady=28, sticky="nw")

        faktura_entry = customtkinter.CTkEntry(master=dodaj_pozycje_frame, width=130, height=30, border_width=1, justify='center')
        faktura_entry.grid(row=0, column=0, columnspan=1, padx=500, pady=28, sticky="nw")

        dodaj_button = customtkinter.CTkButton(dodaj_pozycje_frame, text='Dodaj nową pozyję do listy', width=80, height=40, border_width=1, command=dodaj)
        dodaj_button.grid(row=0, padx=235, pady=70, column=0, sticky="nw")

        usun_label = customtkinter.CTkLabel(dodaj_pozycje_frame, text=("Data" + "    "*7 + "Nr Faktury"), fg_color=("#dbdbdb", "#2b2b2b"))
        usun_label.grid(row=0, column=0, columnspan=1, padx=230, pady=125, sticky="nw")

        data_entry_usun = customtkinter.CTkEntry(master=dodaj_pozycje_frame, width=90, height=30, border_width=1, justify='center')
        data_entry_usun.grid(row=0, column=0, columnspan=1, padx=200, pady=150, sticky="nw")

        faktura_entry_usun = customtkinter.CTkEntry(master=dodaj_pozycje_frame, width=140, height=30, border_width=1, justify='center')
        faktura_entry_usun.grid(row=0, column=0, columnspan=1, padx=300, pady=150, sticky="nw")

        usun_button = customtkinter.CTkButton(dodaj_pozycje_frame, text='Usuń pozyję wpisując date lub nr faktury', width=80, height=40, border_width=1, command=usun)
        usun_button.grid(row=0, padx=200, pady=187, column=0, sticky="nw")

    def tree():
        global firma_nr
        filename = r'Faktury materiały.xlsx'
        if filename:
            try:
                filename = r"{}".format(filename)
                df = pd.read_excel(filename)
                dff = df['Firma'].unique()

            except ValueError:
                print("Nie wykrywam pliku z fakturami, załaduj ponownie faktury.")

            except FileNotFoundError:
                print("Nie wykrywam pliku z fakturami, załaduj ponownie faktury.")

        clear_faktury_tree()

        faktury_tree.column("Firma", anchor='center', stretch=False, width=130)
        faktury_tree.heading("Firma", text="Firma")
        faktury_tree.column("Data", anchor='center', stretch=False, width=130)
        faktury_tree.heading("Data", text="Data")
        faktury_tree.column("Nazwa Materiału", anchor='w', stretch=False, width=470)
        faktury_tree.heading("Nazwa Materiału", text="Materiał")
        faktury_tree.column("Ilość", anchor='center', stretch=False, width=110)
        faktury_tree.heading("Ilość", text="Ilość")
        faktury_tree.column("Cena [zł]", anchor='e', stretch=False, width=110)
        faktury_tree.heading("Cena [zł]", text="Cena [zł]")
        faktury_tree.column("Nr faktury", anchor='e', stretch=False, width=160)
        faktury_tree.heading("Nr faktury", text="Nr faktury")

        columns = ['Firma', 'Data', 'Nazwa Materiału', 'Ilość', 'Cena [zł]', 'Nr faktury']

        if search_word:
            if df['Nazwa Materiału'].str.contains(f"(?i){search_word}").any():
                df = df[df['Nazwa Materiału'].str.contains(f"(?i){search_word}")]
            else:
                df = df[df['Nr faktury'].str.contains(f"(?i){search_word}")]
        else:
            pass

        # Sortowanie ze względu na firmę
        try:
            if firma_nr == 'Wszystkie':
                pass
            else:
                df = df.loc[df['Firma'] == firma_nr]
        except NameError:
            firma_nr = 'Wszystkie'

        df['Data'] = df['Data'].dt.date

        df_rows = df.to_numpy().tolist()
        for row in df_rows:
            faktury_tree.insert("", "end", values=row)

        def faktury_treeview_sort_column(tv, col, reverse):
            leng = [(tv.set(k, col), k) for k in tv.get_children('')]
            leng.sort(reverse=reverse)

            for index, (val, k) in enumerate(leng):
                tv.move(k, '', index)

            tv.heading(col, text=col, command=lambda _col=col:
            faktury_treeview_sort_column(tv, _col, not reverse))

        for col in columns:
            faktury_tree.heading(col, text=col, command=lambda _col=col:
            faktury_treeview_sort_column(faktury_tree, _col, False))

        faktury_tree.grid(row=0, column=2, padx=11, pady=11, sticky="nw")

    search_entry_faktury = customtkinter.CTkEntry(master=faktury_window, width=157, height=40, border_width=1, justify='center')
    search_entry_faktury.grid(row=0, column=0, columnspan=1, padx=5, pady=225, sticky="w")

    search_button_faktury = customtkinter.CTkButton(faktury_window, text='Szukaj/Odśwież', width=140, height=40, border_width=1, command=search_faktury)
    search_button_faktury.grid(row=0, column=0, padx=15, pady=225, sticky="e")
    search_entry_faktury.bind("<Return>", search_faktury())

    box_firma = customtkinter.CTkComboBox(faktury_window, values=list_firmy, border_width=1, width=110, justify='center', command=select_firma)
    box_firma.grid(row=0, column=0, columnspan=1, padx=65, pady=235, sticky="ne")

    firma_label = customtkinter.CTkLabel(faktury_window, text=f"Wyświetl firmę:", fg_color=("#dbdbdb", "#2b2b2b"))
    firma_label.grid(row=0, column=0, columnspan=1, padx=45, pady=235, sticky="nw")

    dodaj_usun = customtkinter.CTkButton(faktury_window, text='Dodaj/usuń pozycję', width=80, height=40, border_width=1, command=dodaj_pozycje)
    dodaj_usun.grid(row=0, padx=95, pady=225, column=0, sticky="se")

    faktury_excel = customtkinter.CTkButton(faktury_window, text='Drukuj', width=50, height=30, border_width=1, command=drukowanie_faktury)
    faktury_excel.grid(row=0, column=0, padx=15, pady=230, sticky="se")

    faktury_window.mainloop()

