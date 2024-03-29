import pandas as pd
import subprocess
from tkinter import filedialog
import os
from CTkMessagebox import CTkMessagebox
import customtkinter
import threading


def main():
    loading_info = ()
    loading_info = 'Ładowanie... Może potrwać ponad 30s'

    def ladowanie_info():
        CTkMessagebox(title="Ładowanie", message=loading_info, option_1="Ok")

    path = filedialog.askopenfilename(initialdir="/", title="Dodaj aktualny obmiar",
                                      filetypes=(("pliki excel", "*.xlsx"), ("wszystkie pliki", "*")))
    spreed_sheet = pd.ExcelFile(path)
    ws = spreed_sheet.sheet_names

    threading.Thread(target=ladowanie_info).start()

    df = pd.read_excel(path, sheet_name=None, header=1)

    df_all = pd.concat(df.values(), ignore_index=True)

    # Plik w sumie ma około 100k wierszy i 120 arkuszy, pętla poglądowo wyświetla przetworzone arkusze
    for i in ws:
        df2 = spreed_sheet.parse(i)
        dimensions = df2.shape
        print(f"Ładowanie nowych danych--->{i}--->{dimensions[0]} wierszy")

    print("-"*60)
    loading_info = 'Grupowanie i filtrowanie danych'
    print('Grupowanie i filtrowanie danych')

    # Kopia do wykresów
    df_plots = df_all.copy()

    # Grupowanie i filtrowanie danych
    df_all['Nr Suwnicy'] = df_all['Nr Suwnicy'].fillna(method='ffill')
    df_all['Data'] = df_all['Data'].fillna(method='ffill')
    df_all = df_all[df_all['Cena'].notna()]

    df_data = df_all[["Nr Suwnicy", "Data", "Opis prac i wykaz materiału", "Cena"]]

    # Wyciągnięcie ilości materiałów z tekstu do nowej kolumny
    pattern = r'[-][ ]([0-9]+.{1,}|[0-9])+ (?:szt|kpl|op|kg|mb|l|m|m3|km)'

    df_data = df_data[["Nr Suwnicy", "Data", "Opis prac i wykaz materiału", "Opis prac i wykaz materiału", "Cena"]]
    df_data.columns = ["Nr Suwnicy", "Data", "Material", "Ilosc", "Cena"]

    df_data['Ilosc'] = df_data['Ilosc'].str.findall(pattern).str.join(", ")

    df_data = df_data.dropna(how='any')
    df_data['Ilosc'] = df_data['Ilosc'].replace(regex=[r','], value='.')

    df_data['Data'] = df_data['Data'].replace(regex=[r'r.'], value='')
    df_data['Data'] = df_data['Data'].str.findall(r"^[0-9]{2}-[0-9]{2}-[0-9]{4}").str.join("")
    df_data['Data'] = df_data['Data'].replace(r'^s*$', float('NaN'), regex=True)
    df_data['Data'] = df_data['Data'].fillna(method='ffill')
    df_data['Material'] = df_data['Material'].str.strip()

    df_data['Ilosc'] = pd.to_numeric(df_data['Ilosc'], errors='coerce')
    df_data['Cena'] = pd.to_numeric(df_data['Cena'], errors='coerce')

    df_data['Wartosc'] = df_data['Ilosc'] * df_data['Cena']
    df_data['Wartosc'] = df_data['Wartosc'].round(decimals=2)

    df_data['Data'] = pd.to_datetime(df_data['Data'], format='%d-%m-%Y', errors='coerce', utc=False)
    df_data = df_data[df_data['Cena'].notna()]
    df_data['Cena'] = df_data['Cena'].round(decimals=3)

    df_data['Material'] = df_data['Material'].replace(regex=[r'[-][ ]([0-9]+.{1,}|[0-9])+ (?:szt|kpl|op|kg|mb|l|m|m3)'], value='').str.join("")

    if os.path.exists("dfall.xlsx"):
        os.remove("dfall.xlsx")

    # Utworzenie bazy danych
    df_data.to_excel("dfall.xlsx", index=False)
    # Ukrywa plik w systemie windows
    subprocess.check_call(["attrib", "+H", "dfall.xlsx"])
    print("-"*60)
    loading_info = 'Ładowanie zakończone'
    print("Ładowanie obmiaru zakończone")

    df_plots = df_plots[["Nr rejestru", "Data", "Opis prac i wykaz materiału", "Cena", "Nr Suwnicy"]]
    df_plots = df_plots[["Nr rejestru", "Data", "Opis prac i wykaz materiału", "Opis prac i wykaz materiału", "Cena", "Nr Suwnicy", "Opis prac i wykaz materiału"]]
    df_plots.columns = ["Nr rejestru", "Data", "Opis prac i wykaz materiału", "Ilosc", "Cena", "Nr Suwnicy", "Rbh"]
    df_plots['Data'] = df_plots['Data'].replace(regex=[r'r.'], value='')
    df_plots['Data'] = df_plots['Data'].replace(regex=[r"^[0-9]{2}-[0-9]{2}-[0-9]{4} - "], value='')
    df_plots['Data'] = df_plots['Data'].str.findall(r"^[0-9]{2}-[0-9]{2}-[0-9]{4}").str.join("")
    df_plots['Data'] = df_plots['Data'].replace(r'^s*$', float('NaN'), regex=True)
    df_plots['Data'] = df_plots['Data'].fillna(method='ffill')
    df_plots['Data'] = pd.to_datetime(df_plots['Data'], format='%d-%m-%Y', errors='coerce', utc=False)

    df_plots['Nr Suwnicy'] = df_plots['Nr Suwnicy'].fillna(method='ffill')

    # Wyciągnięcie ilości materiałów z tekstu do nowej kolumny
    pattern = r'[-][ ]([0-9]+.{1,}|[0-9])+ (?:szt|kpl|op|kg|mb|l|m|m3|km)'
    pattern_2 = r'[-][ ]([0-9]+.[0-9]|[0-9]+)+ (?:rbh)'
    pattern_3 = r'[a-zA-Z]+'

    df_plots['Ilosc'] = df_plots['Ilosc'].str.findall(pattern).str.join(", ")
    df_plots['Ilosc'] = df_plots['Ilosc'].replace(regex=[r','], value='.')

    df_plots['Ilosc'] = pd.to_numeric(df_plots['Ilosc'], errors='coerce')
    df_plots['Cena'] = pd.to_numeric(df_plots['Cena'], errors='coerce')

    df_plots['Wartosc'] = df_plots['Ilosc'] * df_plots['Cena']
    df_plots['Wartosc'] = df_plots['Wartosc'].round(decimals=2)

    df_plots['Rbh'] = df_plots['Rbh'].str.findall(pattern_2).str.join(", ")
    df_plots['Rbh'] = df_plots['Rbh'].replace(regex=[r','], value='.')
    df_plots['Rbh'] = pd.to_numeric(df_plots['Rbh'], errors='coerce')

    df_plots['Nr Suwnicy'] = df_plots['Nr Suwnicy'].replace(regex=[pattern_3], value='')
    df_plots['Nr Suwnicy'] = df_plots['Nr Suwnicy'].replace(regex=[r'-'], value='')
    df_plots['Nr Suwnicy'] = df_plots['Nr Suwnicy'].replace(regex=[r' '], value='')

    df_plots = df_plots[df_plots['Opis prac i wykaz materiału'].str.contains("kz") == False]
    df_plots = df_plots[df_plots['Opis prac i wykaz materiału'].str.contains("Kz") == False]
    df_plots = df_plots[df_plots['Opis prac i wykaz materiału'].str.contains("faktura") == False]
    df_plots = df_plots[df_plots['Opis prac i wykaz materiału'].str.contains("Faktura") == False]

    if os.path.exists("dfplot.xlsx"):
        os.remove("dfplot.xlsx")
    df_plots.to_excel("dfplot.xlsx", index=False)
    subprocess.check_call(["attrib", "+H", "dfplot.xlsx"])
    print('Baza do wykresów została załadowana')
    threading.Thread(target=ladowanie_info).start()


def extract_xlsx():
    customtkinter.set_appearance_mode("dark")
    customtkinter.set_default_color_theme("blue")
    app = customtkinter.CTk()
    app.withdraw()

    threading.Thread(target=main).start()

    app.mainloop()
