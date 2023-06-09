import pandas as pd
import subprocess
from tkinter import filedialog
import os


def extract_xlsx():
    path = filedialog.askopenfilename(initialdir="/", title="Dodaj aktualny obmiar",
                                      filetypes=(("pliki excel", "*.xlsx"), ("wszystkie pliki", "*")))
    spreed_sheet = pd.ExcelFile(path)
    ws = spreed_sheet.sheet_names

    df = pd.read_excel(path, sheet_name=None, header=1)

    df_all = pd.concat(df.values(), ignore_index=True)

    # Plik w sumie ma około 100k wierszy i 120 arkuszy, pętla poglądowo wyświetla przetworzone arkusze
    for i in ws:
        df2 = spreed_sheet.parse(i)
        dimensions = df2.shape
        print(f"Ładowanie nowych danych--->{i}--->{dimensions[0]} wierszy")

    print("-"*60)
    print('Grupowanie i filtrowanie danych')

    # Grupowanie i filtrowanie danych
    df_all['Nr Suwnicy'] = df_all['Nr Suwnicy'].fillna(method='ffill')
    df_all['Data'] = df_all['Data'].fillna(method='ffill')
    df_all = df_all[df_all['Cena'].notna()]

    df_data = df_all[["Nr Suwnicy", "Data", "Opis prac i wykaz materiału", "Cena"]]

    # Wyciągnięcie ilości materiałów z tekstu do nowej kolumny
    pattern = r'[-][ ]([0-9]+.{1,}|[0-9])+ (?:szt|kpl|op|kg|mb|l|m|m3)'

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

    if os.path.exists("dfall.xlsx"):
        os.remove("dfall.xlsx")

    # Utworzenie bazy danych
    df_data.to_excel("dfall.xlsx", index=False)
    # Ukrywa plik w systemie windows
    subprocess.check_call(["attrib", "+H", "dfall.xlsx"])
    print("-"*60)
    print("Ładowanie zakończone")