import re
import camelot
import pandas as pd
import os, glob
from PyPDF2 import PdfReader
import openpyxl
import subprocess
from PIL import Image
from pytesseract import pytesseract
from tkinter import filedialog
import datetime


def faktury_database():
    subprocess.check_call(["attrib", "-H", "outpdf.xlsx"])
    path = filedialog.askdirectory(initialdir="/", title="Wybierz folder z fakturami",)

    miesiac = {"stycznia": "01", "lutego": "02", "marca": "03", "kwietnia": "04", "maja": "05", "czerwca": "06",
               "lipca": "07", "sierpnia": "08", "września": "09", "października": "10", "listopada": "11", "grudnia": "12"}

    def append_df_to_excel():
        if os.path.exists("Faktury materiały.xlsx"):
            print('Plik excel z fakturami istnieje, następuje dodanie nowych wartości')
        else:
            wb = openpyxl.Workbook()
            wb.save('Faktury materiały.xlsx')

        df_excel = pd.read_excel('Faktury materiały.xlsx')
        result = pd.concat([df_excel, df2], ignore_index=True)
        result.to_excel('Faktury materiały.xlsx', index=False)

    for file in glob.glob(os.path.join(path, '*.pdf')):
        tables = camelot.read_pdf(file)
        column_names = tables[0].df.columns.values.tolist()
        print(len(column_names))

        tables[0].df.to_excel('outpdf.xlsx', index=False)
        df1 = pd.read_excel('outpdf.xlsx')

        if len(column_names) == 10:
            df1.columns = ['LP', 'Nazwa Materiału', 'Data', 'Ilość', 'Jednostka', 'Cena', 'Wartość', 'Vat', 'Podatek',
                           'Suma']
            if df1['Data'].str.contains(f"\d+").any():
                df1 = df1[['Data', 'Nazwa Materiału', 'Ilość', 'Cena']]
                df2 = df1.dropna()
                df2 = df2.iloc[1:]
                df2 = df2.replace(regex=[r'\n'], value='')
                df2['Firma'] = 'Windex'
                df2['Nr faktury'] = f'{os.path.basename(file)}'
                df2['Nr faktury'] = df2['Nr faktury'].replace(regex=[r'.pdf'], value='')
                df2['Nr faktury'] = df2['Nr faktury'].replace(regex=[r'_'], value='/')
                df2['Nr faktury'] = df2['Nr faktury'].replace(regex=[r'FS/FS'], value='')
                df2['Nr faktury'] = df2['Nr faktury'].replace(regex=[r'FS'], value='')

                df2['Data'] = pd.to_datetime(df2['Data'], errors='coerce', utc=False)

                df2 = df2[['Firma', 'Data', 'Nazwa Materiału', 'Ilość', 'Cena', 'Nr faktury']]

                append_df_to_excel()
                print(df2)
            else:
                df2 = df1[['Data', 'Nazwa Materiału', 'Ilość', 'Cena']]
                df2 = df2.iloc[:-2]
                df2 = df2.iloc[2:]
                df2['Nazwa Materiału'] = df2['Nazwa Materiału'].replace(regex=[r'\n'], value='')
                df2['Firma'] = 'Sidus'

                sidusdate = ()
                faktury = open("faktury.txt", "w+", encoding="utf-8")
                reader = PdfReader(file)
                page = reader.pages[0]
                text = page.extract_text()
                try:
                    total = re.findall('ORYGINAŁ\n.+', text)
                    total[0] = total[0].replace('ORYGINAŁ\n', '')
                    for i, j in miesiac.items():
                        if i in total[0]:
                            total[0] = total[0].replace(i, f'-{j}-')
                    sidusdate = total[0]
                except IndexError:
                    total = re.findall('Odbiorca \n.+ Przelew', text)
                    total[0] = total[0].replace('Odbiorca \n', '')
                    total[0] = total[0].replace(' Przelew', '')
                    for i, j in miesiac.items():
                        if i in total[0]:
                            total[0] = total[0].replace(i, f'-{j}-')
                    sidusdate = total[0]
                print(text, file=faktury)

                df2['Data'] = sidusdate
                df2['Data'] = df2['Data'].replace(regex=[r' '], value='')
                df2['Data'] = pd.to_datetime(df2['Data'], format='%d-%m-%Y', errors='coerce', utc=False)
                df2['Nr faktury'] = f'{os.path.basename(file)}'
                df2['Nr faktury'] = df2['Nr faktury'].replace(regex=[r'.pdf'], value='')
                df2['Nr faktury'] = df2['Nr faktury'].replace(regex=[r'_'], value='/')
                df2['Nr faktury'] = df2['Nr faktury'].replace(regex=[r'Faktura VAT'], value='')
                df2['Nr faktury'] = df2['Nr faktury'].replace(regex=[r'FV'], value='')
                df2 = df2[['Firma', 'Data', 'Nazwa Materiału', 'Ilość', 'Cena', 'Nr faktury']]

                append_df_to_excel()
                print(df2)

        if len(column_names) == 9:
            df1.columns = ['LP', 'Nazwa Materiału', 'Data', 'Ilość', 'Jednostka', 'Cena', 'Wartość', 'Vat', 'Suma']
            df1 = df1[['Data', 'Nazwa Materiału', 'Ilość', 'Cena']]
            df2 = df1.dropna()
            df2 = df2.iloc[1:]
            df2 = df2.replace(regex=[r'\n'], value='')
            df2['Firma'] = 'Windex'
            df2['Nr faktury'] = f'{os.path.basename(file)}'
            df2['Nr faktury'] = df2['Nr faktury'].replace(regex=[r'.pdf'], value='')
            df2['Nr faktury'] = df2['Nr faktury'].replace(regex=[r'_'], value='/')
            df2['Nr faktury'] = df2['Nr faktury'].replace(regex=[r'FS/FS'], value='')
            df2['Nr faktury'] = df2['Nr faktury'].replace(regex=[r'FS'], value='')

            df2['Data'] = pd.to_datetime(df2['Data'], errors='coerce', utc=False)

            df2 = df2[['Firma', 'Data', 'Nazwa Materiału', 'Ilość', 'Cena', 'Nr faktury']]

            append_df_to_excel()
            print(df2)

        if len(column_names) == 6:
            df1.columns = [['Nazwa Materiału', 'Cena', 'Ilość', 'Jednostka', 'Cenaa', 'Wartość']]
        if len(column_names) == 5:
            df1.columns = [['Nazwa Materiału', 'Cena', 'Ilość', 'Jednostka', 'Cenaa']]
        if len(column_names) == 6 or len(column_names) == 5:
            df1 = df1.iloc[:-2]
            df1 = df1.iloc[2:]
            df2 = df1[['Nazwa Materiału', 'Nazwa Materiału', 'Cena']]
            df2.columns = ['Nazwa Materiału', 'Ilość', 'Cena']
            df2 = df2.replace(r'\n', ' ', regex=True)
            pattern = r'([0-9]{1,})+ (?:szt|kpl|op|kg|mb|l|m|m3)'
            df2['Ilość'] = df2['Ilość'].str.findall(pattern).str.join(", ")
            df2['Nazwa Materiału'] = df2['Nazwa Materiału'].replace(regex=[r'. ([0-9]{1,})+ (?:szt|kpl|op|kg|mb|l|m|m3).'], value='')
            df2['Nazwa Materiału'] = df2['Nazwa Materiału'].replace(regex=[r'[0-9] '], value='')

            df2['Firma'] = 'Sidus'

            sidusdate = ()
            faktury = open("faktury.txt", "w+", encoding="utf-8")
            reader = PdfReader(file)
            page = reader.pages[0]
            text = page.extract_text()
            try:
                total = re.findall('ORYGINAŁ\n.+', text)
                total[0] = total[0].replace('ORYGINAŁ\n', '')
                for i, j in miesiac.items():
                    if i in total[0]:
                        total[0] = total[0].replace(i, f'-{j}-')
                sidusdate = total[0]
            except IndexError:
                total = re.findall('Odbiorca \n.+ Przelew', text)
                total[0] = total[0].replace('Odbiorca \n', '')
                total[0] = total[0].replace(' Przelew', '')
                for i, j in miesiac.items():
                    if i in total[0]:
                        total[0] = total[0].replace(i, f'-{j}-')
                sidusdate = total[0]
            print(text, file=faktury)

            df2['Data'] = sidusdate
            df2['Data'] = df2['Data'].replace(regex=[r' '], value='')
            df2['Data'] = pd.to_datetime(df2['Data'], format='%d-%m-%Y', errors='coerce', utc=False)
            df2['Nr faktury'] = f'{os.path.basename(file)}'
            df2['Nr faktury'] = df2['Nr faktury'].replace(regex=[r'.pdf'], value='')
            df2['Nr faktury'] = df2['Nr faktury'].replace(regex=[r'_'], value='/')
            df2['Nr faktury'] = df2['Nr faktury'].replace(regex=[r'Faktura VAT'], value='')
            df2['Nr faktury'] = df2['Nr faktury'].replace(regex=[r'FV'], value='')
            df2 = df2[['Firma', 'Data', 'Nazwa Materiału', 'Ilość', 'Cena', 'Nr faktury']]
            print(df2)
            append_df_to_excel()

    for file in glob.glob(os.path.join(path, '*.jpg') or os.path.join(path, '*.png')):
        path_to_tesseract = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

        # Tesseract.exe to text
        pytesseract.tesseract_cmd = path_to_tesseract
        text = pytesseract.image_to_string(Image.open(file), lang='pol')

        material = []
        rajnert_material = []

        df = pd.DataFrame(data=None)

        # Possible to get this in 1 line but this will look unreadable and ugly
        material.append(re.findall('1\. .+\n+2|1. .+\n+PODSUMOWANIE|1\. .+\n+', text))
        material.append(re.findall('2\. .+\n+3|2. .+\n+PODSUMOWANIE|.{1,4} .+\n+3\.', text))
        material.append(re.findall('3\. .+\n+4|3. .+\n+PODSUMOWANIE|.{1,4} .+\n+4\.', text))
        material.append(re.findall('4\. .+\n+5|4. .+\n+PODSUMOWANIE|.{1,4} .+\n+5\.', text))
        material.append(re.findall('5\. .+\n+6|5. .+\n+PODSUMOWANIE|.{1,4} .+\n+6\.', text))
        material.append(re.findall('6\. .+\n+7|6. .+\n+PODSUMOWANIE|.{1,4} .+\n+7\.', text))
        material.append(re.findall('7\. .+\n+8|7. .+\n+PODSUMOWANIE|.{1,4} .+\n+8\.', text))
        material.append(re.findall('8\. .+\n+9|8. .+\n+PODSUMOWANIE|.{1,4} .+\n+9\.', text))
        material.append(re.findall('9\. .+\n+10|9. .+\n+PODSUMOWANIE|.{1,4} .+\n+10\.', text))
        material.append(re.findall('10\. .+\n+11|10. .+\n+PODSUMOWANIE|.{1,4} .+\n+11\.', text))
        material.append(re.findall('11\. .+\n+12|11. .+\n+PODSUMOWANIE|.{1,4} .+\n+12\.', text))
        material.append(re.findall('12\. .+\n+13|12. .+\n+PODSUMOWANIE|.{1,4} .+\n+13\.', text))
        material.append(re.findall('13\. .+\n+14|13. .+\n+PODSUMOWANIE|.{1,4} .+\n+14\.', text))
        material.append(re.findall('14\. .+\n+15|14. .+\n+PODSUMOWANIE|.{1,4} .+\n+15\.', text))

        for lista in material:
            rajnert_material.extend(lista)

        df['Nazwa Materiału'] = rajnert_material
        df['Nazwa Materiału'] = df['Nazwa Materiału'].replace(regex=[r'PODSUMOWANIE'], value='')
        df['Nazwa Materiału'] = df['Nazwa Materiału'].replace(regex=[r'\d. '], value='')
        df['Nazwa Materiału'] = df['Nazwa Materiału'].replace(regex=[r'\n\d'], value='')
        df['Nazwa Materiału'] = df['Nazwa Materiału'].replace(regex=[r'\n', r'\n\n'], value='')
        df['Nazwa Materiału'] = df['Nazwa Materiału'].replace(regex=[r'\*', r'\.'], value='')

        ilosc = re.findall('\n+.+ sz.? |\n+.+kp.? ', text)
        df['Ilość'] = ilosc
        df['Ilość'] = df['Ilość'].astype(str)
        df['Ilość'] = df['Ilość'].replace(regex=['\n+'], value='')
        df['Ilość'] = df['Ilość'].replace(regex=['sz.'], value='')
        df['Ilość'] = df['Ilość'].replace(regex=[' '], value='')
        df['Ilość'] = df['Ilość'].replace(regex=['kpl'], value='')

        cena = re.findall('\n+.+ sz. [0-9]+,[0,9]+|\n+.+ sz . [0-9]+,[0,9]+|\n+.+ kpl . [0-9]+,[0,9]+|\n+.+ sz.{1,7} [0-9]+,[0,9]+', text)
        df['Cena'] = cena
        df['Cena'] = df['Cena'].replace(regex=[r'\n+\d+'], value='')
        df['Cena'] = df['Cena'].replace(regex=['... " '], value='')
        df['Cena'] = df['Cena'].replace(regex=['... '], value='')
        df['Cena'] = df['Cena'].replace(regex=['... \* '], value='')
        df['Cena'] = df['Cena'].replace(regex=['sz.'], value='')

        df['Nr faktury'] = f'{os.path.basename(file)}'
        df['Nr faktury'] = df['Nr faktury'].replace(regex=[r'.jpg'], value='')
        df['Nr faktury'] = df['Nr faktury'].replace(regex=[r'Rajnert'], value='')
        df['Nr faktury'] = df['Nr faktury'].replace(regex=[r'_'], value='/')
        df['Nr faktury'] = df['Nr faktury'].replace(regex=[r'rajnert/'], value='')
        df['Nr faktury'] = df['Nr faktury'].replace(regex=[r'rajnert'], value='')

        rajnertdate = re.findall('dnia .+,', text)
        rajnertdate = ' '.join(rajnertdate)
        df['Data'] = rajnertdate
        df['Data'] = df['Data'].replace(regex=[r'dnia '], value='')
        df['Data'] = df['Data'].replace(regex=[r','], value='')
        df['Data'] = pd.to_datetime(df['Data'], format='%d-%m-%Y', errors='coerce', utc=False)

        df['Firma'] = 'Rajnert'

        df2 = df[['Firma', 'Data', 'Nazwa Materiału', 'Ilość', 'Cena', 'Nr faktury']]
        print(df2)

        append_df_to_excel()
        faktury = open("faktury.txt", "w+", encoding="utf-8")
        print(text, file=faktury)

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
    writer.sheets['Sheet1'].set_column(2, 2, 40)
    writer.sheets['Sheet1'].set_column(3, 3, 5)
    writer.sheets['Sheet1'].set_column(4, 4, 7)
    writer.sheets['Sheet1'].set_column(5, 5, 14)

    writer._save()
    subprocess.check_call(["attrib", "+H", "outpdf.xlsx"])

    print("-" * 60)
    print('Faktury zostały pomyślnie załadowane')

    # Archiwum faktur w przypadku zgubienia/zniszczenia głównego pliku
    try:
        archiwum = r'/Faktury Archiwum Obmiar GUI'
        os.makedirs(archiwum, exist_ok=False)

    except FileExistsError:
        archiwum = r'/Faktury Archiwum Obmiar GUI'
        df_arch = pd.read_excel('Faktury materiały.xlsx')
        df_arch.to_excel(os.path.join(archiwum, 'Faktury materiały.xlsx'))
        old_name = r"/Faktury Archiwum Obmiar GUI/Faktury materiały.xlsx"
        new_name = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        newname = f'{old_name} - {new_name}.xlsx'
        os.rename(old_name, newname)
