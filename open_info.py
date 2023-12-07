import customtkinter
from PIL import Image, ImageTk
from import_images import resource_path

def open_info():
    openinfo = customtkinter.CTkToplevel()

    openinfo.title("Informacje od autora")
    openinfo.geometry("790x790")

    #otwiera okno podobnie jak transient
    openinfo.attributes('-topmost', 1)

    openinfo_frame = customtkinter.CTkScrollableFrame(openinfo, width=750, height=750, border_width=1)
    openinfo_frame.grid(row=0, column=0, padx=10, pady=10)
    image1 = resource_path(r"Images/Rozszerzenie pliku xlsx.jpg")
    image2 = resource_path(r"Images/xlsx.jpg")
    image3 = resource_path(r"Images/Kolumna L2.jpg")
    image4 = resource_path(r"Images/Kolumna L.jpg")
    image5 = resource_path(r"Images/1 strzalka.jpg")
    image6 = resource_path(r"Images/2 strzalka.jpg")
    image7 = resource_path(r"Images/3 strzalka.jpg")
    image8 = resource_path(r"Images/Format Obmiaru.jpg")
    im1 = ImageTk.PhotoImage(Image.open(image1))
    im2 = ImageTk.PhotoImage(Image.open(image2))
    im3 = ImageTk.PhotoImage(Image.open(image3))
    im4 = ImageTk.PhotoImage(Image.open(image4))
    im5 = ImageTk.PhotoImage(Image.open(image5))
    im6 = ImageTk.PhotoImage(Image.open(image6))
    im7 = ImageTk.PhotoImage(Image.open(image7))
    im8 = ImageTk.PhotoImage(Image.open(image8))

    labelinfo1 = customtkinter.CTkLabel(openinfo_frame, text="1. Przed załadowaniem nowego Obmiaru należy sprawdzić, czy plik ma rozszerzenie xlsx (zwykły plik excela).\n "
                                                        "Jeżeli ma inne rozszerzenie niż wymagane wystarczy Obmiar zapisać jako Skoroszyt Programu Excel.\n"
                                                        "Próba załadowania Obmiaru z innym rozszerzeniem zakończy się niepowodzeniem.", justify="left")
    labelinfo1.grid(row=0, column=2, padx=1, pady=1, sticky="nw")

    labelinfo2 = customtkinter.CTkLabel(openinfo_frame, image=im1)
    labelinfo2.grid(row=0, column=2, padx=1, pady=48, sticky="nw")

    labelinfo3 = customtkinter.CTkLabel(openinfo_frame, image=im2)
    labelinfo3.grid(row=0, column=2, padx=1, pady=248, sticky="nw")

    labelinfo4 = customtkinter.CTkLabel(openinfo_frame, text=f"2. Aby dało się poprawie utworzyć bazę danych obmiar musi spełniać kilka warunków.\n"
                                                       f"W kolumnie L2 powinien znajdować się tekst Nr Suwnicy oraz w L3 nazwa urządzenia (np. tytuł arkusza.\n"
                                                       f"Nie jest konieczne aby znajdowała się tam formuła, taka jak na obrazie poniżej.\n"
                                                       f"Dodawanie nowego urządzenia wymaga jedynie dodania w kolumnie L2 i L3 powyższych danych.", justify="left")
    labelinfo4.grid(row=0, column=2, padx=1, pady=380, sticky="nw")

    labelinfo5 = customtkinter.CTkLabel(openinfo_frame, image=im3)
    labelinfo5.grid(row=0, column=2, padx=1, pady=450, sticky="nw")

    labelinfo6 = customtkinter.CTkLabel(openinfo_frame, image=im4)
    labelinfo6.grid(row=0, column=2, padx=1, pady=710, sticky="nw")

    labelinfo7 = customtkinter.CTkLabel(openinfo_frame, text="3. Kolejnym warunkiem jest format wpisywania danych do Obmiaru (niezmienny od lat).\n"
                                                             "Data musi być zawsze w formacie dzień-miesiąc-rok+r. (np. 11-02-2023r.).\n"
                                                             "Materiały muszą się zaczynać od myślnika - oraz kończyć myślnikiem - ilścią i jednostką (- 2 szt) (- 55 mb).", justify="left")
    labelinfo7.grid(row=0, column=2, padx=1, pady=940, sticky="nw")

    labelinfo8 = customtkinter.CTkLabel(openinfo_frame, image=im8)
    labelinfo8.grid(row=0, column=2, padx=1, pady=1000, sticky="nw")

    labelinfo9 = customtkinter.CTkLabel(openinfo_frame, text="4. Lista z nr urządzeń oraz lista z latami będą aktualizowane automatycznie.\n"
                                                             "Dla ułatwienia powrotu z wybranego urządzenia do opcji WSZYSTKIE\n"
                                                             f"należy wcisnąć klawisz strzałki w dół \/ oraz wcisnąć ENTER.\n"
                                                             "W polu szukaj można wpisywać materiały oraz nr urządzeń (najlepiej kilka ostatnich cyfr).\n"
                                                             "Wielkość liter nie ma znaczenia. Klikając w tytuły kolumn mamy możliwość sortowania jej elementów", justify="left")
    labelinfo9.grid(row=0, column=2, padx=1, pady=1500, sticky="nw")

    labelinfo10 = customtkinter.CTkLabel(openinfo_frame, image=im5)
    labelinfo10.grid(row=0, column=2, padx=1, pady=1600, sticky="nw")

    labelinfo11 = customtkinter.CTkLabel(openinfo_frame, image=im6)
    labelinfo11.grid(row=0, column=2, padx=1, pady=1800, sticky="nw")

    labelinfo12 = customtkinter.CTkLabel(openinfo_frame, image=im7)
    labelinfo12.grid(row=0, column=2, padx=1, pady=1800, sticky="n")

    labelinfo14 = customtkinter.CTkLabel(openinfo_frame, text="5. Ładowanie faktur. Tylko firmy Sidus, Windex, Rajnert.\n"
                                                              "Należy utworzyć folder, w którym będą znajdowały się wszystkie faktury, a następnie go wybrać.\n"
                                                              "Do przetworzenia plików jpg, png służy program tesseract, a jego ścieżka musi wyglądać w następujący sposób:\n"
                                                              r"C:\Program Files\Tesseract-OCR\tesseract.exe.""\n"
                                                              "Przy ładowanie zostaje utworzone archiwum w wypadku nieznanego błędu. Znajduje się na dysku C:\n"
                                                              "C:/Faktury Archiwum Obmiar GUI/. W przypadku błędów lub wielokrotnego niezałądowania materiałow należy usunąć aktualny plik.", justify="left")
    labelinfo14.grid(row=0, column=2, padx=1, pady=2130, sticky="nw")

    labelinfo15 = customtkinter.CTkLabel(openinfo_frame,
                                         text="6. Karty pracy dostępne są tylko z suwnic. Należy wybrać miesiąc i rok z jakiego mają zostać wygenerowane.\n"
                                            "Następnie wymagane jest załadowanie obmiaru. Plik musi mieć rozszedzenie xlsx, binary nie jest obsługiwany.\n"
                                            "Zostanie utworzony plik karty_pracy_suwnice w miejscu, gdzie znajduje się główny program obmiar gui.exe.\n"
                                            "Zalecam sprawdzenie, czy wyszstkie prace są umieszczone w wygenerowanym pliku. Ani razu nie miałem ani jednego błędu,\n"
                                            "w przypadku źle wypełnionego obmiaru mogą się takowe utworzyć.",
                                         justify="left")
    labelinfo15.grid(row=0, column=2, padx=1, pady=2250, sticky="nw")

    labelinfo16 = customtkinter.CTkLabel(openinfo_frame, text="7. Prace Suwnice. Plik musi być zamknięty aby program zaczął działać. Wpisujemy nazwę akrkusza np 10-2023, 01-2024 etc.\n"
                                                                "Program wyszukuje takich samych nazw 1:1 (w przypadku wyszukiwania podonych nawet w 95% wypełniało nieprawidłowo).\n"
                                                                "Ceny pobierane są z bazy, która jest utworzona poprzez ładowanie nowego obmiaru\n"
                                                                "Jedynym zauważonym przeze mnie błędem jest przeliczanie śrub DIN w zły sposób, prawdopodobnie kiedyś w obmiarze\n"
                                                                "zostało to źle wpisane, poprawienie tych źle wpisanych danych naprawi ten błąd.",
                                         justify="left")
    labelinfo16.grid(row=0, column=2, padx=1, pady=2350, sticky="nw")

    labelinfo17 = customtkinter.CTkLabel(openinfo_frame, text="8. Wykresy. Generowany jest plik w formacie .ini. Znajdują się tam dane takie jak cena za rbg, mateirały, rok.\n"
                                                                "Plik jest w pełni edytowalny, można w nim ręcznie wpisywać wartości i będą czyatne przez program.\n"
                                                                "Wykresy zostają automatycznie zapisywane do plików png w miejscu, gdzie znajduje się główny program .exe.",
                                                        justify="left")
    labelinfo17.grid(row=0, column=2, padx=1, pady=2450, sticky="nw")



    labelinfo13 = customtkinter.CTkLabel(openinfo_frame, text="Autor projektu: Miłosz Maguda\n Wersja: v1.5\n wersja z wykresami\n\n\n"
                                                            "Program działa w wersji 32bit, możliwe jest utworzenie wersji 64bit\n"
                                                            "Kod do programu znajduje się w repozytorium Github pod adresem:\n"
                                                            "https://github.com/takixdd/Quantity-survey.git\n"
                                                            "Kontakt: trzecitaki@gmail.com",
                                         justify="right")
    labelinfo13.grid(row=0, column=2, padx=90, pady=2800, sticky="nw")
    openinfo.mainloop()
