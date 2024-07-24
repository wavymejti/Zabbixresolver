import pandas as pd
import tkinter as tk

# Ustalona ścieżka do plików Excel
EXCEL_FILE_PATH_1 = "Telefony.xlsx"  # Ścieżka do pierwszego pliku
EXCEL_FILE_PATH_2 = "Lista_marketow_budowa.xlsx"  # Ścieżka do drugiego pliku
EXCEL_FILE_PATH_3 = "Umowy_internet.xlsx"  # Ścieżka do trzeciego pliku

def read_excel_and_display_row_1(numer_lokalizacji):
    try:
        # Odczytujemy pierwszy plik Excel
        df = pd.read_excel(EXCEL_FILE_PATH_1)
        
        # Sprawdzamy, czy kolumna 'nr lokalizacji' istnieje
        if 'nr lokalizacji' not in df.columns:
            result_label.config(text="Kolumna 'nr lokalizacji' nie istnieje w pierwszym pliku Excel.")
            return ""
        
        # Szukamy wiersza z podanym numerem lokalizacji
        row = df[df['nr lokalizacji'] == numer_lokalizacji]
        
        # Sprawdzamy czy znaleziono wiersz
        if not row.empty:
            nr_tel_kierownika = row['nr tel. kierownika'].values[0]
            nr_tel_zastepcy = row['nr tel. zastępcy'].values[0]
            result_text = (f"Numer tel. kierownika: {nr_tel_kierownika}\n"
                           f"Numer tel. zastępcy: {nr_tel_zastepcy}")
            result_label.config(text=f"Pierwszy plik:\n{result_text}")
            
            return result_text
        else:
            result_label.config(text=f"Numer lokalizacji {numer_lokalizacji} nie został znaleziony w pierwszym pliku.")
            return ""
    except Exception as e:
        result_label.config(text=f"Wystąpił błąd podczas odczytu pierwszego pliku Excel: {e}")
        return ""

def read_excel_and_display_row_2(numer_gold):
    try:
        # Odczytujemy drugi plik Excel
        df = pd.read_excel(EXCEL_FILE_PATH_2, skiprows=5)  # Pomijamy pierwsze 5 wierszy, aby zacząć od C6
        
        # Sprawdzamy, czy kolumna 'NR GOLD' istnieje
        if 'NR GOLD' not in df.columns:
            result_label_2.config(text="Kolumna 'NR GOLD' nie istnieje w drugim pliku Excel.")
            return ""
        
        # Szukamy wiersza z podanym numerem GOLD
        row = df[df['NR GOLD'] == numer_gold]
        
        # Sprawdzamy czy znaleziono wiersz
        if not row.empty:
            status = row['STATUS'].values[0]  # Pobieramy wartość z kolumny "STATUS" (kolumna K)
            ulica_nr = row['ULICA, NR'].values[0]  # Pobieramy wartość z kolumny "ULICA, NR" (kolumna L)
            result_text = f"Status: {status}\nUlica, Nr: {ulica_nr}"
            result_label_2.config(text=f"Drugi plik:\n{result_text}")
            return result_text
        else:
            result_label_2.config(text=f"Numer GOLD {numer_gold} nie został znaleziony w drugim pliku.")
            return ""
    except Exception as e:
        result_label_2.config(text=f"Wystąpił błąd podczas odczytu drugiego pliku Excel: {e}")
        return ""

def read_excel_and_display_row_3(numer_lokalizacji):
    try:
        # Odczytujemy trzeci plik Excel
        df = pd.read_excel(EXCEL_FILE_PATH_3)
        
        # Sprawdzamy, czy kolumna 'nr lokalizacji' istnieje
        if 'nr lokalizacji' not in df.columns:
            result_label_3.config(text="Kolumna 'nr lokalizacji' nie istnieje w trzecim pliku Excel.")
            return ""
        
        # Szukamy wiersza z podanym numerem lokalizacji
        row = df[df['nr lokalizacji'] == numer_lokalizacji]
        
        # Sprawdzamy czy znaleziono wiersz
        if not row.empty:
            firma = row['Firma '].values[0]  # Pobieramy wartość z kolumny "Firma " (kolumna D)
            result_text = f"Firma: {firma}"
            result_label_3.config(text=f"Trzeci plik:\n{result_text}")
            return result_text
        else:
            result_label_3.config(text=f"Numer lokalizacji {numer_lokalizacji} nie został znaleziony w trzecim pliku.")
            return ""
    except Exception as e:
        result_label_3.config(text=f"Wystąpił błąd podczas odczytu trzeciego pliku Excel: {e}")
        return ""

def submit():
    numer_lokalizacji = entry_numer_lokalizacji.get()
    zmienna = entry_zmienna.get()
    
    # Walidacja numeru lokalizacji (czy jest liczbą)
    try:
        numer_lokalizacji = int(numer_lokalizacji)
    except ValueError:
        result_label.config(text="Numer lokalizacji musi być liczbą.")
        return
    
    print(f"Numer lokalizacji: {numer_lokalizacji}")
    print(f"Zmienna: {zmienna}")
    
    result_text_1 = read_excel_and_display_row_1(numer_lokalizacji)
    result_text_2 = read_excel_and_display_row_2(numer_lokalizacji)
    result_text_3 = read_excel_and_display_row_3(numer_lokalizacji)
    
    # Kopiowanie do schowka
    full_result_text = f"{result_text_1}\n{result_text_2}\n{result_text_3}"
    root.clipboard_clear()
    root.clipboard_append(full_result_text.strip())
    root.update()  # Aby aplikacja nie była "zamrożona"

    # Informacja o skopiowaniu do schowka
    copied_label.config(text="Wszystkie informacje zostały skopiowane do schowka.")

# Tworzenie okna Tkinter
root = tk.Tk()
root.title("Formularz")

tk.Label(root, text="Numer lokalizacji").grid(row=0)
tk.Label(root, text="Zmienna").grid(row=1)

entry_numer_lokalizacji = tk.Entry(root)
entry_zmienna = tk.Entry(root)

entry_numer_lokalizacji.grid(row=0, column=1)
entry_zmienna.grid(row=1, column=1)

tk.Button(root, text='Submit', command=submit).grid(row=2, column=1, sticky=tk.W, pady=4)

result_label = tk.Label(root, text="", anchor="w", justify="left")
result_label.grid(row=3, column=0, columnspan=2, sticky="w")

result_label_2 = tk.Label(root, text="", anchor="w", justify="left")
result_label_2.grid(row=4, column=0, columnspan=2, sticky="w")

result_label_3 = tk.Label(root, text="", anchor="w", justify="left")
result_label_3.grid(row=5, column=0, columnspan=2, sticky="w")

copied_label = tk.Label(root, text="", anchor="w", justify="left")
copied_label.grid(row=6, column=0, columnspan=2, sticky="w")

root.mainloop()