import os
import win32com.client
import tkinter as tk
from tkinter import filedialog
import datetime
import pathlib
import re
import pywintypes, win32file, win32con
from tqdm import tqdm  # Import für den Fortschrittsbalken


def create_file_list(Such_Pfad):
    # Erstellt eine Liste mit allen Dateien im angegebenen Suchpfad
    # Input: Suchpfad
    # Output: Liste aller Dateien (pathlib.Path Objekte)
    
    return list(pathlib.Path(Such_Pfad).rglob('*'))


def load_2column_list(file_path):
    # AH 04.11.24
    # Lädt eine Liste aus einer Textdatei.
    # Die Liste muss pro Zeile 2 Einträge enthalten, die durch Semikolon getrennt sind.
    # Einträge mit der falschen Anzahl an Einträge werden ignoriert
    # Zeilen, die mit einem # beginnen, werden übersprungen.
    # param file_path: Pfad zur Ersetzungslisten-Datei.
    # return: Ein Wörterbuch mit Suchtext als Schlüssel und Ersatztext als Wert.

    liste = {}
    korrekt_geladen = True
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            for line_number, line in enumerate(file, start=1):
                line = line.strip()
                if not line or line.startswith('#'):
                    continue
                parts = line.split(';')
                if len(parts) != 2:
                    print(f"Warnung: Zeile {line_number} in '{file_path}' enthält nicht genau zwei Einträge und wird ignoriert.")
                    continue
                search_text, replace_text = parts
                print(search_text, replace_text)
                liste[search_text] = replace_text
    except FileNotFoundError:
        print(f"Datei '{file_path}' nicht gefunden.")
        korrekt_geladen = False
    except Exception as e:
        print(f"Ein Fehler ist aufgetreten beim Laden der Bereinigungsliste {file_path}: {str(e)}")
        korrekt_geladen = False
    return liste, korrekt_geladen

# Ende der Funktion "load_2column_list"


def Search_Files(Search, all_files):
    # Durchsucht eine vorgegebene Liste nach Dateien, die dem Suchmuster entsprechen
    # Input: Dateinamen (mit Extension), Liste aller Dateien
    # Output: Liste mit Treffern (kompletter Pfad mit Dateiname und Extension)
    Liste = [str(path) for path in all_files if path.name.startswith(Search)]
    return Liste

def Rename_Files_and_new_date(Liste, NewName):
    # Benennt Datein um
    # AH 02.04.24
    # Import: os
    # Input:  Liste mit Namen mit vollem Pfad und Dateinamen (mit Extension)
    #         Neuer Name (mit Extension)
    # Output: -
    #
    # Open Items:
    #   Fehlerabfrage Datein in Liste nicht gefunden
    #   Fehlerabfrage beim Umbennen 
    #

    for path in Liste:

        old_path = path
        new_path = os.path.join(os.path.dirname(path), NewName)
        print(f"Datei {old_path} soll in {new_path} umbenannt")
        os.rename(old_path, new_path)
        print(f"Datei {old_path} wurde in {new_path} umbenannt")
        Log_Eintrag_schreiben(f"Datei {old_path} wurde in {new_path} umbenannt")

        # Falls der neue Filename ein Datum enthält wird das Erstelldatum angepasst
        new_date, flag = Date_Extract(NewName)
        if flag:
            set_creation_date(new_path, new_date)


    # Ende Rename_Files_and_new_date

def Rename_Files(Liste, NewName):
    # Benennt Datein um
    # AH 02.04.24
    # Import: os
    # Input:  Liste mit Namen mit vollem Pfad und Dateinamen (mit Extension)
    #         Neuer Name (mit Extension)
    # Output: -
    #
    # Open Items:
    #   Fehlerabfrage Datein in Liste nicht gefunden
    #   Fehlerabfrage beim Umbennen 
    #
    
    for path in Liste:
        old_path = path
        new_path = os.path.join(os.path.dirname(path), NewName)
        os.rename(old_path, new_path)
        print(f"Datei {old_path} wurde in {new_path} umbenannt")
        Log_Eintrag_schreiben(f"Datei {old_path} wurde in {new_path} umbenannt")

    # Ende Rename_Files

def Target_in_LNK(Liste,New_Target):
    # Benennen Ziele in LNK Dateien anpassen
    for path in Liste:
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(path)
        shortcut.TargetPath = New_Target
        shortcut.Save()
        print(f"Ziel in der Verknüpfung {path} wurde auf {New_Target} geändert")
        Log_Eintrag_schreiben(f"Ziel in der Verknüpfung {path} wurde auf {New_Target} geändert")
    # Ende Target_in_LNK

def choose_filename_and_replacementname(default_path):
    root = tk.Tk()
    root.withdraw()  # Verstecke das Hauptfenster

    file_path = filedialog.askopenfilename(initialdir=default_path)
    if file_path:
        # Zeige den ausgewählten Dateinamen an
        #print("Ausgewählte Datei:", file_path)

        # Extrahiere den ursprünglichen Dateinamen ohne Pfad
        original_filename = os.path.basename(file_path)

        # Extrahiere die Dateierweiterung
        _, file_extension = os.path.splitext(original_filename)

        # Nutzer kann den Dateinamen bearbeiten
        new_filename = filedialog.asksaveasfilename(defaultextension=file_extension, initialfile=original_filename)

        # Speichere den neuen Dateinamen in einer Variable
        if new_filename:
            #print("Neuer Dateiname gespeichert:", new_filename)
            return file_path, new_filename

        else:
            print("Kein neuer Dateiname eingegeben.")
    else:
        print("Keine Datei ausgewählt.")

def choose_filename_with_path():
    root = tk.Tk()
    root.withdraw()  # Verstecke das Hauptfenster
    file_path = filedialog.askopenfilename()
    if file_path:
        return file_path
    else:
        print("Keine Datei ausgewählt.")

def Log_Eintrag_schreiben(message):
    path = r"\\WIN-H7BKO5H0RMC\Dataserver\Python Hilfsdateien\Email_Analyse_Log.txt"
    current_datetime = datetime.datetime.now()
    log_entry = f"{current_datetime.strftime('%Y%m%d %H%M')} - File Rename incl Copies and Links: {message}\n"
   
    with open(path, "a") as log_file:
        log_file.write(log_entry)

    # ende "Log_Eintrag_schreiben"

def show_selection_box(prompt1, prompt2):
    def handle_selection(selection):
        root.selection = selection
        root.destroy()  # Schließe das Hauptfenster

    root = tk.Tk()
    root.title("Optionenauswahl")

    button1 = tk.Button(root, text=prompt1, command=lambda: handle_selection(prompt1))
    button2 = tk.Button(root, text=prompt2, command=lambda: handle_selection(prompt2))

    button1.pack()
    button2.pack()

    # Fenster in der Mitte des Bildschirms positionieren
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    window_width = 300
    window_height = 200
    x_pos = (screen_width - window_width) // 2
    y_pos = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x_pos}+{y_pos}")

    root.mainloop()
    return getattr(root, 'selection', 'Keine Auswahl')


def Date_Extract(string):
    
    date_pattern = re.compile(r'\d{8}')  # Muster für 8-stellige Zahlen
    match = date_pattern.search(string)
    if match:
        date_string = match.group()
    
        try:
            date = datetime.datetime.strptime(date_string, '%Y%m%d')
            return date, True
    
        except ValueError:
            return None, False
            
    else:
        return None, False

def set_creation_date(file_path, new_creation_date):
    #kopiert aus stackoverflow: https://stackoverflow.com/questions/4996405/how-do-i-change-the-file-creation-date-of-a-windows-file

    wintime = pywintypes.Time(new_creation_date)
    winfile = win32file.CreateFile(
        file_path, win32con.GENERIC_WRITE,
        win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
        None, win32con.OPEN_EXISTING,
        win32con.FILE_ATTRIBUTE_NORMAL, None)

    try:
        win32file.SetFileTime(winfile, wintime, None, None)
        #print(f"Erstelldatum für {winfile} auf {wintime} gesetzt")
        Log_Eintrag_schreiben(f"Erstelldatum für {file_path} auf {wintime} gesetzt")
    except:
        #print(f"Fehler: Erstelldatum für {winfile} nicht verändert.")
        Log_Eintrag_schreiben(f"Fehler: Erstelldatum für {file_path} nicht verändert.")

    winfile.close()

"""
Das Programm sucht alle Files mit einem bestimmten Namen und die LNK Files mit dem gleichen Namen
Das Ursprungsfile und alle Kopien, werden umbenannt
Die Verweise in den LNK Files werden angepaßt
Die LNK Files werden umbenannt

Es werden keine LNK Files gesucht, die einen Verweis auf das Ursprungsfile beinhalten, aber selbst einen anderen Filenamen haben

Batch Datei mus alten namen und neuen Namen beides ohne Pfadangabe durch Semikolon getrennte enthalten. Keine Leerzeilen am Ende. 
alter und neuer name muss mit extension sein!

"""


# Der Pfad, in dem Sie nach Dateien suchen möchten
#Such_Pfad = r"C:\Users\ah\Desktop\Email Analyse\Testumgebung für Umbenennung"   # Testumgebung
#Basis_Pfad = r"C:\Users\ah\Desktop\Email Analyse\Testumgebung für Umbenennung"   # Testumgebung

Such_Pfad = pathlib.Path(r"\\WIN-H7BKO5H0RMC\Dataserver")
Basis_Pfad = pathlib.Path(r"\\WIN-H7BKO5H0RMC\Dataserver\Korrespondenz\Post Archiv")

# Pfadangaben für Testumgebung
#Such_Pfad  = r"\\WIN-H7BKO5H0RMC\Dataserver\Python Testumgebungen\File Rename incl Copies and Links\S_Dataserver" 
#Basis_Pfad = r"\\WIN-H7BKO5H0RMC\Dataserver\Python Testumgebungen\File Rename incl Copies and Links\S_Dataserver\Korrespondenz\Post_Archiv"

# Erstelle eine Liste aller Dateien im Suchpfad einmal vor der Schleife

all_files_list = create_file_list(Such_Pfad)

Auswahl = show_selection_box("Einzelfile", "Batchdatei")

if Auswahl == "Batchdatei":
    Bereinigung_File = choose_filename_with_path()
    print(Bereinigung_File)
    if Bereinigung_File:
        Bereinigung_Liste, flagb = load_2column_list(Bereinigung_File)
        if not flagb:
            print(f"Fehler beim Laden der Batch-Liste {Bereinigung_File}")
else:
    try:
        # Benutzer wählt alten und neuen Dateinamen aus
        NameAlt_full, NameNeu_full = choose_filename_and_replacementname(Basis_Pfad)
        NameAlt_Name = pathlib.Path(NameAlt_full).name
        NameNeu_Name = pathlib.Path(NameNeu_full).name
        Bereinigung_Liste = {NameAlt_Name: NameNeu_Name}
        flagb = True
    except:
        print("Dateiauswahl abgebrochen.")
        flagb = False

if flagb:
    for old_name, new_name in Bereinigung_Liste.items():
        FileList = Search_Files(old_name, all_files_list)

        # Entfernt die Erweiterung vom Ursprungsnamen und fügt .lnk hinzu
        NameLNKalt = pathlib.Path(old_name).stem + ".lnk"
        NameLNKneu = pathlib.Path(new_name).stem + ".lnk"

        FileLNKList = Search_Files(NameLNKalt, all_files_list)

        new_path_full = os.path.join(Basis_Pfad, new_name)

        Rename_Files_and_new_date(FileList, new_name)
        Target_in_LNK(FileLNKList, new_path_full)
        Rename_Files(FileLNKList, NameLNKneu)
else:
    print("Programm Ende")
