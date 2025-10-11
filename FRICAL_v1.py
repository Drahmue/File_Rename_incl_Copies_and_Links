import os
import sys
import win32com.client
import tkinter as tk
from tkinter import filedialog
import datetime
import pathlib
import re
import pywintypes, win32file, win32con
from tqdm import tqdm  # Import für den Fortschrittsbalken
from multiprocessing import Pool, cpu_count
from concurrent.futures import ProcessPoolExecutor, as_completed

try:
    from ahlib import StructuredConfigParser, create_extended_logger
except ImportError:
    print("FEHLER: ahlib konnte nicht importiert werden.")
    print("Bitte installieren Sie ahlib mit:")
    print("  pip install git+https://github.com/Drahmue/ahlib.git")
    sys.exit(1)


def _scan_directory_recursive(directory_path):
    """
    Hilfsfunktion für paralleles Scannen eines einzelnen Verzeichnisses rekursiv.

    Args:
        directory_path: Pfad zum zu scannenden Verzeichnis (als String)

    Returns:
        Liste von pathlib.Path Objekten für alle Dateien und Verzeichnisse
    """
    results = []
    try:
        dir_path = pathlib.Path(directory_path)
        if dir_path.is_dir():
            # Rekursiv alle Items sammeln
            for item in dir_path.rglob('*'):
                results.append(item)
    except (PermissionError, OSError):
        # Überspringe Verzeichnisse ohne Zugriffsberechtigung
        pass
    return results


def create_file_list(Such_Pfad):
    """
    Erstellt eine Liste mit allen Dateien im angegebenen Suchpfad.
    Verwendet Multiprocessing für schnelleres Scannen großer Verzeichnisstrukturen.

    Input: Suchpfad
    Output: Liste aller Dateien (pathlib.Path Objekte)
    """
    Such_Pfad = pathlib.Path(Such_Pfad)

    # Sammle erst die direkten Unterverzeichnisse der ersten Ebene
    try:
        first_level_dirs = [item for item in Such_Pfad.iterdir() if item.is_dir()]
    except (PermissionError, OSError):
        # Fallback zur einfachen Variante wenn kein Zugriff
        print("Erstelle Dateiliste (einfacher Modus)...")
        return list(Such_Pfad.rglob('*'))

    # Wenn wenige Verzeichnisse, nutze einfache Variante mit Progress Bar
    if len(first_level_dirs) < 4:
        print("Erstelle Dateiliste (einfacher Modus)...")
        all_items = []
        with tqdm(desc="Scanne Verzeichnisse", unit=" Dateien", dynamic_ncols=True) as pbar:
            for item in Such_Pfad.rglob('*'):
                all_items.append(item)
                pbar.update(1)
        return all_items

    # Multiprocessing für große Verzeichnisstrukturen
    print(f"Erstelle Dateiliste mit {cpu_count()} Prozessoren (parallel)...")
    all_files = []

    # Dateien im Root-Verzeichnis direkt hinzufügen
    try:
        for item in Such_Pfad.iterdir():
            if not item.is_dir():
                all_files.append(item)
    except (PermissionError, OSError):
        pass

    # Paralleles Scannen der Unterverzeichnisse
    num_workers = min(cpu_count(), len(first_level_dirs))

    try:
        with ProcessPoolExecutor(max_workers=num_workers) as executor:
            # Starte parallele Scans für jedes Unterverzeichnis der ersten Ebene
            future_to_dir = {
                executor.submit(_scan_directory_recursive, str(d)): d
                for d in first_level_dirs
            }

            # Sammle Ergebnisse mit Fortschrittsanzeige
            with tqdm(total=len(first_level_dirs),
                     desc="Scanne Verzeichnisse",
                     unit=" Ordner",
                     dynamic_ncols=True) as pbar:
                for future in as_completed(future_to_dir):
                    try:
                        dir_results = future.result()
                        all_files.extend(dir_results)
                        pbar.update(1)
                        # Zeige aktuelle Anzahl gefundener Dateien
                        pbar.set_postfix({"Dateien": len(all_files)})
                    except Exception:
                        # Fehler beim Scannen eines Verzeichnisses ignorieren
                        pbar.update(1)
                        pass
    except Exception:
        # Fallback zur einfachen Variante bei Multiprocessing-Problemen
        print("Fallback zu einfachem Modus...")
        return list(Such_Pfad.rglob('*'))

    return all_files


def create_file_list_simple(Such_Pfad):
    """
    Einfache synchrone Version der Dateilistenerstellung (Fallback).

    Input: Suchpfad
    Output: Liste aller Dateien (pathlib.Path Objekte)
    """
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
        logger.info(f"Datei {old_path} soll in {new_path} umbenannt")
        os.rename(old_path, new_path)
        logger.info(f"Datei {old_path} wurde in {new_path} umbenannt")

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
        logger.info(f"Datei {old_path} wurde in {new_path} umbenannt")

    # Ende Rename_Files

def Target_in_LNK(Liste,New_Target):
    # Benennen Ziele in LNK Dateien anpassen
    for path in Liste:
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(path)
        shortcut.TargetPath = New_Target
        shortcut.Save()
        logger.info(f"Ziel in der Verknüpfung {path} wurde auf {New_Target} geändert")
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

# Logging is now handled by ExtendedLogger from ahlib
# The old Log_Eintrag_schreiben function has been removed

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


def show_preview_window(old_name, new_name, files_to_rename, links_to_adjust):
    """
    Zeigt ein Vorschaufenster mit allen Dateien und Links, die umbenannt/angepasst werden.

    Args:
        old_name: Alter Dateiname
        new_name: Neuer Dateiname
        files_to_rename: Liste der umzubenennenden Dateien (mit Pfad)
        links_to_adjust: Liste der anzupassenden .lnk Dateien (mit Pfad)

    Returns:
        bool: True wenn Benutzer bestätigt, False bei Abbruch
    """
    # Variable für das Ergebnis - außerhalb des Tkinter-Kontexts
    result = {'confirmed': False}

    def handle_confirm():
        result['confirmed'] = True
        root.quit()
        root.destroy()

    def handle_cancel():
        result['confirmed'] = False
        root.quit()
        root.destroy()

    root = tk.Tk()
    root.title("Vorschau - Umbenennungsvorgang")

    # Fenster-Dimensionen
    window_width = 800
    window_height = 600
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_pos = (screen_width - window_width) // 2
    y_pos = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x_pos}+{y_pos}")

    # Header mit altem und neuem Namen
    header_frame = tk.Frame(root, bg="#e0e0e0", padx=10, pady=10)
    header_frame.pack(fill=tk.X, padx=5, pady=5)

    tk.Label(header_frame, text="Alter Name:", font=("Arial", 10, "bold"), bg="#e0e0e0").grid(row=0, column=0, sticky="w")
    tk.Label(header_frame, text=old_name, font=("Arial", 10), bg="#e0e0e0").grid(row=0, column=1, sticky="w", padx=10)

    tk.Label(header_frame, text="Neuer Name:", font=("Arial", 10, "bold"), bg="#e0e0e0").grid(row=1, column=0, sticky="w")
    tk.Label(header_frame, text=new_name, font=("Arial", 10), bg="#e0e0e0").grid(row=1, column=1, sticky="w", padx=10)

    # Hauptbereich mit Scrollbar
    main_frame = tk.Frame(root)
    main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    # Textbereich für die Dateiliste
    text_widget = tk.Text(main_frame, wrap=tk.WORD, font=("Courier", 9))
    scrollbar = tk.Scrollbar(main_frame, command=text_widget.yview)
    text_widget.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Inhalt erstellen
    content = []

    # Dateien zum Umbenennen
    content.append("=" * 80)
    content.append(f"DATEIEN ZUM UMBENENNEN ({len(files_to_rename)})")
    content.append("=" * 80)
    if files_to_rename:
        for file_path in files_to_rename:
            content.append(f"  {file_path}")
    else:
        content.append("  (Keine Dateien gefunden)")
    content.append("")

    # Links zum Anpassen
    content.append("=" * 80)
    content.append(f"VERKNÜPFUNGEN ZUM ANPASSEN ({len(links_to_adjust)})")
    content.append("=" * 80)
    if links_to_adjust:
        for link_path in links_to_adjust:
            content.append(f"  {link_path}")
    else:
        content.append("  (Keine Verknüpfungen gefunden)")
    content.append("")

    # Zusammenfassung
    content.append("=" * 80)
    content.append("ZUSAMMENFASSUNG")
    content.append("=" * 80)
    content.append(f"  Dateien zum Umbenennen:      {len(files_to_rename)}")
    content.append(f"  Verknüpfungen zum Anpassen:  {len(links_to_adjust)}")
    content.append(f"  Gesamt:                      {len(files_to_rename) + len(links_to_adjust)}")

    # Text einfügen
    text_widget.insert("1.0", "\n".join(content))
    text_widget.configure(state="disabled")  # Readonly

    # Button-Frame
    button_frame = tk.Frame(root, padx=10, pady=10)
    button_frame.pack(fill=tk.X)

    cancel_button = tk.Button(button_frame, text="Abbrechen", command=handle_cancel,
                              bg="#ff6b6b", fg="white", font=("Arial", 10, "bold"),
                              padx=20, pady=5)
    cancel_button.pack(side=tk.RIGHT, padx=5)

    confirm_button = tk.Button(button_frame, text="Umbenennung durchführen", command=handle_confirm,
                               bg="#51cf66", fg="white", font=("Arial", 10, "bold"),
                               padx=20, pady=5)
    confirm_button.pack(side=tk.RIGHT, padx=5)

    root.mainloop()
    return result['confirmed']


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
        logger.info(f"Erstelldatum für {file_path} auf {wintime} gesetzt")
    except:
        logger.error(f"Fehler: Erstelldatum für {file_path} nicht verändert.")

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

# Globale Variable für logger (wird im Hauptprozess initialisiert)
logger = None


def perform_rename_operation(all_files_list, Basis_Pfad, logger_instance):
    """
    Führt einen einzelnen Umbenennungsvorgang durch.

    Args:
        all_files_list: Liste aller Dateien im Suchpfad
        Basis_Pfad: Basispfad für neue Dateien
        logger_instance: Logger-Instanz für Logging

    Returns:
        bool: True wenn erfolgreich, False bei Abbruch
    """
    # Verwende den übergebenen Logger
    global logger
    logger = logger_instance

    Auswahl = show_selection_box("Einzelfile", "Batchdatei")

    if Auswahl == "Batchdatei":
        Bereinigung_File = choose_filename_with_path()
        logger.info(f"Batch-Datei ausgewählt: {Bereinigung_File}")
        if Bereinigung_File:
            Bereinigung_Liste, flagb = load_2column_list(Bereinigung_File)
            if not flagb:
                logger.error(f"Fehler beim Laden der Batch-Liste {Bereinigung_File}")
        else:
            logger.warning("Keine Batch-Datei ausgewählt")
            return False
    else:
        try:
            # Benutzer wählt alten und neuen Dateinamen aus
            NameAlt_full, NameNeu_full = choose_filename_and_replacementname(Basis_Pfad)
            NameAlt_Name = pathlib.Path(NameAlt_full).name
            NameNeu_Name = pathlib.Path(NameNeu_full).name
            Bereinigung_Liste = {NameAlt_Name: NameNeu_Name}
            flagb = True
            logger.info(f"Einzelfile-Modus: {NameAlt_Name} -> {NameNeu_Name}")
        except:
            logger.warning("Dateiauswahl abgebrochen")
            return False

    if flagb:
        for old_name, new_name in Bereinigung_Liste.items():
            logger.info(f"Verarbeite: {old_name} -> {new_name}")
            FileList = Search_Files(old_name, all_files_list)
            logger.info(f"  {len(FileList)} Dateien gefunden")

            # Entfernt die Erweiterung vom Ursprungsnamen und fügt .lnk hinzu
            NameLNKalt = pathlib.Path(old_name).stem + ".lnk"
            NameLNKneu = pathlib.Path(new_name).stem + ".lnk"

            FileLNKList = Search_Files(NameLNKalt, all_files_list)
            logger.info(f"  {len(FileLNKList)} LNK-Dateien gefunden")

            # Zeige Vorschau und warte auf Bestätigung
            logger.info("Zeige Vorschaufenster...")
            user_confirmed = show_preview_window(old_name, new_name, FileList, FileLNKList)

            if not user_confirmed:
                logger.warning(f"Benutzer hat Umbenennung abgebrochen: {old_name} -> {new_name}")
                continue  # Überspringe diesen Eintrag, gehe zum nächsten

            logger.info("Benutzer hat Umbenennung bestätigt")
            new_path_full = os.path.join(Basis_Pfad, new_name)

            # Führe Umbenennung durch
            Rename_Files_and_new_date(FileList, new_name)
            Target_in_LNK(FileLNKList, new_path_full)
            Rename_Files(FileLNKList, NameLNKneu)
            logger.info(f"Umbenennung abgeschlossen: {old_name} -> {new_name}")
        return True
    else:
        logger.warning("Keine gültigen Daten zum Umbenennen")
        return False




if __name__ == '__main__':
    # Load configuration from INI file
    config = StructuredConfigParser()
    config_file = pathlib.Path(__file__).parent / "FRICAL_v1.ini"

    if not config_file.exists():
        print(f"FEHLER: Konfigurationsdatei '{config_file}' nicht gefunden.")
        print(f"Bitte erstellen Sie die Datei 'FRICAL_v1.ini' im gleichen Verzeichnis wie das Skript.")
        print(f"Erforderlicher Inhalt:")
        print(f"[Files]")
        print(f"Such_Pfad = <Ihr Suchpfad>")
        print(f"Basis_Pfad = <Ihr Basispfad>")
        print(f"logfile = <Pfad zur Logdatei>")
        sys.exit(1)

    config.read(config_file, encoding='utf-8')

    # Read logfile path
    logfile_path = config.get_structured('Files', 'logfile', fallback=None)
    if logfile_path is None:
        print(f"FEHLER: 'logfile' nicht in Konfigurationsdatei '{config_file}' gefunden.")
        print(f"Bitte fügen Sie 'logfile = <Pfad>' im Abschnitt [Files] hinzu.")
        sys.exit(1)

    # Initialize logger
    logger = create_extended_logger(logfile_path, screen_output=True, script_name='FRICAL_v1')
    logger.info("=== FRICAL v1 gestartet ===")
    logger.info(f"Konfigurationsdatei: {config_file}")

    # Read Such_Pfad
    Such_Pfad_str = config.get_structured('Files', 'Such_Pfad', fallback=None)
    if Such_Pfad_str is None:
        logger.error(f"'Such_Pfad' nicht in Konfigurationsdatei gefunden.")
        logger.error(f"Bitte fügen Sie 'Such_Pfad = <Pfad>' im Abschnitt [Files] hinzu.")
        sys.exit(1)
    Such_Pfad = pathlib.Path(Such_Pfad_str)

    # Read Basis_Pfad
    Basis_Pfad_str = config.get_structured('Files', 'Basis_Pfad', fallback=None)
    if Basis_Pfad_str is None:
        logger.error(f"'Basis_Pfad' nicht in Konfigurationsdatei gefunden.")
        logger.error(f"Bitte fügen Sie 'Basis_Pfad = <Pfad>' im Abschnitt [Files] hinzu.")
        sys.exit(1)
    Basis_Pfad = pathlib.Path(Basis_Pfad_str)

    logger.info(f"Such_Pfad: {Such_Pfad}")
    logger.info(f"Basis_Pfad: {Basis_Pfad}")

    # Erstelle eine Liste aller Dateien im Suchpfad einmal zu Beginn
    logger.info("Erstelle Dateiliste im Suchpfad...")
    all_files_list = create_file_list(Such_Pfad)
    logger.info(f"Dateiliste erstellt: {len(all_files_list)} Dateien gefunden")

    # Hauptschleife für mehrere Umbenennungsvorgänge
    continue_processing = True
    while continue_processing:
        logger.info("=== Neuer Umbenennungsvorgang ===")

        # Führe Umbenennungsvorgang durch
        success = perform_rename_operation(all_files_list, Basis_Pfad, logger)

        if success:
            logger.info("Umbenennungsvorgang erfolgreich abgeschlossen")

        # Frage Benutzer, ob weitere Vorgänge durchgeführt werden sollen
        user_choice = show_selection_box("Weiteren Vorgang durchführen", "Programm beenden")

        if user_choice == "Programm beenden":
            continue_processing = False
            logger.info("Benutzer hat Programm beendet")
        else:
            logger.info("Benutzer startet weiteren Vorgang")
            logger.info("Verwende vorhandene Dateiliste weiter")

    logger.info("=== FRICAL v1 beendet ===")
    print("Programm beendet")
