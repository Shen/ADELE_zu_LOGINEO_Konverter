import os
import os.path
import sys
import pandas as pd
from datetime import datetime
from bs4 import BeautifulSoup
import codecs

# This tool imports ADELE-txt-exports and refactors it for LOGINEO NRW Import.

# Created by Johannes Schirge
# Mail: johannes.schirge@zfsl-bielefeld.nrw.schule

# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# any later version.

# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.

# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

print("")
print("Created by Johannes Schirge")
print("ZfsL Bielefeld")
print("E-Mail: johannes.schirge@zfsl-bielefeld.nrw.schule")
print("")
print("This program comes with ABSOLUTELY NO WARRANTY")
print("This is free software, and you are welcome to redistribute it under certain conditions.")
print("For details look into LICENSE file (GNU GPLv3).")
print("")

# determine if running in a build package (frozen) or from seperate python script
frozen = 'not'
if getattr(sys, 'frozen', False):
    # we are running in a bundle
    appdir = os.path.dirname(os.path.abspath(sys.executable))
    # print("Executable is in frozen state, appdir set to: " + appdir) # for debug
else:
    # we are running in a normal Python environment
    appdir = os.path.dirname(os.path.abspath(__file__))
    # print("Executable is run in normal Python environment, appdir set to: " + appdir) # for debug

# read config from xml file
configfile = codecs.open(os.path.join(
    appdir, 'config.xml'), mode='r', encoding='utf-8')
config = configfile.read()
configfile.close()

# load config values into variables
config_xmlsoup = BeautifulSoup(config, "html.parser")  # parse
config_txtfile = config_xmlsoup.find('txtfile').string  # import-file-name
config_txtfile_delimiter = config_xmlsoup.find(
    'txtfile_delimiter').string  # import-file-delimiter
print(config_txtfile_delimiter)
# set if AdeleID or IdentNr is primary key in LOGINEO
config_primary_key = config_xmlsoup.find('primary_key').string
config_gruppe_laa_lehramt = config_xmlsoup.find(
    'gruppe_laa_lehramt').string  # group LAA_LEHRAMT
config_gruppe_laa_lehramt_jg = config_xmlsoup.find(
    'gruppe_laa_lehramt_jg').string  # group LAA_LEHRAMT_JAHRGANG
config_gruppe_laa_seminare = config_xmlsoup.find(
    'gruppe_laa_seminare').string  # groups Seminare

# logineo Info-Text
print("")
print("###################################################################################")
print("# Inoffizielles LAA-ADELE-Export zu LOGINEO NRW-Import-Tool für ZfsL-Instanzen    #")
print("# VERSION: 1.7                                                                   #")
print("# Dieses Tool erstellt aus einem unveränderten LAA-ADELE-.txt/xlsx-Export         #")
print("# eine Exceldatei (.xlsx), für den LOGINEO NRW-LAA-Nutzerdatenimport.             #")
print("#                                                                                 #")
print("# Es werden automatische Gruppen erzeugt: LAA, LAA_Lehramt (z. B. LAA_GyGe)       #")
print("# Fehlerhafte Zeilen der Datei oder Zeilen, die keine Ident-Nr. enthalten,        #")
print("# werden in einer gesonderten Excel-Datei ausgegeben.                             #")
print("###################################################################################")
print("")
print("")
print("Wenn Sie sicher sind, dass Ihre Einstellungen in der config.xml korrekt sind,")
print("drücken Sie eine beliebige Taste, um fortzufahren.")
print("ACHTUNG: Dieses Tool ist ausschließlich für den Import von LAA-Daten (Referendare)")
print("vorgesehen, nicht für Fachleitungen oder anderes Personal!")
input("Andernfalls brechen Sie den Prozess mit [STRG + C] ab.")
print("")
print("Ihre LAA-ADELE-Datei wird nun eingelesen.")
print("")

# adds date and time as string to variable
now = datetime.now()
dt_string = now.strftime("%Y-%m-%d_%H-%M-%S")

# set/create output-directory and output-names
output_dir = 'output'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# check if user-ADELE-Export-file exists
if not os.path.isfile(config_txtfile):
    print("FEHLER!")
    print("Die txt-Datei (" + config_txtfile + "), die Sie in der config.xml eingetragen haben, existiert nicht. Bitte speichern Sie die Datei '" +
          config_txtfile + "' im Hauptverzeichnis des Scripts oder bearbeiten Sie die config.xml")
    input("Drücken Sie eine beliebige Taste, um zu bestätigen und den Prozess zu beenden.")
    sys.exit(1)

# import user-ADELE-Export-file
if config_txtfile.endswith('.xls') or config_txtfile.endswith('.xlsx'):
    df1 = pd.read_excel(config_txtfile, dtype=str)
elif config_txtfile.endswith('.txt'):
    df1 = pd.read_table(
        config_txtfile, sep=config_txtfile_delimiter, encoding='mbcs')
else:
    print("FEHLER!")
    print("Die Datei (" + config_txtfile +
          "), die Sie in der config.xml eingetragen haben, konnte nicht eingelesen werden.")
    input("Drücken Sie eine beliebige Taste, um zu bestätigen und den Prozess zu beenden.")
    sys.exit(1)

df1.fillna('', inplace=True)

# create dataframes
data = {"AdeleID": [], "IdentNr": [], "Nachname": [], "Vorname": [],
        "Typ": [], "Seminar": [], "Lehramt": [], "Jahrgang": [], "Kernseminar": []}


datafail = {"AdeleID": [], "IdentNr": [], "Nachname": [],
            "Vorname": [],  "Typ": [],  "Lehramt": []}

# dictionary for number-teaching type assignment (df1.iloc[i]['Lehramt'])
lehraemter = {}
lehraemter[0] = 'kein Eintrag'  # Primarstufe
lehraemter[1] = 'unbekannt'
lehraemter[2] = 'unbekannt'
lehraemter[3] = 'unbekannt'
lehraemter[4] = 'G'  # an Grundschulen
lehraemter[5] = 'unbekannt'
lehraemter[6] = 'unbekannt'
lehraemter[7] = 'unbekannt'
lehraemter[8] = 'SF'  # für sonderpädagogische Förderung
lehraemter[9] = 'SF'  # Sonderpädagogik
lehraemter[10] = 'unbekannt'
lehraemter[11] = 'unbekannt'
lehraemter[12] = 'unbekannt'
lehraemter[13] = 'unbekannt'
lehraemter[14] = 'SF'  # Sonderpädagogik
lehraemter[15] = 'G'  # an Grund-, Haupt-, Real- und Gesamtschulen
lehraemter[16] = 'HRSGe'  # an Grund-, Haupt-, Real- und Gesamtschulen
lehraemter[17] = 'HRSGe'  # an Haupt-, Real- und Gesamtschulen
lehraemter[18] = 'HRSGe'  # an Haupt-, Real-, Sekundar- und Gesamtschulen
lehraemter[19] = 'unbekannt'
lehraemter[20] = 'unbekannt'  # Sekundarstufe I
lehraemter[21] = 'unbekannt'
lehraemter[22] = 'unbekannt'
lehraemter[23] = 'unbekannt'
lehraemter[24] = 'GyGe'  # Sekundarstufe II und Sekundarstufe I
lehraemter[25] = 'unbekannt'
lehraemter[26] = 'unbekannt'
lehraemter[27] = 'GyGe'  # an Gymnasien und Gesamtschulen
lehraemter[28] = 'unbekannt'
lehraemter[29] = 'unbekannt'  # Sekundarstufe II
lehraemter[30] = 'unbekannt'
lehraemter[31] = 'unbekannt'
lehraemter[32] = 'unbekannt'  # Sekundarstufe II mit berufl. Fachrichtung
lehraemter[33] = 'unbekannt'
lehraemter[34] = 'unbekannt'
lehraemter[35] = 'BK'  # an Berufskollegs

# dictionary for instituteID-teaching-type assignment (df1.iloc[i]['Seminar'])
seminare = {}
# ZfsL Bielefeld
seminare[510749] = 'G'
seminare[510750] = 'HRSGe'
seminare[510762] = 'SF'
seminare[510774] = 'GyGe'
seminare[510786] = 'BK'


# Functions

def add_adeleid(source, target):
    """
    Reads ADELE-ID and adds it to dataset
    column: AdeleID
    """
    if 'Nr' in source and source['Nr'] != "":
        target['AdeleID'].append(str(source['Nr']))
    else:
        target['AdeleID'].append('AdeleID fehlt')


def add_identnr(source, target):
    """
    Reads IdentNr and adds it to dataset
    column: IdentNr
    """
    if 'Identnummer' in source and (source['Identnummer']) != '' and len(str(source['Identnummer'])) > 9:
        if len(source['Identnummer']) == 10:
            target['IdentNr'].append("0" + str(source['Identnummer']))
        else:
            target['IdentNr'].append(str(source['Identnummer']))
    else:
        target['IdentNr'].append('IdentNr fehlt')


def add_nachname(source, target):
    """
    Reads lastname and adds it to dataset
    column: Nachname
    """
    if 'Name' in source:
        target['Nachname'].append(str(source['Name']))
    elif 'Familienname' in source and source['Familienname'] != '' and 'Namensvorsatz' in source and source['Namensvorsatz'] != '':
        target['Nachname'].append(
            str(source['Namensvorsatz']) + ' ' + str(source['Familienname']))
    elif 'Familienname' in source and source['Familienname'] != '':
        target['Nachname'].append(source['Familienname'])
    else:
        target['Nachname'].append('FEHLER')


def add_vorname(source, target):
    """
    Reads surname and adds it to dataset
    column: Vorname
    """
    if (source['Vorname']) != '':
        target['Vorname'].append(source['Vorname'])
    else:
        target['Vorname'].append('FEHLER')


def add_status(source, target):
    """
    Adds status (LAA/SAB) to dataset
    column: Typ
    """
    target['Typ'].append(source)


def add_seminar(source, target):
    """
    Reads Seminar and adds Seminar_Lehramt to dataset
    column: Seminar
    """
    if 'Lehramt' in source and source['Lehramt'] != "":
        if source['Lehramt'] in lehraemter:
            target['Seminar'].append(
                'Seminar_'+str(lehraemter[source['Lehramt']]))
        else:
            target['Seminar'].append('')
    elif 'Lehramt1' in source and source['Lehramt1'] != "":
        if source['Lehramt1'] in lehraemter:
            target['Seminar'].append(
                'Seminar_'+str(lehraemter[source['Lehramt1']]))
        else:
            target['Seminar'].append('')
    elif 'Seminar' in source and source['Seminar'] != "":
        if source['Seminar'] in seminare:
            target['Seminar'].append(
                'Seminar_'+str(seminare[source['Seminar']]))
        else:
            target['Seminar'].append('')
    else:
        target['Seminar'].append('')


def add_lehramt(source, target):
    """
    Reads Lehramt and adds LAA_Lehramt it to dataset
    column: Lehramt
    """
    if 'Lehramt' in source and source['Lehramt'] != "":
        if source['Lehramt'] in lehraemter:
            target['Lehramt'].append('LAA_'+str(lehraemter[source['Lehramt']]))
        else:
            target['Lehramt'].append('')
    elif 'Lehramt1' in source and source['Lehramt1'] != "":
        if source['Lehramt1'] in lehraemter:
            target['Lehramt'].append(
                'LAA_'+str(lehraemter[source['Lehramt1']]))
        else:
            target['Lehramt'].append('')
    elif 'Seminar' in source and source['Seminar'] != "":
        if source['Seminar'] in seminare:
            target['Lehramt'].append('LAA_'+str(seminare[source['Seminar']]))
        else:
            target['Lehramt'].append('')
    else:
        target['Lehramt'].append('')


def add_jahrgang(source, target):
    """
    Reads Jahrgang and adds it LAA_Seminar_Jahrgang dataset
    column: Jahrgang
    """
    if 'Lehramt' in source and source['Lehramt'] != "":
        if source['Lehramt'] in lehraemter and source['VD1_von'] != '':
            if len(str(source['VD1_von'])) == 10:
                target['Jahrgang'].append('LAA_'+str(lehraemter[source['Lehramt']])+'_'+(
                    str(source['VD1_von'])[-4:])+'-'+(str(source['VD1_von'])[-7:-5]))
            elif len(str(source['VD1_von'])) == 19:
                if config_txtfile.endswith('.xls') or config_txtfile.endswith('.xlsx'):
                    target['Jahrgang'].append('LAA_'+str(lehraemter[source['Lehramt']])+'_'+(
                        str(source['VD1_von'])[-19:-15])+'-'+(str(source['VD1_von'])[-14:-12]))
                elif config_txtfile.endswith('.txt'):
                    target['Jahrgang'].append('LAA_'+str(lehraemter[source['Lehramt']])+'_'+(
                        str(source['VD1_von'])[-13:-9])+'-'+(str(source['VD1_von'])[-16:-14]))
                else:
                    target['Jahrgang'].append('')
            else:
                target['Jahrgang'].append('')
    elif 'Lehramt1' in source and source['Lehramt1'] != "":
        if source['Lehramt1'] in lehraemter and source['VD1_von'] != '':
            if len(str(source['VD1_von'])) == 10:
                target['Jahrgang'].append('LAA_'+str(lehraemter[source['Lehramt1']])+'_'+(
                    str(source['VD1_von'])[-4:])+'-'+(str(source['VD1_von'])[-7:-5]))
                if len(str(source['VD1_von'])) == 19:
                    if config_txtfile.endswith('.xls') or config_txtfile.endswith('.xlsx'):
                        target['Jahrgang'].append('LAA_'+str(lehraemter[source['Lehramt1']])+'_'+(
                            str(source['VD1_von'])[-19:-15])+'-'+(str(source['VD1_von'])[-14:-12]))
                    elif config_txtfile.endswith('.txt'):
                        target['Jahrgang'].append('LAA_'+str(lehraemter[source['Lehramt1']])+'_'+(
                            str(source['VD1_von'])[-13:-9])+'-'+(str(source['VD1_von'])[-16:-14]))
                    else:
                        target['Jahrgang'].append('')
                else:
                    target['Jahrgang'].append('')
    elif 'Seminar' in source and source['Seminar'] != "":
        if source['Seminar'] in seminare and source['VD1_von'] != '':
            if len(str(source['VD1_von'])) == 10:
                target['Jahrgang'].append('LAA_'+str(seminare[source['Seminar']])+'_'+(
                    str(source['VD1_von'])[-4:])+'-'+(str(source['VD1_von'])[-7:-5]))
            if len(str(source['VD1_von'])) == 19:
                if config_txtfile.endswith('.xls') or config_txtfile.endswith('.xlsx'):
                    target['Jahrgang'].append('LAA_'+str(seminare[source['Seminar']])+'_'+(
                        str(source['VD1_von'])[-19:-15])+'-'+(str(source['VD1_von'])[-14:-12]))
                elif config_txtfile.endswith('.txt'):
                    target['Jahrgang'].append('LAA_'+str(seminare[source['Seminar']])+'_'+(
                        str(source['VD1_von'])[-13:-9])+'-'+(str(source['VD1_von'])[-16:-14]))
                else:
                    target['Jahrgang'].append('')
            else:
                target['Jahrgang'].append('')
    else:
        target['Jahrgang'].append('')

def add_kernseminar(source, target):
    """
    Reads Hsem/Hsem_Leiter and adds it to dataset
    column: Kernseminar
    """
    if 'HSem' in source and source['HSem'] != "" and 'HSem_Leiter' in source and ['HSem_Leiter'] != "":
        target['Kernseminar'].append('Seminar_'+str(source['HSem'])+'_'+str(source['HSem_Leiter']))
    else:
        target['Kernseminar'].append('')

# Fill dataframes
for i, j in df1.iterrows():
    # adding a new row (be careful to ensure every column gets another value)
    if config_primary_key == 'IdentNr':
        if (df1.iloc[i]['Identnummer']) != '' and len(str(df1.iloc[i]['Identnummer'])) > 9:
            add_identnr(df1.iloc[i], data)
            add_adeleid(df1.iloc[i], data)
            add_nachname(df1.iloc[i], data)
            add_vorname(df1.iloc[i], data)
            add_status("LAA", data)
            add_seminar(df1.iloc[i], data)
            if config_gruppe_laa_lehramt == 'ja':
                add_seminar(df1.iloc[i], data)
            if config_gruppe_laa_lehramt == 'ja':
                add_lehramt(df1.iloc[i], data)
            if config_gruppe_laa_lehramt_jg == 'ja':
                add_jahrgang(df1.iloc[i], data)
            if config_gruppe_laa_seminare == 'ja':
                add_kernseminar(df1.iloc[i], data)                

        else:
            add_identnr(df1.iloc[i], datafail)
            add_adeleid(df1.iloc[i], datafail)
            add_nachname(df1.iloc[i], datafail)
            add_vorname(df1.iloc[i], datafail)
            add_status("LAA", datafail)
            add_lehramt(df1.iloc[i], datafail)

    elif config_primary_key == 'AdeleID':
        if 'Nr' in df1.iloc[i] and df1.iloc[i]['Nr'] != "":
            add_adeleid(df1.iloc[i], data)
            add_identnr(df1.iloc[i], data)
            add_nachname(df1.iloc[i], data)
            add_vorname(df1.iloc[i], data)
            add_status("LAA", data)
            if config_gruppe_laa_lehramt == 'ja':
                add_seminar(df1.iloc[i], data)
            if config_gruppe_laa_lehramt == 'ja':
                add_lehramt(df1.iloc[i], data)
            if config_gruppe_laa_lehramt_jg == 'ja':
                add_jahrgang(df1.iloc[i], data)
            if config_gruppe_laa_seminare == 'ja':
                add_kernseminar(df1.iloc[i], data)
        else:
            add_adeleid(df1.iloc[i], datafail)
            add_identnr(df1.iloc[i], datafail)
            add_nachname(df1.iloc[i], datafail)
            add_vorname(df1.iloc[i], datafail)
            add_status("LAA", datafail)
            add_lehramt(df1.iloc[i], datafail)
    else:
        print("")
        print("\nFEHLER - FEHLER - FEHLER.")
        print("\nDie Angaben zum Primary Key in der config.xml sind falsch.")
        print("Bitte überprüfen Sie die Einstellungen.")
        input("\nDrücken Sie eine beliebige Taste, um das Programm zu beenden.")
        sys.exit(1)
# print(data)

# safe results in new dataframes
df2 = pd.DataFrame(data, columns=[
                   'AdeleID', 'IdentNr', 'Nachname', 'Vorname', 'Typ', 'Seminar', 'Lehramt', 'Jahrgang', 'Kernseminar'])
df3 = pd.DataFrame(datafail, columns=[
                   'AdeleID', 'IdentNr', 'Nachname', 'Vorname', 'Typ', 'Lehramt'])

# display error-results for logineo-users
if not df3.empty:
    print(df3)
    output_filename_error = dt_string+'_Referendare_FEHLER.xlsx'
    output_filepath_error = os.path.join(output_dir, output_filename_error)
    df3.to_excel(output_filepath_error, 'Referendare-FEHLER')
    print("")
    print("WARNUNG - WARNUNG - WARNUNG")
    print("\nBei einigen importierten Zeilen sind Probleme/Fehler aufgetreten (siehe oben).")
    print("Prüfen Sie ggf. die Primärquelle (ADELE), ob diese Fehler in ADELE behoben werden können.")
    print("Erstellen Sie nach der Fehlerbehebung eine neue Export-Datei mit ADELE.")
    print("Sie finden eine Excel-Datei mit der Liste der Fehler im Output-Ordner.")
    print("")
    print("")
    input("Wenn Sie die Fehler ignorieren möchten, drücken Sie eine beliebige Taste, um fortzufahren. Wenn nicht, drücken Sie Strg+C, um abzubrechen.")

# display expected results for logineo-users
if not df2.empty:
    print("")
    print("Hier eine Übersicht der Tabellen-Struktur und der anzulegenden Nutzer:")
    print("")
    print(df2)

# ask user to check values and continue
    print("\nÜberprüfen Sie, ob die Daten für die zu generierenden Excel-Dateien korrekt sind.")
    input("Wenn alles gut aussieht, drücken Sie eine beliebige Taste, um fortzufahren. Wenn nicht, drücken Sie Strg+C, um abzubrechen.")
    output_filename = dt_string+'_referendare.xlsx'
    output_filepath = os.path.join(output_dir, output_filename)
    df2.to_excel(output_filepath, 'Referendare')
    print("")
    print("\nEs wurde erfolgreich eine Datei mit den Referendaren im Output-Ordner angelegt.")
    print("Sie können diese Datei nun in der Nutzerverwaltung LOGINEO NRWs importieren.")
    input("\nDrücken Sie eine beliebige Taste, um das Programm zu beenden.")
    exit
else:
    print("\nIhre Tabelle enthält keine gültigen Werte. Der Prozess wird abgebrochen.")
    input("\nDrücken Sie eine beliebige Taste, um das Programm zu beenden.")
    exit
