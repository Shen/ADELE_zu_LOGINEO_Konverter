import os
import os.path
import sys
import pandas as pd
from datetime import datetime
from bs4 import BeautifulSoup
import codecs
import re
import traceback
import logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# If set DEBUG = True, Error-Msg will be visible
DEBUG = False
def debug():
    if DEBUG == True:
        traceback.print_last()


        

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
try:
    configfile = codecs.open(os.path.join(appdir, 'config.xml'), mode='r', encoding='utf-8')
    config = configfile.read()
    configfile.close()
except Exception:
    print("FEHLER!")
    print("Die config.xml wurde nicht gefunden.")
    print("Bitte prüfen Sie, ob sich die config.xml im selben Verzeichnis befindet, wie die Script-Datei.")
    debug()
    input("\nDrücken Sie eine beliebige Taste, um zu bestätigen und den Prozess zu beenden.")
    debug()
    sys.exit(1)

# load config values into variables
config_xmlsoup = BeautifulSoup(config, "html.parser")  # parse
config_txtfile = config_xmlsoup.find('txtfile').string  # import-file-name
config_txtfile_delimiter = config_xmlsoup.find('txtfile_delimiter').string  # import-file-delimiter
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
print("# VERSION: 1.8                                                                    #")
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
    input("\nDrücken Sie eine beliebige Taste, um zu bestätigen und den Prozess zu beenden.")
    sys.exit(1)

# import user-ADELE-Export-file
if config_txtfile.endswith('.xls') or config_txtfile.endswith('.xlsx'):
    df1 = pd.read_excel(config_txtfile, dtype=str)
elif config_txtfile.endswith('.txt'):
    df1 = pd.read_table(
    config_txtfile, sep=config_txtfile_delimiter, encoding='mbcs', engine='python')
else:
    print("FEHLER! FEHLER! FEHLER!")
    print("Die Datei (" + config_txtfile +
    "), die Sie in der config.xml eingetragen haben, hat kein zulässiges Dateiformat.")
    print("Zulässig sind Dateien mit den Endungen: .txt /.xls /.xlsx")
    input("\nDrücken Sie eine beliebige Taste, um zu bestätigen und den Prozess zu beenden.")
    sys.exit(1)

df1.fillna('', inplace=True)

# create dataframes
data = {"AdeleID": [], "IdentNr": [], "Nachname": [], "Vorname": [],
        "Typ": [], "Seminar": [], "Lehramt": [], "Jahrgang": [], "Kernseminar": [], 
        "Fachseminar_1": [], "Fachseminar_2": []}


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

def rmspaces(string):
    """ Replace spaces with underscores"""
    return re.sub('\s+', '_', string)

def read_column(source, column, datatype):
    """ Reads the content of a column """
    try:
        value = datatype(source[column])
    except:
        return False
    else:
        return value


def append_to_dataset(target, key, value, datatype):
    """Adds source to dataset"""
    target[key].append(datatype(value))


def add_adeleid(source, target):
    """
    Reads ADELE-ID and adds it to dataset
    column: AdeleID
    """
    value = read_column(source, 'Nr', str)
    append_to_dataset(target, 'AdeleID', value, int)


def add_identnr(source, target):
    """
    Reads IdentNr and adds it to dataset
    column: IdentNr
    """
    value = read_column(source, 'Identnummer', str)
    if len(value) > 9:
        if len(value) == 10:
            append_to_dataset(target, 'IdentNr', "0" + value, str)
        else:
            append_to_dataset(target, 'IdentNr', value, str)
    else:
        append_to_dataset(target, 'IdentNr', 'IdentNr fehlt', str)


def add_nachname(source, target):
    """
    Reads lastname and adds it to dataset
    column: Nachname
    """
    nachname = read_column(source, 'Name', str)
    familienname = read_column(source, 'Familienname', str)
    namensvorsatz = read_column(source, 'Namensvorsatz', str)
    if nachname != False:
        append_to_dataset(target, 'Nachname', nachname, str)
    elif familienname != False and namensvorsatz != False:
        append_to_dataset(target, 'Nachname', namensvorsatz + ' ' + familienname, str)
    elif familienname != False and namensvorsatz == False:
        append_to_dataset(target, 'Nachname', familienname, str)
    else:
        append_to_dataset(target, 'Nachname', "FEHLER", str)


def add_vorname(source, target):
    """
    Reads surname and adds it to dataset
    column: Vorname
    """
    value = read_column(source, 'Vorname', str)
    append_to_dataset(target, 'Vorname', value, str)    


def add_status(source, target):
    """
    Adds status (LAA/SAB) to dataset
    column: Typ
    """
    value = source
    append_to_dataset(target, 'Typ', value, str)     


def add_seminar(source, target):
    """
    Reads Seminar and adds Seminar_Lehramt to dataset
    column: Seminar
    """
    lehramt = read_column(source, 'Lehramt', str)
    lehramt1 = read_column(source, 'Lehramt1', str)
    seminar = read_column(source, 'Seminar', str)
    
    if lehramt != False and lehramt !="":
        if int(float(lehramt)) in lehraemter and lehramt !="":
           append_to_dataset(target, 'Seminar', "Seminar_" + lehraemter[int(float(lehramt))], str)
        else:
            append_to_dataset(target, 'Seminar', "Seminar_???", str)
    elif lehramt1 != False and lehramt1 !="":
        if int(float(lehramt1)) in lehraemter:
           append_to_dataset(target, 'Seminar', "Seminar_" + lehraemter[int(float(lehramt1))], str)
        else:
            append_to_dataset(target, 'Seminar', "Seminar_???", str)
    elif seminar != False and seminar !="":
        if int(float(seminar)) in seminare:
             append_to_dataset(target, 'Seminar', 'Seminar_' + seminare[int(float(seminar))], str)
        else:
             append_to_dataset(target, 'Seminar', "Seminar_???", str)
    else:
        append_to_dataset(target, 'Seminar', "Seminar_???", str)


def add_lehramt(source, target):
    """
    Reads Lehramt and adds LAA_Lehramt it to dataset
    column: Lehramt
    """
    lehramt = read_column(source, 'Lehramt', str)
    lehramt1 = read_column(source, 'Lehramt1', str)
    seminar = read_column(source, 'Seminar', str)

    if lehramt != False and lehramt !="":
        if int(float(lehramt)) in lehraemter:
           append_to_dataset(target, 'Lehramt', "LAA_" + lehraemter[int(float(lehramt))], str)
        else:
            append_to_dataset(target, 'Lehramt', "LAA_???", str)
    elif lehramt1 != False and lehramt1 !="":
        if int(float(lehramt1)) in lehraemter:
           append_to_dataset(target, 'Lehramt', "LAA_" + lehraemter[int(float(lehramt1))], str)
        else:
            append_to_dataset(target, 'Lehramt', "LAA_???", str)
    elif seminar != False and seminar != "":
        if int(float(seminar)) in seminare:
             append_to_dataset(target, 'Lehramt', 'LAA_' + seminare[int(float(seminar))], str)
        else:
             append_to_dataset(target, 'Lehramt', "LAA_???", str)
    else:
        append_to_dataset(target, 'Lehramt', "LAA_???", str)


def split_year_short(string):
    """Splits year from date in format YYYY-MM-DD"""
    return string[-4:]

def split_year_long_xls(string):
    """Splits year from date in format YYYY-MM-DD 00:00:00 in xls/xlsx-files"""
    return string[-19:-15]

def split_year_long_txt(string):
    """Splits year from date in format YYYY-MM-DD 00:00:00 in txt-files"""
    return string[-13:-9]

def split_month_short(string):
    """Splits month from date in format YYYY-MM-DD"""
    return str(string[-7:-5])

def split_month_long_xls(string):
    """Splits month from date in format YYYY-MM-DD 00:00:00 in xls/xlsx-files"""
    return string[-14:-12]

def split_month_long_txt(string):
    """Splits month from date in format YYYY-MM-DD 00:00:00 in txt-files"""
    return string[-16:-14]

def add_jahrgang(source, target):
    """
    Reads Jahrgang and adds it LAA_Seminar_Jahrgang dataset
    column: Jahrgang
    """
    lehramt = read_column(source, 'Lehramt', str)
    lehramt1 = read_column(source, 'Lehramt1', str)
    seminar = read_column(source, 'Seminar', str)   
    vd1_von = read_column(source, 'VD1_von', str)
    
    jahrgang = "LAA_"
    if lehramt != False and lehramt != '':
        if int(float(lehramt)) in lehraemter:
            jahrgang += str(lehraemter[int(float(lehramt))]) + "_"
        else:
            jahrgang += "???_"
    elif lehramt1 != False and lehramt1 !="":
        if int(float(lehramt1)) in lehraemter:
            jahrgang += str(lehraemter[int(float(lehramt1))]) + "_"
        else:
            jahrgang += "???_"
    elif seminar != False and seminar !="":
        if int(float(seminar)) in seminare:
            jahrgang += str(seminare[int(float(seminar))]) + "_"
        else:
            jahrgang += "???_"
    else:
        jahrgang += "???_"
    
    if vd1_von != False:
        if len(vd1_von) == 10:
            start_date_short = split_year_short(vd1_von) + '-' + split_month_short(vd1_von)
            jahrgang += start_date_short
        elif len(source['VD1_von']) == 19:
            if config_txtfile.endswith('.xls') or config_txtfile.endswith('.xlsx'):
                start_date_long_xls = split_year_long_xls(vd1_von) + '-' + split_month_long_xls(vd1_von)
                jahrgang +=  start_date_long_xls
            elif config_txtfile.endswith('.txt'):
                start_date_long_txt = split_year_long_txt(vd1_von) + '-' + split_month_long_txt(vd1_von)
                jahrgang +=  start_date_long_txt
    else:
        jahrgang +=  "???"
       
    append_to_dataset(target, 'Jahrgang', jahrgang, str)


def add_kernseminar(source, target):
    """
    Reads Hsem/Hsem_Leiter and adds it to dataset
    column: Kernseminar
    """
    if 'HSem' in source and source['HSem'] != "" and 'HSem_Leiter' in source and ['HSem_Leiter'] != "":
        target['Kernseminar'].append('Seminar_'+rmspaces(str(source['HSem_Leiter']))+'_'+rmspaces(str(source['HSem'])))
    else:
        target['Kernseminar'].append('')

def add_fachseminar_1(source, target):
    """
    Reads FSem1/FSem1_Leiter and adds it to dataset
    column: Fachseminar_1
    """
    if 'FSem1' in source and source['FSem1'] != "" and 'FSem1_Leiter' in source and ['FSem1_Leiter'] != "":
        target['Fachseminar_1'].append('Seminar_'+rmspaces(str(source['FSem1_Leiter']))+'_'+rmspaces(str(source['FSem1'])))
    else:
        target['Fachseminar_1'].append('')

def add_fachseminar_2(source, target):
    """
    Reads FSem2/FSem2_Leiter and adds it to dataset
    column: Fachseminar_2
    """
    if 'FSem2' in source and source['FSem2'] != "" and 'FSem2_Leiter' in source and ['FSem2_Leiter'] != "":
        target['Fachseminar_2'].append('Seminar_'+rmspaces(str(source['FSem2_Leiter']))+'_'+rmspaces(str(source['FSem2'])))
    else:
        target['Fachseminar_2'].append('')

# Fill dataframes
for i, j in df1.iterrows():
    # adding a new row (be careful to ensure every column gets another value)
    if config_primary_key == 'IdentNr':
        if (df1.iloc[i]['Identnummer']) != '' and len(str(df1.iloc[i]['Identnummer'])) > 9:
            try:
                add_identnr(df1.iloc[i], data)
                add_adeleid(df1.iloc[i], data)
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
                    add_fachseminar_1(df1.iloc[i], data)
                    add_fachseminar_2(df1.iloc[i], data)
            except Exception:
                print("")
                print("\nFEHLER - FEHLER - FEHLER. (#9.1)")
                print("Bei der Ausführung ist etwas schiefgelaufen.")
                print("Überprüfen Sie die Einstellungen (config.xml) und die Quell-Datei.")
                input("\nDrücken Sie eine beliebige Taste, um das Programm zu beenden.")
                debug()
                sys.exit(1)                                           

        else:
            try:
                add_identnr(df1.iloc[i], datafail)
                add_adeleid(df1.iloc[i], datafail)
                add_nachname(df1.iloc[i], datafail)
                add_vorname(df1.iloc[i], datafail)
                add_status("LAA", datafail)
                add_lehramt(df1.iloc[i], datafail)
            except Exception:
                print("")
                print("\nFEHLER - FEHLER - FEHLER. (#9.2)")
                print("Bei der Ausführung ist etwas schiefgelaufen.")
                print("Überprüfen Sie die Einstellungen (config.xml) und die Quell-Datei.")
                input("\nDrücken Sie eine beliebige Taste, um das Programm zu beenden.")
                debug()
                sys.exit(1)              

    elif config_primary_key == 'AdeleID':
        if 'Nr' in df1.iloc[i] and df1.iloc[i]['Nr'] != "":
            try:
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
                    add_fachseminar_1(df1.iloc[i], data)
                    add_fachseminar_2(df1.iloc[i], data)
            except Exception:
                print("")
                print("\nFEHLER - FEHLER - FEHLER. (#9.3)")
                print("Bei der Ausführung ist etwas schiefgelaufen.")
                print("Überprüfen Sie die Einstellungen (config.xml) und die Quell-Datei.")
                input("\nDrücken Sie eine beliebige Taste, um das Programm zu beenden.")
                debug()
                sys.exit(1)
        else:
            try:
                add_adeleid(df1.iloc[i], datafail)
                add_identnr(df1.iloc[i], datafail)
                add_nachname(df1.iloc[i], datafail)
                add_vorname(df1.iloc[i], datafail)
                add_status("LAA", datafail)
                add_lehramt(df1.iloc[i], datafail)
            except Exception:
                print("")
                print("\nFEHLER - FEHLER - FEHLER. (#9.4)")
                print("Bei der Ausführung ist etwas schiefgelaufen.")
                print("Überprüfen Sie die Einstellungen (config.xml) und die Quell-Datei.")
                input("\nDrücken Sie eine beliebige Taste, um das Programm zu beenden.")
                debug()
                sys.exit(1)
    else:
        print("")
        print("\nFEHLER - FEHLER - FEHLER. (#9.5)")
        print("\nDie Angaben zum Primary Key in der config.xml sind falsch.")
        print("Bitte überprüfen Sie die Einstellungen.")
        input("\nDrücken Sie eine beliebige Taste, um das Programm zu beenden.")
        sys.exit(1)

## Test-Block
#print(data)
#print(f'AdeleID {len(data["AdeleID"])}')
#print(f'IdentNr {len(data["IdentNr"])}')
#print(f'Nachname {len(data["Nachname"])}')
#print(f'Vorname {len(data["Vorname"])}')
#print(f'Typ {len(data["Typ"])}')
#print(f'Seminar {len(data["Seminar"])}')
#print(f'Lehramt {len(data["Lehramt"])}')
#print(f'Jahrgang {len(data["Jahrgang"])}')
#print(f'Kernseminar {len(data["Kernseminar"])}')
#print(f'Fachseminar_1 {len(data["Fachseminar_1"])}')
#print(f'Fachseminar_2 {len(data["Fachseminar_2"])}')
#print(f'Data[Jahrgang] {data["Jahrgang"]}')

# safe results in new dataframes
df2 = pd.DataFrame(data, columns=[
                   'AdeleID', 'IdentNr', 'Nachname', 'Vorname', 'Typ', 'Seminar', 'Lehramt', 'Jahrgang', 'Kernseminar', 'Fachseminar_1', 'Fachseminar_2'])
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
    print("Hier eine Übersicht der finalen Tabellen-Struktur und der anzulegenden Nutzer:")
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
