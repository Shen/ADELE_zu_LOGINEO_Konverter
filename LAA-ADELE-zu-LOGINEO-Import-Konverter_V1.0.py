import os
import os.path
import sys
import pandas as pd
from datetime import datetime
from bs4 import BeautifulSoup
import html
import string
import requests
import codecs

# This tool imports ADELE-txt-exports and refactors it for LOGINEO NRW Import.

# Created by Johannes Schirge
# Mail: johannes.schirge@zfs-bielefeld.nrw.schule

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
print("Github: https://github.com/Shen/ADELE_to_LOGINEO")
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
  ## print("Executable is in frozen state, appdir set to: " + appdir) # for debug
else:
  # we are running in a normal Python environment
  appdir = os.path.dirname(os.path.abspath(__file__))
  ## print("Executable is run in normal Python environment, appdir set to: " + appdir) # for debug

# read config from xml file
configfile = codecs.open(os.path.join(appdir,'config.xml'),mode='r', encoding='utf-8')
config = configfile.read()
configfile.close()

# load config values into variables
config_xmlsoup = BeautifulSoup(config, "html.parser") # parse
config_txtfile = config_xmlsoup.find('txtfile').string


# logineo Info-Text
print("")
print("###################################################################################")
print("# Inoffizielles LAA-ADELE-Export zu LOGINEO NRW-Import-Tool für ZfsL-Instanzen    #")
print("# VERSION: 1.0                                                                    #")
print("# Dieses Tool erstellt aus einem unveränderten LAA-ADELE-.txt-Export              #")
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
print("Ihre LAA-ADELE-.txt-Datei wird nun eingelesen.")
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
      print("Die txt-Datei (" + config_txtfile + "), die Sie in der config.xml eingetragen haben, existiert nicht. Bitte speichern Sie die Datei '" + config_txtfile + "' im Hauptverzeichnis des Scripts oder bearbeiten Sie die config.xml")
      input("Drücken Sie eine beliebige Taste, um zu bestätigen und den Prozess zu beenden.")  
      sys.exit(1)

# import user-ADELE-Export-file
df1 = pd.read_table(config_txtfile, sep='\t', encoding='mbcs')
df1.fillna('', inplace=True)

# create dataframes
data = {"Quell-Id":[], "Nachname":[], "Vorname":[],  "Typ":[],  "Lehramt":[]}
datafail = {"Quell-Id":[], "Nachname":[], "Vorname":[],  "Typ":[],  "Lehramt":[]}

# dictionary for number-teaching type assignment (df1.iloc[i]['Lehramt'])
lehraemter = {}
lehraemter[0] = 'kein Eintrag'
lehraemter[1] = 'unbekannt'
lehraemter[2] = 'unbekannt'
lehraemter[3] = 'unbekannt'
lehraemter[4] = 'G'
lehraemter[5] = 'unbekannt'
lehraemter[6] = 'unbekannt'
lehraemter[7] = 'unbekannt'
lehraemter[8] = 'SF'
lehraemter[9] = 'unbekannt'
lehraemter[10] = 'unbekannt'
lehraemter[11] = 'unbekannt'
lehraemter[12] = 'unbekannt'
lehraemter[13] = 'unbekannt'
lehraemter[14] = 'SF'
lehraemter[15] = 'G'
lehraemter[16] = 'HRSGe'
lehraemter[17] = 'unbekannt'
lehraemter[18] = 'HRSGe'
lehraemter[19] = 'unbekannt'
lehraemter[20] = 'unbekannt'
lehraemter[21] = 'unbekannt'
lehraemter[22] = 'unbekannt'
lehraemter[23] = 'unbekannt'
lehraemter[24] = 'unbekannt'
lehraemter[25] = 'unbekannt'
lehraemter[26] = 'unbekannt'
lehraemter[27] = 'GyGe'
lehraemter[28] = 'unbekannt'
lehraemter[29] = 'unbekannt'
lehraemter[30] = 'unbekannt'
lehraemter[31] = 'unbekannt'
lehraemter[32] = 'unbekannt'
lehraemter[33] = 'unbekannt'
lehraemter[34] = 'unbekannt'
lehraemter[35] = 'BK'


for i, j in df1.iterrows():
    # adding a new row (be careful to ensure every column gets another value)
    if (df1.iloc[i]['Identnummer']) != '':
        data['Quell-Id'].append(df1.iloc[i]['Identnummer'])
        if (df1.iloc[i]['Familienname']) != '' and (df1.iloc[i]['Namensvorsatz']) != '':
            data['Nachname'].append(str(df1.iloc[i]['Namensvorsatz']) + ' ' + str(df1.iloc[i]['Familienname']))
        elif (df1.iloc[i]['Familienname']) != '':
            data['Nachname'].append(df1.iloc[i]['Familienname'])
        else:
            data['Nachname'].append("FEHLER")
        if (df1.iloc[i]['Vorname']) != '':
            data['Vorname'].append(df1.iloc[i]['Vorname'])
        else:
            data['Vorname'].append("FEHLER")
        data['Typ'].append('LAA')
        if df1.iloc[i]['Lehramt'] in lehraemter:
            data['Lehramt'].append('LAA_'+str(lehraemter[df1.iloc[i]['Lehramt']]))
        else:
            data['Lehramt'].append("FEHLER")
    else:
        datafail['Quell-Id'].append('IdentNr fehlt')
        if df1.iloc[i]['Familienname'] != "" and (df1.iloc[i]['Namensvorsatz']) != '':
            datafail['Nachname'].append(str(df1.iloc[i]['Namensvorsatz']) + ' ' + str(df1.iloc[i]['Familienname']))
        elif df1.iloc[i]['Familienname'] != "":
            datafail['Nachname'].append(df1.iloc[i]['Familienname'])
        else:
            datafail['Nachname'].append('FEHLER')
        if df1.iloc[i]['Vorname'] != '':
            datafail['Vorname'].append(df1.iloc[i]['Vorname'])
        else:
            datafail['Vorname'].append('FEHLER')
        datafail['Typ'].append('LAA')
        if df1.iloc[i]['Lehramt'] != "":
            if df1.iloc[i]['Lehramt'] in lehraemter:
                datafail['Lehramt'].append('LAA_'+str(lehraemter[df1.iloc[i]['Lehramt']]))
            else:
                datafail['Lehramt'].append('FEHLER')
        else:
            datafail['Lehramt'].append('FEHLER')

# safe results in new dataframes
df2 = pd.DataFrame(data, columns = ['Quell-Id', 'Nachname', 'Vorname', 'Typ', 'Lehramt'])
df3 = pd.DataFrame(datafail, columns = ['Quell-Id', 'Nachname', 'Vorname', 'Typ', 'Lehramt'])

# display error-results for logineo-users
if not df3.empty:
    print(df3)
    output_filename_error = dt_string+'_Referendare_FEHLER.xlsx'
    output_filepath_error = os.path.join( output_dir, output_filename_error )
    df3.to_excel(output_filepath_error, 'Referendare-FEHLER')
    print("")
    print("WARNUNG - WARNUNG - WARNUNG")
    print("Bei einigen importierten Zeilen sind Probleme/Fehler aufgetreten (siehe oben).")
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
    output_filepath = os.path.join( output_dir, output_filename )
    df2.to_excel(output_filepath, 'Referendare')
    print("")
    print("\nEs wurde erfolgreich eine Datei mit den Referendaren im Output-Ordner angelegt.")
    print("\nSie können diese Datei nun in der Nutzerverwaltung LOGINEO NRWs importieren.")
    exit
else:
    print("\nIhre Tabelle enthält keine gültigen Werte. Der Prozess wird abgebrochen.")
    exit
