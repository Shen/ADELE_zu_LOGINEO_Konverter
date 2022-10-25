# Über das ADELE_zu_LOGINEO-Script für Referendare

# Beschreibung
Das ADELE_zu_LOGINEO-Script konvertiert die txt-Ausgabedatei von ADELE in eine Excel-Datei, die einfach in LOGINEO NRW importiert werden kann.
Das Programm funktioniert ausschließlich mit der txt-Ausgabedatei für Referendare, nicht für Seminarausbilder!

Die ursprüngliche ADELE-Ausgabedatei ist als solche nicht für den Import in LOGINEO NRW geeignet.
![output-adele-file](https://user-images.githubusercontent.com/81589/197849771-cebcfba8-eaf7-47fe-856d-1d906b2e6a10.png)

Das Script beschränkt die Ausgabe auf die notwendigen Spalten und fügt zudem in LOGINEO NRW gut nutzbare Gruppen hinzu (Lehramt, Jahrgang, Seminarzugehörigkeit).
![output-excel-file](https://user-images.githubusercontent.com/81589/197849774-66da92f6-8955-4013-8eba-be15e6838ac3.png)

Die Excel-Ausgabedatei kann problemlos im Admin-Bereich von LOGINEO NRW importiert werden.
![import-logineonrw](https://user-images.githubusercontent.com/81589/197849765-faecbafe-717a-494c-9d84-67bc4b28852d.png)


__Wichtige Hinweise__
* Als __Primärschlüssel__ ist unbedingt die __ADELE-ID__ zu empfehlen, da diese immer vorhanden ist. Die IdentNr, die an Schulen als Primärschlüssel vorgesehen ist, fehlt häufig.
* Das Script kann an Windows-Rechnern über die .exe-Datei ausgeführt werden oder alternativ über das Python-Script.

__Bekannte Probleme__
* Einige Personengruppen (z. B. Fachlehrer in Ausbildung, EU-Anpassungsgänge) sind keinem Lehramt zugeordnet. Für diese Personen werden keine Gruppen generiert.

__Disclaimer__

Ich übernehme keinerlei Haftung für die Verwendung dieses Programms. Es wurde nach bestem Wissen und Gewissen programmiert und soll die Arbeit mit LOGINEO NRW im Zusammenspiel mit ADELE erleichtern.

# Anleitung (Windows)
* Laden Sie sich unter https://github.com/Shen/ADELE_to_LOGINEO/releases die aktuelle Version des Scripts herunter.
* Entpacken Sie die Datei in einem Verzeichnis Ihrer Wahl.
* Kopieren Sie Ihre in ADELE erstelle Referendare.txt-Datei in das soeben erstellte Verzeichnis. Wichtig ist, dass die Datei mit der .exe-Datei des Scripts in einem Verzeichnis liegt.
* Öffnen Sie die __config.xml__ mit einem Editor wie zum Beispiel Notepad++ (https://notepad-plus-plus.org/downloads/). Sie werden zumindest den Dateinamen ändern müssen. Schauen Sie auch, ob Ihre ADELE-txt-Datei mit Tabstopps oder mit Semikolon getrennt ist und passen Sie auch dies in der config.xml an. Speichern Sie abschließend die Datei.
![config-xml](https://user-images.githubusercontent.com/81589/197859934-60643211-b4c7-4810-9564-56ddc364fd46.png)
* Führen Sie die Datei __ADELE-zu-LOGINEO-Import-Konverter.exe__ aus und folgen Sie den Anweisungen.
* Wenn das Script erfolgreich durchlaufen wurde, befindet sich in dem Unterordner "output" die neue Excel-Datei, die nun in LOGINEO NRW importiert werden kann.
![script](https://user-images.githubusercontent.com/81589/197861674-7375d4be-8045-4ea7-b0a9-6c8d4cc0c055.png)
