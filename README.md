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
