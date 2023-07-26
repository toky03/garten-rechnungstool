# Garten Rechnungen erstellen
## Vorbedingungen
Node mit npm installiert mit mindestens Version 18 siehe [node.js Downloads](https://nodejs.org/en/download)

## Einrichtung
1. diese Git Projekt klonen
2. `npm ci` im Projektverzeichniss ausführen
3. Abhängige Dateien in einen Unterordner namens _data_ hinzufügen
  a. _logo.png_ Logo welches als Absender auf den Brief gedruckt werden soll
  b. _mitgliederliste.xlsx_ Excel Datei mit den Daten siehe [Struktur in der mitgliederliste](#struktur-in-mitgliederliste)
  c. _bills_ Ordner in den die Rechnungen als Pdf erstellt werden sollen


#### Struktur im Excel **mitgliederliste.xlsx**
Damit die Daten korrekt aus dem Excel gelesen werden können müssen die folgenden Reiter mit extakt dem angegebenen Namen vorhanden sein

##### _Mitgliederliste_
Die folgende Tabelle Zeigt die Mindestandorderung an das Tabellenblatt _Mitgliederliste_. Wichtig ist vor allem die Reihenfolge.
|Parz.	        | Name | Vorname | Adresse | PLZ | Ort | Tel. | Aa | Spr. | Vorstand |
|----------|------|---------|---------|--------|-----|------------|-----------|-----------|----------|		
|Parzellen nummer|Name als text|Vorname als Text|Strasse und Hausnummer|Postleitzahl|Ortschaftsnamen|Telefon (wird nicht benötigt)|Anzahl Aren als Zahlenwert|Sprache D oder F|'J' falls Mitglied Vorstansmitglied ist ansonsten leer|	

##### _pachtzins_
Diese Tabelle _muss_ explizit so ausgefüllt werden
alle Einträge sind in derselben Währung als Zahl anzugeben. Im gegensatz zur Tabelle [Mitgliederliste](#mitgliederliste) zählt hier nur die erste Zeile nach dem Titel.
|pachtzinz | wasserbezug | GF Abonement | Strom | Versicherung | Mitgliederbeitrag | Reparaturfonds | Verwaltungskosten |
|---------|--------------|--------------|------|--------|-----------|-----------|-----------|
| Zins pro Are| Wasserkosten pro Are|Kosten für das Abonement des "Gartenfreund"|Stromkosten pauschal| Versicherungskosten pauschal |Mitgliederbeitrag pauschal|Beitrag an den Reparaturfonds pauschal|Beitrag an die Verwaltungskosten pauschal|

##### _Rechnungsdetails_
Auch diese Tabelle _muss_ explizit so ausgefüllt werden
Hierbei sind die Spalten massgebend. Die erste Spalte ist die Bezeichnung, während die zweite Spalte die Werte beinhaltet

| | |
|---- | ---- |
|Name | Name in der Adresszeile|
| Adresse |	Adresse ohne Hausnummer |
| Adress Nummer | Hausnummer |
| Postleitzahl | Postleitzahl |
| Stadt | Ortsangabe ohne plz |
|Iban Nummer | gültige Iban Nummer |
| Ueberschrift DE | Überschrift im Brief auf Deutsch |
|Ueberschrift FR | Überschrift im Brief auf Französisch |

## Anwendung

### Build für eine mit Node ausführbare .js Datei
Das Projekt ist in Typescript geschrieben. Damit es mit Node ausführen kann muss das Script `npm run build` ausgeführt werden.
Nach erfolgreichen ausführen ist die transpilierte JavaScript Datei _index.js_ unter _dist_.

### Programm mit Node laufen lassen
> Vorbedingung hierfür ist ein vorhergehendes `npm run build` und die benötigten Dateien im Ordner [Vorbedingungen](#einrichtung) _data_

Skript um das Programm zu starten `npm run start`
Nach erfolgreichem Lauf befinden sich die Rechnungen unter _data/bills_

### Bundle erstellen
Nur notwendig, wenn die Datei in einer einzigen JavaScript datei ohne Abhängigkeiten benötigt wird z.B. [um daraus eine Ausführbare Datei zu erstellen](#ausführbare-dateien-erstellen).
Skript `npm run bundle`.
Nach erfolgreichem Bundle befindet sich die JavaScript Datei _main.js_ unter dem Ordner _dist_.

### Ausführbare Dateien erstellen
> Vorbedingung hierfür ist, dass eine [gebündelte main.js Datei vorhanden](#bundle-erstellen) ist.

Skript `npm run exec`.
Nach erfolgreichem erstellen, befinden sich im Arbeitsverzeichniss drei neue Dateien
- main-win.exe (für Windows)
- main-linux (für Linux)
- main-macos (für Mac)

Welche auf den jeweiligen Systemen (und dem Ordner _data_) direkt ausführen lassen.
