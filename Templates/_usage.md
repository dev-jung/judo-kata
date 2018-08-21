<Since the most users are currently german this is initially kept in german and will be translated later.>

# Benutzerdokumentation

Bitte lies das, bevor Du das Programmpaket verwendest. Wenn Du Fragen hast, wende Dich bitte direkt an den Autor.

## Prinzipielle Funktionsweise

Mit dem Programmpaket kann man für die Ausrichtung von Katameisterschaften gemäß der IJF Regularien _Bewertungsbögen generieren_, _Bewertungen eingeben_, _Ergebnistabellen erzeugen_ und _Gesamtstatistiken erheben_, und Vergleiche zwischen den einzelnen Bewertern anstellen.

Das Paket is darauf ausgelegt, international eingesetzt werden zu können, daher sind die verwendeten Texte übersetzbar.

Dazu gibt es zunächst die Excel-Datei "Blanks", die die Bewertungsvorlagen enthält, so wie sie auch die Judges an den Tischen bekommen werden. Zusätzlich sind im Excel noch Auswertungsblätter, die es erlauben, eine erste Auswertung für ein Paar zu erzeugen.

Dann gibt es  die Excel-Datei "List", die die Starterlisten enthält, und die Funktion zum Generieren der Bewertungsbögen.

Die Excel-Datei "Results" enthält die Ergebnistabellen, und die Funktion zum Extrahieren aller Bewertungen aus den Generierten Bewertungsbögen.

Um die Bewerter untereinander zu vergleichen, und um eine Beurteilung der Bewerter sowie eine Qualitätskontrolle der Ergebnisse zu ermöglichen gibt es die Datei "Statistics".

Im Verzeichnis "examples" findest Du übrigens eine beispielhaft ausgefüllte Fassung mit Testdaten.

## Setup und Konfiguration

Zunächst brauchst Du ein leeres Verzeichnis, in dem Du für eine Meisterschaft die folgenden Dateien kopierst:

* Blank.xlsm
* List.xlsm
* Results.xlsm
* Statistics.xlsm

Wenn Du auf den Bögen der Bewerter ein spezielles Logo oder sonstige Dinge haben möchtest, öffne die Datei "Blank" und ändere dort die Grafiken und das Layout, so wie Du Dir das vorstellst. 
Fast alle der Texte werden übrigens später gemäß der Konfiguration überschrieben, daher solltest Du Dich auf die Anpassung der Grafiken beschränken. 
In erster Linie sind die Bögen für den DJB entworfen, daher tragen sie auch das "Kata im DJB" Logo. 
Wenn Du eine Meisterschaft in einem Landesverband, oder jedenfalls nicht im Rahmen des DJB ausrichtest, _nimm bitte dieses Logo heraus_.

Als weitere Vorbereitung musst Du die Datei List.xlsm öffnen, und dort zunächst das Blatt __"Configuration"__ editieren.
Hier stellst Du als erstes den Namen der Meisterschaft ein, und das Datum (Zellen B3 und B6).
Wenn Du eine Unterscheidung zwischen Vor- und Finalrunde hast, gib das entsprechend in der Zelle B4 an.
Ansonsten lass die Zelle einfach leer.
Die restlichen Zellen sind für die Internationalisierung notwendig, hier kannst Du also im Bereich B5, B7:B48 die verwendeten Begriffe verändern.

Auf dem Blatt __"Configuration"__ kannst Du auch noch einen wichtigen Parameter in Zelle H8 einstellen.
Wenn Du in die Zelle den Wahrheitswert "FALSCH" oder "FALSE" einträgst, wird beim Generieren der Bögen (siehe unten) __nicht__ auf einen echten Drucker geschrieben, sondern es werden nur PDF DOkumente erzeugt.
Wenn Du in diese Zelle den Wahrheitswert "WAHR" oder "TRUE" einträgst, wird jedes PDF Dokument, was generiert wird, auch an den Standarddrucker geschickt. 
Gerade wenn man viele Paare am Start hat, ist das hilfreich, um nicht jedes PDF extra nochmal öffnen und manuell ausdrucken zu müssen.
Zunächst solltest Du das also auf "FALSCH" stehen lassen, bis die erzeugten PDFs richtig sind. 
Dann kannst Du das auf "WAHR" stellen, alle PDFs nochmal generieren, dann werden sie automatisch gedruckt.

## Vorbereiten der Startlisten

Als zweiten Schritt musst Du die Daten der Starter und der Bewerter eingeben. 
Dazu wechselst Du zunächst in das Tabellenblatt einer Kata, z.B. "1 Nage".

Hier gibt es zwei Bereiche, die Du ausfüllen sollst: B5:H?? und K5:M13

### Daten der Athleten


## Generieren der Bewertungsdateien

## Eingeben der Bewertungen

## Erzeugen von Ergebnislisten

## Erzeugen der Gesamtstatistik