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

Hier gibt es zwei Bereiche, die Du ausfüllen sollst: B5:H?? und K5:M13.

### Daten der Athleten
B5:H?? enthält die Daten der Starter, Zusätzlich musst Du in Spalte A die Startnummern ergänzen. 
Bei der Listenerzeugung werden nur Einträge berücksichtigt, die auch eine Startnummer in Spalte A haben. 

### Daten der Wertungsrichter
K5:M13 enthält die Daten der Wertungsrichter. Für jeden eingetragenen Wertungsrichter wird später ein Bogen generiert. 
In die tatsächliche Wertung werden nur die Wertungsrichter 1 bis 5 und ggf. der Springer einbezogen. 
Der Beobachter wird als Referenz für die Berechnung der Abweichungen verwendet. 
Die Hospitanten werden nicht weiter bei der Berechnung berücksichtigt. 

### Wichtige Regeln

_Beim Befüllen der Spalte H musst Du darauf achten, dass die richtigen Länderkennzeichner verwendet werden. 
Diese werden später mit den Länderkennzeichnern der Wertungsrichter verglichen, um "Springer" ein- bzw. auszuwechseln, daher __müssen die Länderkennzeichen in Spalte H und in Spalte M übereinstimmen__._

_Wenn __weniger__ als 5 Wertungsrichter eingetragen sind (z.B. 3 für manche Meisterschaften), wird die Gesamtpunktzahl einfach aufaddiert. 
Wenn 5 Wertungsrichter eingetragen sind, dann wird die Gesamtpunktzahl ermittelt als die Summe der 3 "mittleren" Wertungen, d.h. die höchste und die niedrigste Wertung fallen aus der Summe heraus._

_Wenn __weniger__ als 5 Wertungsrichter eingetragen sind solltest Du die Länderkennzeichnung auf "n/a" setzen. 
In diesem Fall wird es nämlich auch keine Springer geben, und das Programm ist nicht darauf getestet, auch bei weniger als 3 Wertungsrichtern die Springer-Logik richtig umzusetzen._

## Generieren der Bewertungsdateien

Jedes der Blätter für die einzelnen Kata enthält über einen "GenerateXYZ" Schalter.
Zusätzlich gibt es einen "Generate All" Schalter auf dem Blatt "Configuration".

Die Schalter starten das Generieren der Bögen, d.h. wenn Du ihn betätigst, passiert folgendes:
* Ein Verzeichnis für die Meisterschaft wird angelegt, darin wird ein Verzeichnis für die jeweilige Kata angelegt.
* In diesem Verzeichnis wird für jedes startende Paar eine Datei für die Bewertung erzeugt. Darin ist für jeden Wertungsrichter ein Blatt vorgesehen.
* Wenn du den "Print" Parameter aktiviert hast, werden alle Bewertungsbögen auch tatsächlich gedrucht.
* Die Konfigurationsparameter aus dem Blatt "Configuration" werden in die Dateien "Results" und "Statistics" kopiert, damit alle Dokumente die gleichen Parameter (z.B. Meisterschaftsname) verwenden.

## Eingeben der Bewertungen
Wenn die Wertungsrichter ihre Bewertung gemacht haben, bringt der Mattenläufer die Bögen zur Auswertung.
Zunächst solltest Du die Blätter auf die korrekte Reihenfolge prüfen. 
Die Reihenfolge zur Eingabe ist __immer__ Wertung 1-5, dann Springer, dann Beobachter.
Der Tausch von Springer und "normalem" Bewerter wird später bei der Berechnung gemacht, das sollst Du nicht selbst ändern.

In den Feldern F13:K35 überträgst Du exakt das, was auf den Bögen der Bewerter steht. 
Aus Grund der Automatisierung der Auswertung sind einige Spalten und Zeilen verborgen, das ist kein Fehler und sollte beim Befüllen keine Schwierigkeiten machen.

Nach dem Übertragen prüfst Du (oder am besten eine zweite Person) nochmals die Richtigkeit des Übertrags. 
Danach trägst Du in Zelle L37 Dein Namenskürzel ein, das startet gleichzeitig die Berechnung.

### Die Auswertung

Das Blatt "Auswertung" enthält die Zusammenfassung aller Wertungsrichter. 

Zeile 7 enthält die Wertung des Beobachters als Referenz.
In den Zeilen 8 bis 13 sind die für die Berechnung des Ergebnisses relevanten Zeilen.
Auf diesem Blatt wird die Ersetzung eines Wertungsrichters durch den Springer vorgenommen, hier siehst Du also, dass ggf. Zeilen vertauscht sind.

Die Zeile 3 enthält das Endergebnis der Bewertung. 
Wie bereits beschrieben ist hier die Berechnungsregel, die Werte von Zeile 8 bis Zeile 12 aufzusummieren, wobei die höchste und die nierdrigste Wertung herausgenommen werden.
Falls eine der Zeilen 8 bis 12 leer ist, geht die Berechnungslogik davon aus, dass Du nur 3 Wertungsrichter einsetzt und summiert Zeilen 8 bis 12 auf, ohne Minimum bzw. Maximum auszusortieren.

Die Ergebnistabelle umfasst eine farbige Markierung, die Abweichungen zwischen den Wertungsrichtern verdeutlicht.

### Wichtige Regeln

_So lange in den Feldern für die Bestätigung der korrekten Datenübertragung nichts steht, werden keine Ergebnisse berechnet._

_Die Felder in der Bewertungstabelle müssen mit "x" bei den Fehlern und mit "+" bzw. "-" bei der Korrektur befüllt werden, alle anderen Werte werden beim Editieren abgewiesen._

_Falls Du auf Grund einer falschen Eingabe Daten kopierst, musst Du sehr vorsichtig sein, dass Du nichts in die verborgenen Zeilen oder Spalten hineinkopierst. Das verfälscht ansonsten die Zählung!_

_Die Blätter werden nach dem generieren geschützt, d.h. Du kannst nicht "versehentlich" Formeln oder Berechnungen überschreiben. Wenn Du die Berechnung nachvollziehen musst, das Passwort für den Blattschutz ist 'generated'._

## Erzeugen von Ergebnislisten

Wenn die ersten Bewertungen übertragen sind, und mindestens eine der Dateien der Starter fertig ausgefüllt ist, kannst Du beginnen, die Ergebnislisten anzufertigen. 
Dazu öffnest Du die Datei "Results" und wechselst auf das Blatt mit der jeweiligen Kata.

Jedes Blatt hat zwei Schalter: 
* "Clear" löscht alle bisher übertragenen Ergebnisse.
* "Import" holt die Ergebnisse aus den Bewertungsdateein und kopiert sie ein eine Ergebnistabelle.

"Clear" brauchst Du normalerweise nicht, der Import löscht als erstes auch alle bestehenden Daten.

Wenn Du "Import" betätigst, liest das Programm aus dem entsprechenden Verzeichnis alle Bewertungsdateien aus, und überträgt aus dem jeweiligen Blatt "Auswertung" die Daten aus Zeile 3 in die Zwischentabelle AV2:CE?
Den Import kannst Du also zu jeder Zeit ausführen, falls ein Paar noch nicht gelaufen ist und die entsprechende Datei leer ist, wird lediglich das leere Ergebnis übertragen.

Wenn alle Daten aus den Bewertungsbögen übertragen sind, kopiert das Programm die Daten in die Ergebnistabelle A5:AG?
Anschließend wird dort die Sortierung vorgenommen, zunächst nach der Gesamtpunktzahl in Spalte D, nachrangig (bei Punktgleichstand) nach der Zahl der großen, dann mittleren, dann kleinen Fehlern.

Der Druckbereich des Blatts ist so eingestellt, dass Du direkt die Ergebnistabelle drucken kannst (A1:AG?). 

## Erzeugen der Gesamtstatistik

Diese Funktion ist aktuell noch nicht vollständig umgesetzt.