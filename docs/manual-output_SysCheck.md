# itc-ToolBox – System-Check

Wenn Sie zuvor heruntergeladene System-Check-Dateien in ein XLSX-Format konvertieren möchten, 
betätigen Sie einfach den Button: *SysCheck CSV -> xlsx*. Anschließend wählen Sie den Ort aus an welchem sich die 
Dateien befinden und wählen einen Namen und einen Speicherort für die zu generierende Excel-Tabelle aus. 

Wie sind die Inhalte in dieser Tabelle zu deuten?

### Tabelle Syscheck

In dieser Tabelle sind weitere eventuell interessante Daten einer Testsitzung 
aufgelistet:

| Spaltenbezeichnung | Bedeutung |
| :------------- | :---------- |
|ID|ID der Testsitzung wie in den anderen Tabellen|
|Start at|Beginn des ersten Ladens der Testinhalte nach Auswahl des Booklets durch die Testperson (Zeitstempel).|
|loadcomplete after|Dauer des Ladevorganges in Millisekunden|
|loadspeed|Ladegeschwindigkeit als Quotient aus Bookletgröße (aus der zusätzlich zugewiesenen txt-Datei) und Ladedauer. Wenn die Bookletgröße in Bytes und die Dauer in Millisekunden angegeben werden (wie hier aktuell im Testcenter), dann ist die Einheit des Wertes kBytes/sec|
|firstUnitRunning after|Zeit zwischen Start des Ladens der Testinhalte und Eintritt in die erste Unit. Achtung: In Abhängigkeit von Testhefteinstellungen kann die erste Unit angezeigt werden, bevor alle Testinhalte geladen wurden.|
|os|Betriebssystem (operating system)|
|browser|Name und Version|
|screen|Breite x Höhe in Pixels|
