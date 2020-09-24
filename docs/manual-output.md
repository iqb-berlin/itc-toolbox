firstUnitRunning# itc-ToolBox – Antworten
Über den Admin-Bereich des Testcenters lassen sich vor allem zwei Dateiarten 
herunterladen: Responses und Logs. Diese Rohdaten sind schlecht auswertbar. 
Die Funktion "Antworten und Logs csv -> xlsx" transformiert diese Daten. 
Dieser Text beschreibt die Struktur dieser erzeugten Daten.

Der Anwendung wird zunächst ein Verzeichnis mitgeteilt, in dem die 
Response- und Log-Daten im CSV-Format liegen. Bei kleineren Erhebungen 
sind dies zwei Dateien, bei größeren Studien könnte eine Aufteilung in 
viele Dateien erforderlich sein. Zusätzlich kann im Ordner eine oder mehrere yaml-Dateien 
hinterlegt sein. Diese Dateien liefern zusätzliche Informationen für die Transformation 
(s. unten).

Als Ausgabe wird eine Xlsx-Datei erzeugt. Diese enthält in drei Tabellen die 
gewünschten Daten. Nachfolgend wird die Bedeutung der Spalten jeder dieser 
Tabellen beschrieben:

## Tabelle Responses

| Spaltenbezeichnung | Bedeutung |
| :------------- | :---------- |
|ID|Kombination aus anderen (nachfolgenden) Informationen. Diese ID wird benötigt, um eine Zeile eindeutig zu identifizieren. Es handelt sich um eine Testsitzung, also eine Testperson beantwortet ein konkretes Booklet. Diese Kombination ist nötig, weil eine Testperson mehrere Booklets haben könnte und in einem Booklet theoretisch dieselbe Unit enthalten sein kann (z. B. Motivationsabfrage). Es muss eindeutig sein, in welchem Booklet diese Unit platziert war.<br> Diese ID wird auch in den anderen zwei Tabellen verwendet, so dass hierüber eine Zusammenführung der Informationen erfolgen kann.|
|Group|Gruppe, in der das Anmelde-Login platziert war. Dies ist normalerweise nur ein Ordnungsmerkmal für das Monitoring der Durchführung.|
|Login+Code|Entsprechend der Anmeldung der Testperson|
| Booklet | ID des Booklets |
| Variablen nach dem Schema<br>`<Unit-ID>##<innere ID>`<br>z. B. `EL105R##canvasElement10` | Der Player des Testcenters speichert im bisherigen Modell die Antwortdaten als Paarung ID->Wert ab, wobei nicht definiert ist, was ID kennzeichnet (Item, Aspekt eines Items, Eingabeelement des Formulars usw.; hier mal als innere ID bezeichnet). Sicher ist nur, dass diese ID innerhalb der Unit eindeutig ist, und da die Unit-ID eindeutig für das Testheft ist, erlaubt die Kombination Unit-ID mit dieser inneren ID eine eindeutige Zuordnung des Antwort-Wertes zu einer Testperson in einem Booklet, wodurch sich die übliche zweidimensionale Struktur der Antwortdaten ergibt.<br> Es werden nur Units berücksichtigt, die tatsächlich Antwortdaten produziert haben. Reine Textseiten, die z. B. nur Instruktionen enthalten, werden nicht in die Tabelle aufgenommen.<br>Die Variablenspalten werden alphabetisch sortiert ausgegeben.<br>Sollte eine Unit mehrfach in einem Test vorkommen, fügt das System ab dem zweiten Vorkommen der Unit automatisch ein Suffix hinzu:<br>`<Unit-ID>%<n>`<br>n steht hier für die fortlaufende Nummerierung, beginnend mit 1 bei dem zweiten Vorkommen der Unit|

Die Zeilen dieser Tabelle sind nach ID sortiert. Sollte eine Testperson den Test nur gestartet, aber keine Antwortdaten abgeschickt haben, erscheint sie nicht in der Liste.

## Tabelle TimeOnPage

Für die weitere Beurteilung der Antworten schickt das IQB-Testcenter eine größere 
Menge zeitpunktbezogener Daten, sog. Log-Daten. Hierbei wird stets ein Zeitstempel 
mitgeliefert (Datum und Uhrzeit auf dem Computer der Testperson) sowie Art des 
Ereignisses und ggf. weitere Informationen. Aus dieser Folge von Ereignissen lässt 
sich die Navigation zwischen Units und Seiten und somit die Zeit ermitteln, die eine 
Testperson während des Tests auf einer bestimmten Seite verbracht hat.

Eine Unit besteht aus mindestens einer Seite, auf denen die visuellen Elemente 
sowie die Antwort-Elemente platziert sind. Sollte es bei mehreren Seiten eine Seite 
geben, die ständig angezeigt wird, wird diese nicht extra aufgeführt, sondern deren 
Zeit müsste als Summe aller anderen Seitenzeiten ermittelt werden. Die Verteilung 
der Antwort-Elemente (z. B. Items) auf die Seiten wird nicht erfasst.

Die erste Spalte enthält die ID (s. Tabelle Responses), dann folgen für jede Seite 
zwei Spalten:
* `<Unit-ID>##<Seiten-ID>##topTotal`<br>
Zeit, die die Seite angezeigt wurde. Diese Verweildauer ist eine Summe aller 
Besuche in Millisekunden. Konnte für eine Seite kein Endzeitpunkt festgestellt 
werden (z. B. wegen eines Neuladen des Tests), wird hier eine Null eingetragen.
* `<Unit-ID>##<Seiten-ID>##topCount`<br>
Anzahl aller Besuche dieser Seite.

Die Unitseiten-Spalten sind alphabetisch sortiert. Es werden nur Units beachtet, 
die in der Tabelle Responses aufgeführt sind, d. h. nur Units mit Antwortdaten.

Die Verweilzeiten werden nach folgender Methode ermittelt: Startzeitpunkt 
wird durch PAGENAVIGATIONCOMPLETE festgestellt. Dieses Ereignis wird durch 
den Player ausgelöst, wenn alle Elemente der Seite vollständig dargestellt 
sind und ggf. eine vorherige Beantwortung rekonstruiert wurde. Ein Ende der 
Anzeige wird angenommen, wenn ein Ereignis auftritt, das nicht in der folgenden 
Liste aufgeführt ist: RESPONSESCOMPLETE, PRESENTATIONCOMPLETE und UNITTRYLEAVE.

## Tabelle TechData

In dieser Tabelle sind weitere eventuell interessante Daten einer Testsitzung 
aufgelistet:

| Spaltenbezeichnung | Bedeutung |
| :------------- | :---------- |
|ID|ID der Testsitzung wie in den anderen Tabellen|
|Start at|Beginn des ersten Ladens der Testinhalte nach Auswahl des Booklets durch die Testperson. Es handelt sich um eine in JavaScript über Date.now() ermittelte Anzahl der Millisekunden, die seit dem 01.01.1970 00:00:00 UTC vergangen sind. Für Excel muss man den Wert umrechnen: `=<ts>/(1000*60*60*24) + 25569` und dann als Datum+Zeit formatieren: TT.MM.JJJJ h:mm:ss|
|loadcomplete after|Dauer des Ladevorganges in Millisekunden|
|loadspeed|Ladegeschwindigkeit als Quotient aus Bookletgröße (aus der zusätzlich zugewiesenen txt-Datei) und Ladedauer. Wenn die Bookletgröße in Bytes und die Dauer in Millisekunden angegeben werden (wie hier aktuell im Testcenter), dann ist die Einheit des Wertes kBytes/sec|
|firstUnitRunning after|Zeit zwischen Start des Ladens der Testinhalte und Eintritt in die erste Unit. Achtung: In Abhängigkeit von Testhefteinstellungen kann die erste Unit angezeigt werden, bevor alle Testinhalte geladen wurden.|
|os|Betriebssystem (operating system)|
|browser|Name und Version|
|screen|Breite x Höhe in Pixels|
|pages visited ratio|Anteil der in dieser Testsitzung besuchten Seiten an der Gesamtzahl der für dieses Booklet bekannten Seiten. Bekannt sind Seiten durch den Besuch irgendeiner Testperson, d. h. zu einem frühen Zeitpunkt der Testung mögen nicht alle Seiten bekannt sein. Es werden nur Seiten entsprechend der Tabelle TimeOnPages berücksichtigt, d. h. Seiten von Units, zu denen Antworten gespeichert wurden.|
|units fully responded ratio|Anteil der vollständig beantworteten Units an allen bekannten Units des Booklets. Es wird zur Ermittlung das Ereignis RESPONSESCOMPLETE: „all“ verwendet, das der Player schickt. Es werden nur Units berücksichtigt, zu denen Antworten gespeichert wurden. Es kann sein, dass nachträglich Antworten wieder entfernt wurden, was nicht berücksichtigt wird.|

## Konfiguration über YAML-Datei
Eine YAML-Datei ist eine einfache Textdatei, in der über eine spezielle Syntax Informationen 
hinterlegt sind. Es können drei Konfigurationseinträge `bookletSizes`, `omitUnits`und `variables` definiert werden. 
Hier ein Beispiel:

```
bookletSizes:
  THETLK: 19138814
omitUnits:
- Ex_MC
- Ex_SA
- Ex_MM
- Ex_Audio
variables:
  EL639v2:
    EL639v201:
    - canvasElement10
    - canvasElement11
    - canvasElement12
    - canvasElement13
    EL639v202:
    - canvasElement14
    - canvasElement15
    - canvasElement16
    - canvasElement17
    __omit__:
    - canvasElement3
    - canvasElement4
    - pagesViewed
    - responsesGiven
  EL666:
    EL66601:
    - canvasElement18
    EL66602:
    - canvasElement19
    EL66604:
    - canvasElement20
    EL66605:
    - canvasElement21
    EL66606:
    - canvasElement22
    __omit__:
    - canvasElement17
    - canvasElement34
    - canvasElement51
    - canvasElement37
    - pagesViewed
    - responsesGiven
```

### bookletSizes
Hier wird eine ID eines Testheftes erwartet, Doppelpunkt, und dahinter die Größe des Testheftes in Bytes.
Die Größe der Testhefte erhält man im IQB-Testcenter über die Admin-Funktion `Arbeitsbereich prüfen`. Bitte 
vor dem Übertragen jeweils die Punkte aus der Zahl entfernen. Mit Hilfe der Testheftgröße kann die 
Download-Geschwindigkeit berechnet werden (s. oben).

### omitUnits
Die hier aufgeführten Units werden bei der Berichtstransformation ignoriert. Normalerweise sind den eigentlichen 
Testaufgaben Probe- bzw. Trainingsaufgaben vorgeschaltet, die der Einstellung der Lautstärke usw. dienen. Deren 
Ergebnisse interessieren üblicherweise nicht.

### variables
Für jede Unit (Angabe der ID) werden neue Variablen definiert. Die darunter jeweils aufgeführten Bezeichner 
entsprechen IDs von Daten, die in den Csv zu finden sind. Hier gibt es die Fälle
* Umbenennen: Wenn nur eine einzige Csv-ID genannt ist, erhält die neue Variable deren Wert
* Transformation von Radiobutton-Gruppen: Wenn mehrere Csv-IDs aufgeführt geht der Transformator davon aus, dass 
es sich um Radiobuttons handelt, bei denen nur ein Wert true (also ausgewählt) ist. Dann wird der neuen Variable 
eine Zahl zugewiesen, die der Position der Csv-ID in der Liste entspricht. Eine 0 oder leer zeigt an, dass 
keine Option ausgewählt wurde.
* Ignorieren: Wenn das Schlüsselwort `__omit__` verwendet wird, dann ignoriert der Transformator die hier gelisteten 
IDs. Damit kann man z. B. Audioelemente oder Beispielitems entfernen, deren Wert nicht interessiert.