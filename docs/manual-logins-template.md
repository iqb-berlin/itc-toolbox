# itc-ToolBox – Logins aus Vorlage erzeugen

Diese Funktion geht davon aus, dass die Schlüsselinformationen für die Erzeugung von Logins in einer speziellen für diese Zwecke vorbereiteten Excel-Tabelle hinterlegt sind. Diese Datei enthält Markierungen, die dann bei der Verarbeitung gefunden werden und das gezielte Ergänzen von Codes erlauben. Eine Vorlage für diese Login-Xlsx finden Sie [hier](/Logins-Vorlage.xlsx).

Die erzeugten Logins werden in diese Login-Datei hineingeschrieben.

## Tabelle Testgruppen

Es gibt zwei Spalten für die Bezeichnung der Testgruppe (z. B. Schule und Klasse), und in den anschließenden Spalten wird die Anzahl der gewünschten Logins festgelegt. Die ID der Testgruppe kann auch vorab festgelegt werden, wenn darüber bestimmte Zusatzinformationen geliefert werden sollen (z. B. Nummer der Schule in der Länderliste). Ist sie nicht festgelegt, wird sie generiert.

## Tabelle Logins

Diese Tabelle wird durch die itc-ToolBox gefüllt. Im Dialog beim Erzeugen kann die Länge der Login-Codes festgelegt werden. Die Länge für das Password kann 0 sein - dann wird kein Password erzeugt (z. B. für das direkte Einloggen über einen Link).

Die Spalten `Schule`, `Klasse`, `ID Gruppe` und `Modus` bereiten das spätere Generieren der Testtaker.Xml vor.

## Tabelle Textersetzungen

Eine Testtaker.Xml kann Ersetzungen enthalten für den Dialog mit der Testperson während der Testdurchführung. Wenn dies in die generierten Xml-Dateien eingefügt werden soll, dann sind diese Key-Value-Paare hier einzutragen.

## Erzeugen der Testtaker.Xml

Nachdem die Logins erzeugt wurden und ggf. manuell nachbearbeitet wurden, kann das Erzeugen der Testtaker.Xml angestoßen werden. Es kann eine einzige Xml mit allen Logins oder mehrere Dateien mit den Logins je einer Gruppe erzeugt werden. Als Testheft wird stets nur eines zugewiesen, dessen Bezeichnung man vorgeben kann.