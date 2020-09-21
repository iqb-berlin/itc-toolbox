# itc-ToolBox – Logins
Über den Admin-Bereich des Testcenters sind die Zugangsdaten für 
Testpersonen, aber auch für Reviewer und die Testleitung zu hinterlegen. Struturiert 
in Gruppen, bestehen diese meist aus Benutzername und Kennwort, ggf. auch ergänzt 
durch Personencodes. Näheres siehe Wiki des [IQB-Testcenters](https://github.com/iqb-berlin/testcenter-frontend/wiki).

Die Verfahren zur Erzeugung und Verwaltung der Login-Daten sind sehr vielfältig. Mal 
hat man Monitor-Accounts, mal Code-basierte, mal nicht, und die Weiterverarbeitung bzw. 
Dokumentation (Verschicken an die Schulen, Zettelchen auf dem Platz) hat nochmal 
gesonderte Anforderungen. Außerdem kann man oft nicht alle Logins auf einmal erzeugen, 
weil die Rückmeldungen der Schulen nur schleppend eintreffen. Dann muss man zu 
vorhandenen Logins neue erzeugen. Eventuell sind Aspekte des Datenschutzes zu beachten.

Im Moment scheint es daher nicht sinnvoll, ein formales Verfahren für Logins festzulegen. Es 
gibt allerdings einen Schritt hierbei, der sehr anstrengend ist: Festlegen von Codes für 
Logins (Benutzername/-kennzeichen, Kennwort) oder Personencode. Hierbei 
sind eine Reihe von Kriterien zu beachten:
* Sie müssen eindeutig sein: Kein Login-Benutzername darf doppelt vorkommen.
* Die Zeichen müssen gut merkbar sein, damit bei der Übertragung vom Zettel in den 
Computer kein Fehler passiert.
* Die Tasten sollen auf der Computer-Tastatur gut findbar sein. Es sollte zur Eingabe nur 
eine Taste nötig sein. Großbuchstaben und die meisten Sonderzeichen sind also ungünstig.
* Die Zeichen müssen gut lesbar sein: Optisch sehr ähnliche Zeichen wie "n" und "m" 
oder "1" und "l" sind zu vermeiden.
* Die Codes müssen gut sprechbar und akustisch verständlich sein: Sollte z. B. die 
Testleiterin einem Schüler den Code ansagen, darf es keine Fehler geben.
* Es sollte keine Gefahr stehen, dass durch nicht erkannte Zeichenkodierung von Dateien 
Probleme mit Sonderzeichen (Umlaute!) auftreten.

Über die Funktion "Logins-Xlsx" erzeugt die itc-ToolBox eine Excel-Tabelle mit vier 
Spalten: Zweistellige, dreistellige, vierstellige und fünfstellige Codes. Das IQB 
verwendet eine Auswahl aus Kleinbuchstaben und Ziffern - keine Sonderzeichen und Umlaute. 
Kleinbuchstaben wechseln sich dabei mit Ziffern ab. Die Codes kommen in der Tabelle 
jeweils nur einmal vor.

Diese Tabelle soll längere Zeit für eine Studie als eine Quelle benutzt werden, aus 
der man nach Bedarf Codes entnimmt (Copy & Paste) und in die eigentlichen Dokumente zur 
Login-Verwaltung überträgt. Die genutzten Codes sollten gelöscht oder markiert werden, um 
eine Mehrfachverwendung auszuschließen.