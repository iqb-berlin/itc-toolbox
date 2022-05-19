# itc-ToolBox – Logins in Docx

Wenn Logins erzeugt wurden und in Form von Testtaker.Xml in das Testsystem hochgeladen wurden (siehe Funktion der itc-ToolBox [Logins aus Vorlage erzeugen](/docs/manual-logins-template.md)), muss man diese Logins den Testpersonen irgendwie mitteilen. Die Funktion "Logins in Docx" liest eine für diese Zwecke vorbereitete Excel-Tabelle und erzeugt Word-Docx-Dateien, die dann - ggf. nachdem sie manuell geändert wurden - ausgedruckt und zerschnitten oder als PDF-Datei verschickt werden können.

## Vorlage

Für die Erzeugung ist eine Vorlage-Datei erforderlich. Dies ist eine normale Docx-Datei. Hier ist für ein Login über Platzhalter festgelegt, was ein Login-Eintrag enthalten soll. Dieser Text in der Vorlage mit den Platzhaltern wird für jeden Login kopiert und die Platzhalter mit den entsprechenden Werten ersetzt.

Als Platzhalter dienen Inhaltselemente vom Typ *Nur-Text-Inhaltssteuerelement*. Dieses Element kann man auswählen, wenn man sich im Menüband das Register *Entwicklertools* aktiviert. Das Hantieren mit diesen Elementen ist etwas fummelig und erfordert Übung. Man kann beispielsweise den *Entwurfsmodus* einschalten, was manchmal aber auch hinderlich ist.

Über den Schalter *Eigenschaften* kann man die Eigenschaft *Tag* für ein Element festlegen. Hier ist ein Schlüsselwort einzutragen (s. Liste unten). Beim Erzeugen der Logins-Docx wird dann diese Markierung erkannt und das gesamte Inhaltselement ersetzt durch den entsprechenden Text bzw. die Grafik.

| Schlüsselwort | Funktion |
| -------- | --------- |
| `server-url` | Es wird die Internet-Adresse des Servers eingefügt. Diese Adresse wird nicht aus der Logins-Xlsx genommen, sondern wird im Dialog des Erzeugens abgefragt |
| `login` | Es wird der Benutzername (Login) eingefügt |
| `password` | Es wird das Kennwort / Password eingefügt |
| `link` | Es wird ein direkter Link eingefügt. Es handelt sich um eine Erweiterung der Server-Adresse um den Benutzernamen. Wenn kein Kennwort vergeben wurde, gelangt die Testperson darüber direkt zur Auswahl des Testheftes und muss nicht das Anmeldeformular ausfüllen |
| `link-qr` | Es wird eine Grafik eingefügt. Es handelt sich um einen QR-Code des Links. Wenn man bei der Testung z. B. Tablets verwendet, kann die Testperson den Code scannen und so die Seite schnell aufrufen. |
| `testgroup-name` | Der Name der Testgruppe wird eingefügt. |
| `testgroup-id` | Die ID der Testgruppe wird eingefügt. |
| `mode` | Der Modus ist ein Schlüsselwort, das in der Testtaker.Xml die Art der Anmeldung steuert. Es kann bei der Fehlersuche hilfreich sein, den Modus mit auszugeben. |

#### Beispiele für Vorlagen

Es sind zwei Vorlagen vorbereitet:
* [Logins-Vorlage1.docx](/Logins-Vorlage1.docx): Alle Daten werden ausgegeben außer der QR-Code des Links.
* [Logins-Vorlage2.docx](/Logins-Vorlage2.docx): Wenige Daten werden ausgegeben und zusätzlich der QR-Code des Links.

#### Vorlage testen!

Das hier vorgestellte Verfahren soll nicht alle erdenklichen Szenarien abdecken und ist gegenüber vielen Varianten der Gestaltung nicht robust. Die Ersetzungstechnik funktioniert in den beiden Vorlagen, aber ein Layout über Tabellen beispielsweise wird nicht unterstützt. Änderungen sollten also unbedingt getestet werden.
