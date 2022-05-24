# itc-ToolBox – Testhefte (Booklets)
Über den Admin-Bereich des Testcenters sind Xml-Dateien zu hinterlegen, die 
festlegen, welche Aufgaben den Testpersonen in welcher Reihenfolge 
vorgelegt werden sollen. Diese Dateien zu pflegen kann aufwändig sein. 
Sollte sich beispielsweise der Titel einer Aufgabe ändern, muss dies in allen 
Testheften geändert werden, in denen diese Aufgabe erscheint.

Die Struktur eines Testheftes am IQB ist recht einheitlich: Aufgaben werden 
in Blöcke platziert, und Testhefte werden aus diesen Blöcken zusammengesetzt. Blöcke 
und damit Aufgaben können in mehreren Testheften vorkommen, um z. B. Positionseffekte 
während einer Pilotierung zu untersuchen. Zeitbeschränkungen sowie Ablaufsynchronisation 
werden auf Blockebene definiert.

Die Funktion "Booklet-Xlsx" der itc-ToolBox liest zunächst eine vorbereitete 
Excel-Datei ein. Hier sind alle nötigen Informationen für die Testhefte in einer 
Art hinterlegt, die gut lesbar und kommunizierbar ist. Spezielle Markierungen innerhalb 
der Datei stellen sicher, dass die itc-ToolBox die Informationen findet. In einem 
zweiten Schritt erzeugt die itc-ToolBox Xml-Dateien für jedes gefundene Testheft.

Eine Vorlage für die Testheft-Xlsx finden Sie [hier](/Booklet-Template.xlsx).

**Beachten Sie bitte:** Änderungen an Tabellenkopfbezeichner und Spalten, können ein ordnungsgemäßes Einlesen der Excelinhalte durch **itc-Toolbox** verhindern! Betrachten Sie dazu das nachfolgende Bild. In diesem sind Spalten, Zellen und Blattbezeichner farblich gekennzeichnet. Die Farben haben die folgende Bedeutung:<br>

* **rot:** Diese Spalten dürfen nicht gelöscht werden!
* **gelb:** Die textuellen Inhalte dürfen nicht verändert werden!
* **blau:** Die textuellen Inhalte können geändert werden. **itc-Toolbox** findet die entsprechenden für die XML-Dateien benötigten Inhalte. Um die Excel-Dateien einheitlich zu halten, wird jedoch von einer Unbenennung abgeraten. 

![Wichtiges zum Editieren der Excel-Vorlage](https://github.com/iqb-berlin/iqb-berlin.github.io/blob/master/assets/TC_FE_ToolBox_ExcelEdit_final.png)

**Erzeugen / Bearbeiten der Booklet.xml**

Nach Auswahl des ersten Punktes: *Booklet-XLSX* muss der Speicherort der entsprechende Excel-Datei angegeben werden.
**itc-Toolbox** überprüft anschließend die Struktur der Excel-Datei und versucht die dort angegebenen Zusatzinformationen auf den anderen
Excel Blattseiten zu finden. Können Bezeichner nicht gefunden oder zugeordnet werden, werden entsprechende Fehlermeldung generiert.
Wurden alle Bezeichner gefunden und konnten entsprechend zugeordnet werden, erscheint eine Meldung über die erfolgreiche Erzeugung
einer `Booklet.xml`. Die `Booklet.xml` wird dann im selben Verzeichnis gespeichert in dem sich auch die Excel-Datei, die als Vorlage diente, 
befindet.

Um die Inhalte und Strukturen dieser Vorlage besser zu verstehen, betrachten Sie bitte die nachfolgenden Abbildungen. Diese gehen näher 
auf die Strukturen der Excel-Inhalte und den Auswirkungen auf die so generierte `Booklet.xml` ein. Achten Sie auch die farbliche Hervorhebung 
in Bild und Code.

![iqb online assessment applications with relations: testcenter](https://github.com/iqb-berlin/iqb-berlin.github.io/blob/master/assets/TC_FE_ToolBox_Excel_final.png)

Auf der Grundlage dieser im obigen Bild dargestellten Excel-Datei und deren Inhalten wird die folgende `Booklet.xml` erzeugt:

![iqb online assessment applications with relations: testcenter](https://github.com/iqb-berlin/iqb-berlin.github.io/blob/master/assets/TC_FE_ToolBox_BookletXml_final.png)
