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
