[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=flat-square)](https://opensource.org/licenses/MIT)


# itc Toolbox

Für die Durchführung von Kompetenztests und Befragungen auf dem Computer 
entwickelt das IQB schrittweise Online-Anwendungen. Mitunter werden diese 
Online-Anwendungen durch Programme ergänzt, die man auf einem Windows-Computer 
installieren muss. 

Die Anwendung "itc-Toolbox" generiert XML-Steuerdateien aus Xlsx-Vorlagen und 
konvertiert die Antworten und Log-Daten, die durch das IQB-Testcenter im CSV-Format erzeugt werden, in eine 
besser auswertbare Xlsx-Datei.

Dokumentationen:
* [Erzeugen von Testheft-Xlm](docs/manual-booklet.md)
* [Erzeugen von Codes für Logins](docs/manual-codes.md)
* [Erzeugen von Logins und von Testtaker.Xml](docs/manual-logins-template.md)
* [Erzeugen von Handzetteln mit Logindaten (Docx)](docs/manual-logins-docx.md)
* [Aufbereitung von SystemCheck Daten](docs/manual-output_SysCheck.md).
* [Aufbereitung von Antworten und Logs](docs/manual-output.md).

## Installieren

Es handelt sich um eine Windows-Anwendung, die ohne Administrationsrechte installiert und genutzt werden kann. Die Anwendung ist mit einem gültigen Zertifikat des IQB signiert und sollte keine Warnungen provozieren. Zum Installieren gehen Sie bitte auf folgende Internet-Seite:

[www.iqb.hu-berlin.de/institut/ab/it/itc-ToolBox](https://www.iqb.hu-berlin.de/institut/ab/it/itc-ToolBox)

## Entwickeln
Nach dem Download des Codes sind über den Paketmanager nuget einige Pakete zu installieren. Die Entwicklung erfolgte mit dem Visual Studio 2022 von Microsoft.

## Credits
* DocumentFormat.OpenXml by Microsoft
* Newtonsoft.Json by James Newton-King
* QRCoder by Raffael Herrmann
* YamlDotNet by Antoine Aubry
