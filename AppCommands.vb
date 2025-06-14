﻿Public Class AppCommands
    Public Shared ReadOnly AppExit As RoutedUICommand = New RoutedUICommand("Beenden", "AppExit", GetType(FrameworkElement))
    Public Shared ReadOnly ImportFromTestcenter As RoutedUICommand = New RoutedUICommand("Einlesen von Testcenter", "Laden von TC", GetType(FrameworkElement))
    Public Shared ReadOnly ImportFromJson As RoutedUICommand = New RoutedUICommand("Einlesen von JSON", "Import JSON", GetType(FrameworkElement))
    Public Shared ReadOnly ImportBookletsFromJson As RoutedUICommand = New RoutedUICommand("Einlesen Booklet-Größen von JSON", "Import JSON Booklets", GetType(FrameworkElement))
    Public Shared ReadOnly ImportFromCsv As RoutedUICommand = New RoutedUICommand("Einlesen von CSV-Datei", "Import CSV", GetType(FrameworkElement))
    Public Shared ReadOnly ExportToJson As RoutedUICommand = New RoutedUICommand("Export in JSON-Dateien", "Export JSON", GetType(FrameworkElement))
    Public Shared ReadOnly ExportToXlsx As RoutedUICommand = New RoutedUICommand("Export in Xlsx", "Export Xlsx", GetType(FrameworkElement))
    Public Shared ReadOnly ExportResponsesToCsv As RoutedUICommand = New RoutedUICommand("Export Antworten Langformat CSV", "Export Csv", GetType(FrameworkElement))
    Public Shared ReadOnly DBNew As RoutedUICommand = New RoutedUICommand("Neue Datenbank-Datei anlegen", "DB Neu", GetType(FrameworkElement))
    Public Shared ReadOnly DBOpen As RoutedUICommand = New RoutedUICommand("Datenbank-Datei öffnen", "DB Open", GetType(FrameworkElement))
    Public Shared ReadOnly DBCheck As RoutedUICommand = New RoutedUICommand("Datenbank-Datei prüfen", "DB Check", GetType(FrameworkElement))
    Public Shared ReadOnly DBCopyTo As RoutedUICommand = New RoutedUICommand("Kopie von Datenbank erzeugen", "DB Kopieren", GetType(FrameworkElement))
End Class
