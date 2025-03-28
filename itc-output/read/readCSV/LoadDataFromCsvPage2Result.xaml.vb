﻿Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Packaging
Imports System.ComponentModel

Public Class LoadDataFromCsvPage2Result
    Private WithEvents myBackgroundWorker As BackgroundWorker = Nothing

    Private Sub Me_Loaded() Handles Me.Loaded
        Me.MBUC.AddMessage("Bitte warten!")
        Me.BtnCancelClose.IsEnabled = True
        Me.BtnCancelClose.Content = "Abbrechen"

        myBackgroundWorker = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
        myBackgroundWorker.RunWorkerAsync()
    End Sub
    Private Sub myBackgroundWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) Handles myBackgroundWorker.RunWorkerCompleted
        Me.BtnCancelClose.IsEnabled = True
        Me.APBUC.UpdateProgressState(0)
        Me.MBUC.AddMessage("i: Beendet.")
        BtnCancelClose.IsEnabled = True
        BtnCancelClose.Content = "Schließen"
    End Sub

    Private Sub BtnCancelClose_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        If myBackgroundWorker Is Nothing Then
            Dim parentDlg As LoadDataFromCsvDialog = Me.Parent
            parentDlg.DialogResult = True
        Else
            If myBackgroundWorker.WorkerSupportsCancellation AndAlso myBackgroundWorker.IsBusy Then
                myBackgroundWorker.CancelAsync()
                BtnCancelClose.IsEnabled = False
                BtnCancelClose.Content = "Bitte warten"
                Me.MBUC.AddMessage("w: Abbruch - bitte warten!")
            Else
                Dim parentDlg As LoadDataFromCsvDialog = Me.Parent
                parentDlg.DialogResult = True
            End If
        End If
    End Sub

    Private Sub myBackgroundWorker_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs) Handles myBackgroundWorker.ProgressChanged
        If e.ProgressPercentage > 0 Then
            Me.TBInfo.Text = "Zeile " + e.ProgressPercentage.ToString("N0")
        Else
            Me.TBInfo.Text = "-"
            If e.ProgressPercentage < 0 Then Me.APBUC.UpdateProgressState(-e.ProgressPercentage)
        End If

        If Not String.IsNullOrEmpty(e.UserState) Then Me.MBUC.AddMessage(e.UserState)
    End Sub

    Private Sub myBackgroundWorker_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles myBackgroundWorker.DoWork
        Dim myworker As ComponentModel.BackgroundWorker = sender
        Dim parentDlg As LoadDataFromCsvDialog = Me.Parent

        Dim targetXlsxFilename As String = My.Settings.lastfile_OutputTargetXlsx
        Dim myTemplate As Byte() = Nothing
        If parentDlg.WriteToXls Then
            Try
                Dim TmpZielXLS As SpreadsheetDocument = SpreadsheetDocument.Create(targetXlsxFilename, SpreadsheetDocumentType.Workbook)
                Dim myWorkbookPart As WorkbookPart = TmpZielXLS.AddWorkbookPart()
                myWorkbookPart.Workbook = New Workbook()
                myWorkbookPart.Workbook.AppendChild(Of Sheets)(New Sheets())
                TmpZielXLS.Close()

                myTemplate = IO.File.ReadAllBytes(targetXlsxFilename)
            Catch ex As Exception
                myworker.ReportProgress(0.0#, "e: Konnte Datei '" + targetXlsxFilename + "' nicht schreiben (noch geöffnet?)" + vbNewLine + ex.Message)
            End Try
        End If

        If myTemplate IsNot Nothing OrElse Not parentDlg.WriteToXls Then
            Dim Events As New List(Of String)
            parentDlg.AllVariables = New List(Of String)
            Dim AllUnitsWithResponses As New List(Of String)
            Dim LogEntryCount As Long = 0

            Dim SearchDir As New IO.DirectoryInfo(My.Settings.lastdir_OutputSource)
            Dim csvSeparator As String = ";"
            globalOutputStore.clear()

            For Each fi As IO.FileInfo In SearchDir.GetFiles("*.csv", IO.SearchOption.AllDirectories)
                If myworker.CancellationPending Then
                    e.Cancel = True
                    Exit For
                End If

                Dim readFile
                Dim line As String = ""
                Try
                    readFile = New IO.StreamReader(fi.FullName)
                    line = readFile.ReadLine().Replace("""", "")
                Catch ex As Exception
                    readFile = Nothing
                    myworker.ReportProgress(0.0#, "e:Fehler mein Lesen von " + fi.Name + "; noch geöffnet?")
                End Try
                If readFile IsNot Nothing Then
                    myworker.ReportProgress(0.0#, "Lese " + fi.Name)
                    If line = LogSymbols.LogFileFirstLine2024 OrElse line = LogSymbols.LogFileFirstLineLegacy Then
                        If parentDlg.WriteToXls Then
                            LogEntryCount += 1
                            Dim fileType As CsvLogFileType = CsvLogFileType.Legacy
                            If line = LogSymbols.LogFileFirstLine2024 Then fileType = CsvLogFileType.v2024
                            Dim lineNumber As Long = 0
                            While readFile.Peek() >= 0 AndAlso Not myworker.CancellationPending
                                line = readFile.ReadLine()
                                lineNumber += 1
                                myworker.ReportProgress(lineNumber)
                                Dim logData As UnitLineDataLog = UnitLineDataLog.fromCsvLine(line, fileType)
                                If logData IsNot Nothing Then globalOutputStore.personDataFull.AddLogEntry(logData)
                            End While
                            If myworker.CancellationPending Then e.Cancel = True
                        End If
                    Else
                        '#########################
                        Dim lineNumber As Long = 0
                        Dim csvType As CsvResponseFileType = CsvResponseFileType.v2024
                        If line = ResponseSymbols.ResponsesFileFirstLineLegacy Then
                            csvType = CsvResponseFileType.Legacy
                        ElseIf line = ResponseSymbols.ResponsesFileFirstLine2019 Then
                            csvType = CsvResponseFileType.v2019
                        End If
                        While readFile.Peek() >= 0 AndAlso Not myworker.CancellationPending
                            line = readFile.ReadLine()
                            lineNumber += 1
                            myworker.ReportProgress(lineNumber)
                            Dim unitData As UnitLineDataResponses = UnitLineDataResponses.fromCsvLine(line, csvSeparator, csvType)
                            If unitData.subforms IsNot Nothing AndAlso unitData.subforms.Count > 0 AndAlso unitData.subforms.First.responses.Count > 0 Then
                                If Not AllUnitsWithResponses.Contains(unitData.unitname) Then AllUnitsWithResponses.Add(unitData.unitname)
                                For Each entry As SubForm In unitData.subforms
                                    For Each respData As ResponseData In entry.responses
                                        If Not parentDlg.AllVariables.Contains(unitData.unitname + "##" + respData.id) Then parentDlg.AllVariables.Add(unitData.unitname + "##" + respData.id)
                                    Next
                                Next
                                globalOutputStore.personDataFull.AddUnitData(unitData)
                            End If
                        End While
                        If myworker.CancellationPending Then e.Cancel = True
                    End If
                End If
            Next
            myworker.ReportProgress(0.0#, "beendet.")


            If Not myworker.CancellationPending Then
                If parentDlg.WriteToXls Then
                    Dim config As New WriteXlsxConfig With {
                        .targetXlsxFilename = targetXlsxFilename,
                        .writeResponsesCodes = False,
                        .writeResponsesScores = False,
                        .writeResponsesStatus = False,
                        .writeResponsesValues = True,
                        .writeSessions = False
                        }
                    WriteOutputToXlsx.Write(myTemplate, myworker, e, config)
                Else
                    Dim maxProgressValue As Integer = globalOutputStore.personDataFull.Count
                    Dim progressValue As Integer = 1
                    For Each p As KeyValuePair(Of String, Person) In globalOutputStore.personDataFull
                        If myworker.CancellationPending Then Exit For
                        myworker.ReportProgress(progressValue * -100 / maxProgressValue, p.Key)
                        progressValue += 1
                        parentDlg.sqliteConnection.addPerson(p.Value)
                    Next
                    parentDlg.sqliteConnection.WriteDbInfoData(True)
                End If
            End If
        End If
    End Sub

End Class
