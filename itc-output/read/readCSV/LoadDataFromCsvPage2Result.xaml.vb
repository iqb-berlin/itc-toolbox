Imports Newtonsoft.Json
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Packaging
Imports iqb.lib.openxml
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
            Dim parentDlg As OutputDialog = Me.Parent
            parentDlg.DialogResult = True
        Else
            If myBackgroundWorker.WorkerSupportsCancellation AndAlso myBackgroundWorker.IsBusy Then
                myBackgroundWorker.CancelAsync()
                BtnCancelClose.IsEnabled = False
                BtnCancelClose.Content = "Bitte warten"
                Me.MBUC.AddMessage("w: Abbruch - bitte warten!")
            Else
                Dim parentDlg As OutputDialog = Me.Parent
                parentDlg.DialogResult = True
            End If
        End If
    End Sub

    Private Sub myBackgroundWorker_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs) Handles myBackgroundWorker.ProgressChanged
        Me.TBInfo.Text = IIf(e.ProgressPercentage > 0, "Zeile " + e.ProgressPercentage.ToString("N0"), "-")
        If Not String.IsNullOrEmpty(e.UserState) Then Me.MBUC.AddMessage(e.UserState)
    End Sub

    Private Sub myBackgroundWorker_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles myBackgroundWorker.DoWork
        Dim myworker As ComponentModel.BackgroundWorker = sender
        Dim parentDlg As OutputDialog = Me.Parent

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

            'Dim LogData As New Dictionary(Of String, Dictionary(Of String, Long))
            Dim SearchDir As New IO.DirectoryInfo(My.Settings.lastdir_OutputSource)
            Dim csvSeparator As String = ";"
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
                            Dim unitData As UnitLineDataResponses = UnitLineDataResponses.fromCsvLine(line, parentDlg.outputConfig.variables,
                                                                                    csvSeparator, parentDlg.segregateBigdata, csvType)
                            If unitData.subforms IsNot Nothing AndAlso unitData.subforms.Count > 0 AndAlso unitData.subforms.First.responses.Count > 0 AndAlso
                                    (parentDlg.outputConfig.omitUnits Is Nothing OrElse Not parentDlg.outputConfig.omitUnits.Contains(unitData.unitname)) Then
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


            If Not myworker.CancellationPending AndAlso parentDlg.WriteToXls Then WriteOutputToXlsx.Write(myTemplate, myworker, e, targetXlsxFilename)
        End If
    End Sub

End Class
