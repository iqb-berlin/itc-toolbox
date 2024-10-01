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
            For Each fi As IO.FileInfo In SearchDir.GetFiles("*.csv", IO.SearchOption.AllDirectories)
                If myworker.CancellationPending Then
                    e.Cancel = True
                    Exit For
                End If

                Dim readFile
                Dim line As String = ""
                Try
                    readFile = New IO.StreamReader(fi.FullName)
                    line = readFile.ReadLine()
                Catch ex As Exception
                    readFile = Nothing
                    myworker.ReportProgress(0.0#, "e:Fehler mein Lesen von " + fi.Name + "; noch geöffnet?")
                End Try
                If readFile IsNot Nothing Then
                    myworker.ReportProgress(0.0#, "Lese " + fi.Name)
                    If line = OutputDialog.LogFileFirstLine Then
                        Dim lineNumber As Long = 0
                        While readFile.Peek() >= 0 AndAlso Not myworker.CancellationPending
                            line = readFile.ReadLine()
                            lineNumber += 1
                            myworker.ReportProgress(lineNumber)
                            Dim lineSplits As String() = line.Split({""";"}, StringSplitOptions.RemoveEmptyEntries)
                            If lineSplits.Count <> 7 Then lineSplits = line.Split({";"}, StringSplitOptions.None)
                            If lineSplits.Count = 7 Then
                                LogEntryCount += 1
                                Dim group As String = IIf(lineSplits(0).Substring(0, 1) = """", lineSplits(0).Substring(1), lineSplits(0))
                                Dim login As String = IIf(lineSplits(1).Substring(0, 1) = """", lineSplits(1).Substring(1), lineSplits(1))
                                Dim code As String = ""
                                If Not String.IsNullOrEmpty(lineSplits(2)) Then code = IIf(lineSplits(2).Substring(0, 1) = """", lineSplits(2).Substring(1), lineSplits(2))
                                Dim booklet As String = IIf(lineSplits(3).Substring(0, 1) = """", lineSplits(3).Substring(1).ToUpper, lineSplits(3).ToUpper)
                                Dim unit As String = ""
                                If Not String.IsNullOrEmpty(lineSplits(4)) Then unit = IIf(lineSplits(4).Substring(0, 1) = """", lineSplits(4).Substring(1), lineSplits(4))
                                Dim timestampStr As String = IIf(lineSplits(5).Substring(0, 1) = """", lineSplits(5).Substring(1), lineSplits(5))
                                Dim timestampInt As Long = 0
                                If timestampStr.IndexOf("E+") > 0 Then
                                    timestampInt = Long.Parse(timestampStr, System.Globalization.NumberStyles.Float)
                                Else
                                    timestampInt = Long.Parse(timestampStr)
                                End If
                                Dim entry As String = lineSplits(6)
                                If Not String.IsNullOrEmpty(entry) AndAlso entry.Substring(0, 1) = """" Then
                                    entry = entry.Substring(1, entry.Length - 2).Replace("""""", """")
                                End If

                                Dim key As String = entry
                                Dim parameter As String = ""
                                If key.IndexOf(" : ") > 0 Then
                                    parameter = key.Substring(key.IndexOf(" : ") + 3)
                                    If parameter.IndexOf("""") = 0 AndAlso parameter.LastIndexOf("""") = parameter.Length - 1 Then
                                        parameter = parameter.Substring(1, parameter.Length - 2)
                                        parameter = parameter.Replace("""""", """")
                                        parameter = parameter.Replace("\\", "\")
                                    End If
                                    key = key.Substring(0, key.IndexOf(" : "))
                                ElseIf key.IndexOf(" = ") > 0 Then
                                    parameter = key.Substring(key.IndexOf(" = ") + 3)
                                    key = key.Substring(0, key.IndexOf(" = "))
                                End If

                                globalOutputStore.personData.AddLogEntry(group, login, code, booklet, timestampInt, unit, key, parameter)
                            End If
                        End While
                        If myworker.CancellationPending Then e.Cancel = True
                    Else
                        '#########################
                        Dim csvSeparator As String = ";"
                        Dim lineNumber As Long = 0
                        While readFile.Peek() >= 0 AndAlso Not myworker.CancellationPending
                            line = readFile.ReadLine()
                            lineNumber += 1
                            myworker.ReportProgress(lineNumber)
                            Dim unitData As UnitLineData = UnitLineData.fromCsvLine(line, parentDlg.outputConfig.variables, csvSeparator, parentDlg.segregateBigdata)
                            If unitData.subforms IsNot Nothing AndAlso unitData.subforms.Count > 0 AndAlso unitData.subforms.First.responses.Count > 0 AndAlso
                                    (parentDlg.outputConfig.omitUnits Is Nothing OrElse Not parentDlg.outputConfig.omitUnits.Contains(unitData.unitname)) Then
                                If Not AllUnitsWithResponses.Contains(unitData.unitname) Then AllUnitsWithResponses.Add(unitData.unitname)
                                For Each entry As SubForm In unitData.subforms
                                    For Each respData As ResponseData In entry.responses
                                        If Not parentDlg.AllVariables.Contains(unitData.unitname + "##" + respData.id) Then parentDlg.AllVariables.Add(unitData.unitname + "##" + respData.id)
                                    Next
                                Next
                                globalOutputStore.personData.AddUnitData(unitData)
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
