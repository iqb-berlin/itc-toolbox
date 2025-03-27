Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports System.ComponentModel

Public Class LoadDataFromTestcenterPage4Responses
    Private WithEvents myBackgroundWorker As BackgroundWorker = Nothing

    Private Sub Me_Loaded() Handles Me.Loaded
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
        If ParentDlg.target = DataTarget.JsonFiles Then
            Dim folderpicker As New Forms.FolderBrowserDialog With {.Description = "Zielverzeichnis für Dateien",
                                                            .ShowNewFolderButton = True, .SelectedPath = My.Settings.lastfolder_OutputTarget}
            If folderpicker.ShowDialog() AndAlso Not String.IsNullOrEmpty(folderpicker.SelectedPath) Then
                My.Settings.lastfolder_OutputTarget = folderpicker.SelectedPath
                My.Settings.Save()

                myBackgroundWorker = New BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
                myBackgroundWorker.RunWorkerAsync()
            Else
                ParentDlg.DialogResult = False
            End If
        Else
            myBackgroundWorker = New BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
            myBackgroundWorker.RunWorkerAsync()
        End If
    End Sub

    Private Sub myBackgroundWorker_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles myBackgroundWorker.DoWork
        Dim myBW As BackgroundWorker = sender
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent

        Dim LogEntryCount As Long = 0
        Dim maxProgressValue As Integer = ParentDlg.selectedDataGroups.Count * 2
        Dim progressValue As Integer = 0
        Dim fatalError As Boolean = False

        If ParentDlg.target = DataTarget.Datastore Then globalOutputStore.clear()

        myBW.ReportProgress(3.0#, "Lese Booklets")
        Dim bookletSizes As Dictionary(Of String, Long) = ParentDlg.itcConnection.getBookletSizes()
        If Not String.IsNullOrEmpty(ParentDlg.itcConnection.lastErrorMsgText) Then
            myBW.ReportProgress(3.0#, "e: Problem beim Lesen der Booklets: " +
                ParentDlg.itcConnection.lastErrorMsgText)
        End If
        If Not myBW.CancellationPending Then
            If ParentDlg.target = DataTarget.JsonFiles Then
                myBW.ReportProgress(3.0#, "Schreibe Booklets.JSON")
                Try
                    JsonReadWrite.WriteBooklets(My.Settings.lastfolder_OutputTarget + IO.Path.DirectorySeparatorChar + "booklets.json")
                Catch ex As Exception
                    myBW.ReportProgress(3.0#, "Fehler beim Schreiben von Booklets.JSON: " + ex.Message)
                    fatalError = True
                End Try
            ElseIf ParentDlg.target = DataTarget.Datastore Then
                globalOutputStore.bookletSizes = bookletSizes
            ElseIf ParentDlg.target = DataTarget.Sqlite Then
                ParentDlg.sqliteConnection.addBooklets(bookletSizes)
            End If
        End If

        For Each dataGroupId As String In ParentDlg.selectedDataGroups
            If myBW.CancellationPending Then Exit For
            Dim personDataList As New PersonList
            If myBW.CancellationPending Then
                e.Cancel = True
                Exit For
            End If
            If fatalError Then Exit For
            myBW.ReportProgress(progressValue * 100 / maxProgressValue, "Lese '" + dataGroupId + "': ")
            Dim logData As List(Of LogEntryDTO) = ParentDlg.itcConnection.getLogs(dataGroupId)
            If myBW.CancellationPending Then Exit For
            If Not String.IsNullOrEmpty(ParentDlg.itcConnection.lastErrorMsgText) Then
                myBW.ReportProgress(progressValue * 100 / maxProgressValue, "e: Problem bei Logingruppe '" + dataGroupId + "': " +
                        ParentDlg.itcConnection.lastErrorMsgText + " (Logs)")
                fatalError = True
            Else
                For Each log As LogEntryDTO In logData
                    LogEntryCount += 1
                    Dim logEntry As UnitLineDataLog = UnitLineDataLog.fromTestcenterAPI(log)
                    If logData IsNot Nothing Then personDataList.AddLogEntry(logEntry)
                Next
            End If
            progressValue += 1

            myBW.ReportProgress(progressValue * 100 / maxProgressValue)
            Dim responseDataList As List(Of ResponseDTO) = ParentDlg.itcConnection.getResponses(dataGroupId)
            If myBW.CancellationPending Then Exit For
            If Not String.IsNullOrEmpty(ParentDlg.itcConnection.lastErrorMsgText) Then
                myBW.ReportProgress(progressValue * 100 / maxProgressValue, "e: Problem bei Logingruppe '" + dataGroupId + "': " +
                    ParentDlg.itcConnection.lastErrorMsgText + " (Responses)")
                fatalError = True
            Else
                For Each responseData As ResponseDTO In responseDataList
                    Dim unitData As UnitLineDataResponses = UnitLineDataResponses.fromTestcenterAPI(responseData)
                    If unitData.subforms IsNot Nothing AndAlso unitData.subforms.Count > 0 AndAlso
                            unitData.subforms.First.responses.Count > 0 Then personDataList.AddUnitData(unitData)
                Next
            End If

            If ParentDlg.target = DataTarget.Datastore OrElse ParentDlg.target = DataTarget.Xlsx Then
                For Each p As KeyValuePair(Of String, Person) In personDataList
                    globalOutputStore.personDataFull.Add(p.Key, p.Value)
                Next
            ElseIf ParentDlg.target = DataTarget.JsonFiles Then
                myBW.ReportProgress(progressValue * 100 / maxProgressValue, "Schreibe '" + dataGroupId + "'")

                Try
                    JsonReadWrite.Write(My.Settings.lastfolder_OutputTarget + IO.Path.DirectorySeparatorChar + dataGroupId + ".json", personDataList)
                Catch ex As Exception
                    myBW.ReportProgress(3.0#, "Fehler beim Schreiben von " + dataGroupId + ".JSON: " + ex.Message)
                    fatalError = True
                End Try
            ElseIf ParentDlg.target = DataTarget.Sqlite Then
                For Each p As KeyValuePair(Of String, Person) In personDataList
                    ParentDlg.sqliteConnection.addPerson(p.Value)
                Next
            End If
            personDataList.Clear()
            logData.Clear()
            responseDataList.Clear()
            GC.Collect()
            GC.WaitForFullGCComplete()

            progressValue += 1
        Next
        If ParentDlg.target = DataTarget.Xlsx Then
            Dim targetXlsxFilename As String = My.Settings.lastfile_OutputTargetXlsx
            Dim myTemplate As Byte() = Nothing
            Try
                Dim TmpZielXLS As SpreadsheetDocument = SpreadsheetDocument.Create(targetXlsxFilename, SpreadsheetDocumentType.Workbook)
                Dim myWorkbookPart As WorkbookPart = TmpZielXLS.AddWorkbookPart()
                myWorkbookPart.Workbook = New Workbook()
                myWorkbookPart.Workbook.AppendChild(Of Sheets)(New Sheets())
                TmpZielXLS.Close()

                myTemplate = IO.File.ReadAllBytes(targetXlsxFilename)
            Catch ex As Exception
                myBW.ReportProgress(0.0#, "e: Konnte Datei '" + targetXlsxFilename + "' nicht schreiben (noch geöffnet?)" + vbNewLine + ex.Message)
            End Try

            If myTemplate IsNot Nothing Then
                Dim config As New WriteXlsxConfig With {
                        .targetXlsxFilename = My.Settings.lastfile_OutputTargetXlsx,
                        .writeResponsesCodes = False,
                        .writeResponsesScores = False,
                        .writeResponsesStatus = False,
                        .writeResponsesValues = True,
                        .writeSessions = False
                        }
                WriteOutputToXlsx.Write(myTemplate, myBW, e, config)
            End If
        ElseIf ParentDlg.target = DataTarget.Sqlite Then
            ParentDlg.sqliteConnection.WriteDbInfoData(False)
        End If

        myBW.ReportProgress(0.0#, "beendet.")
    End Sub

    Private Sub myBackgroundWorker_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs) Handles myBackgroundWorker.ProgressChanged
        Me.APBUC.UpdateProgressState(e.ProgressPercentage)
        If Not String.IsNullOrEmpty(e.UserState) Then Me.MBUC.AddMessage(e.UserState)
    End Sub

    Private Sub myBackgroundWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) Handles myBackgroundWorker.RunWorkerCompleted
        APBUC.Value = 0.0#
        MBUC.AddMessage("beendet")
        BtnCancelClose.Content = "Schließen"
        BtnCancelClose.IsEnabled = True
        If e.Cancelled Then MBUC.AddMessage("durch Nutzer abgebrochen.")
        Me.APBUC.UpdateProgressState(0.0#)
    End Sub

    Private Sub BtnCancelClose_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        If myBackgroundWorker IsNot Nothing AndAlso myBackgroundWorker.IsBusy Then
            myBackgroundWorker.CancelAsync()
            BtnCancelClose.IsEnabled = False
            MBUC.AddMessage("Abbruch - bitte warten")
        Else
            Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
            ParentDlg.DialogResult = False
        End If
    End Sub

End Class
