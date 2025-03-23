Imports DocumentFormat.OpenXml
Imports System.ComponentModel

Public Class LoadDataFromTestcenterPage4Responses
    Private WithEvents myBackgroundWorker As BackgroundWorker = Nothing

    Private Sub Me_Loaded() Handles Me.Loaded
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
        Me.BtnCancelClose.IsEnabled = False
        If ParentDlg.write Then
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

        If ParentDlg.write Then globalOutputStore.clear()

        myBW.ReportProgress(3.0#, "Lese Booklets")
        globalOutputStore.bookletSizes = globalOutputStore.itcConnection.getBookletSizes()
        If ParentDlg.write Then
            myBW.ReportProgress(3.0#, "Schreibe Booklets.JSON")
            Try
                JsonReadWrite.WriteBooklets(My.Settings.lastfolder_OutputTarget + IO.Path.DirectorySeparatorChar + "booklets.json")
            Catch ex As Exception
                myBW.ReportProgress(3.0#, "Fehler beim Schreiben von Booklets.JSON: " + ex.Message)
                fatalError = True
            End Try
        End If

        For Each dataGroupId As String In ParentDlg.selectedDataGroups
            If ParentDlg.write Then globalOutputStore.clear()
            If myBW.CancellationPending Then
                e.Cancel = True
                Exit For
            End If
            If fatalError Then Exit For
            myBW.ReportProgress(progressValue * 100 / maxProgressValue, "Lese '" + dataGroupId + "': ")
            Dim logData As List(Of LogEntryDTO) = globalOutputStore.itcConnection.getLogs(dataGroupId)
            If Not String.IsNullOrEmpty(globalOutputStore.itcConnection.lastErrorMsgText) Then
                myBW.ReportProgress(progressValue * 100 / maxProgressValue, "e: Problem bei Logingruppe '" + dataGroupId + "': " +
                        globalOutputStore.itcConnection.lastErrorMsgText + " (Logs)")
                fatalError = True
            Else
                For Each log As LogEntryDTO In logData
                    LogEntryCount += 1
                    Dim logEntry As UnitLineDataLog = UnitLineDataLog.fromTestcenterAPI(log)
                    If logData IsNot Nothing Then globalOutputStore.personDataFull.AddLogEntry(logEntry)
                Next
            End If
            progressValue += 1

            myBW.ReportProgress(progressValue * 100 / maxProgressValue)
            Dim responseDataList As List(Of ResponseDTO) = globalOutputStore.itcConnection.getResponses(dataGroupId)
            If Not String.IsNullOrEmpty(globalOutputStore.itcConnection.lastErrorMsgText) Then
                myBW.ReportProgress(progressValue * 100 / maxProgressValue, "e: Problem bei Logingruppe '" + dataGroupId + "': " +
                    globalOutputStore.itcConnection.lastErrorMsgText + " (Responses)")
                fatalError = True
            Else
                For Each responseData As ResponseDTO In responseDataList
                    Dim unitData As UnitLineDataResponses = UnitLineDataResponses.fromTestcenterAPI(responseData, ParentDlg.segregateBigdata)
                    If unitData.subforms IsNot Nothing AndAlso unitData.subforms.Count > 0 AndAlso
                            unitData.subforms.First.responses.Count > 0 Then globalOutputStore.personDataFull.AddUnitData(unitData)
                Next
            End If

            If ParentDlg.write Then
                myBW.ReportProgress(progressValue * 100 / maxProgressValue, "Schreibe '" + dataGroupId + "'")

                Try
                    JsonReadWrite.Write(My.Settings.lastfolder_OutputTarget + IO.Path.DirectorySeparatorChar + dataGroupId + ".json")
                Catch ex As Exception
                    myBW.ReportProgress(3.0#, "Fehler beim Schreiben von " + dataGroupId + ".JSON: " + ex.Message)
                    fatalError = True
                End Try

                If ParentDlg.segregateBigdata Then
                    Try
                        JsonReadWrite.WriteBigData(My.Settings.lastfolder_OutputTarget)
                    Catch ex As Exception
                        myBW.ReportProgress(3.0#, "Fehler beim Schreiben von Bigdata: " + ex.Message)
                        fatalError = True
                    End Try
                End If
            End If

            progressValue += 1
        Next
        If ParentDlg.write Then globalOutputStore.clear()
        myBW.ReportProgress(0.0#, "beendet.")
    End Sub

    Private Sub myBackgroundWorker_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs) Handles myBackgroundWorker.ProgressChanged
        Me.APBUC.UpdateProgressState(e.ProgressPercentage)
        If Not String.IsNullOrEmpty(e.UserState) Then Me.MBUC.AddMessage(e.UserState)
    End Sub

    Private Sub myBackgroundWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) Handles myBackgroundWorker.RunWorkerCompleted
        Me.BtnCancelClose.IsEnabled = True
        Me.APBUC.UpdateProgressState(0.0#)
    End Sub

    Private Sub BtnCancelClose_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
        ParentDlg.DialogResult = True
    End Sub

End Class
