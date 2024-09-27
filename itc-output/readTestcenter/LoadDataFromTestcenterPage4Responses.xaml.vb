Imports Newtonsoft.Json
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Packaging
Imports System.ComponentModel

Public Class LoadDataFromTestcenterPage4Responses
    Private WithEvents myBackgroundWorker As BackgroundWorker = Nothing

    Private Sub Me_Loaded() Handles Me.Loaded
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
        Me.BtnCancelClose.IsEnabled = False
        If ParentDlg.write Then
            Dim defaultDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            If Not String.IsNullOrEmpty(My.Settings.lastfile_OutputTargetXlsx) Then defaultDir = IO.Path.GetDirectoryName(My.Settings.lastfile_OutputTargetXlsx)
            Dim filepicker As New Microsoft.Win32.SaveFileDialog With {.FileName = My.Settings.lastfile_OutputTargetXlsx, .Filter = "Excel-Dateien|*.xlsx",
                                                            .InitialDirectory = defaultDir, .DefaultExt = "xlsx", .Title = "Antworten Zieldatei wählen"}
            If filepicker.ShowDialog Then
                My.Settings.lastfile_OutputTargetXlsx = filepicker.FileName
                My.Settings.Save()

                myBackgroundWorker = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
                myBackgroundWorker.RunWorkerAsync()
            Else
                ParentDlg.DialogResult = False
            End If
        Else
            myBackgroundWorker = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
            myBackgroundWorker.RunWorkerAsync()
        End If
    End Sub

    Private Sub myBackgroundWorker_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles myBackgroundWorker.DoWork
        Dim myBW As BackgroundWorker = sender
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
        Dim targetXlsxFilename As String = My.Settings.lastfile_OutputTargetXlsx
        Dim myTemplate As Byte() = Nothing
        If ParentDlg.write Then
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
        End If

        If myTemplate IsNot Nothing OrElse Not ParentDlg.write Then
            myBW.ReportProgress(3.0#, "Lese Booklets")
            Dim booklets As List(Of BookletDTO) = globalOutputStore.itcConnection.getBooklets()
            globalOutputStore.bookletSizes = (From b As BookletDTO In booklets).ToDictionary(Of String, Long)(Function(b) b.id, Function(b) b.info.totalSize)

            ParentDlg.AllVariables = New List(Of String)
            Dim AllUnitsWithResponses As New List(Of String)
            Dim LogEntryCount As Long = 0

            Dim maxProgressValue As Integer = ParentDlg.selectedDataGroups.Count * 2
            Dim progressValue As Integer = 0
            Dim fatalError As Boolean = False
            For Each dataGroupId As String In ParentDlg.selectedDataGroups
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
                        Dim key As String = log.logentry
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

                        globalOutputStore.personData.AddLogEntry(log.groupname, log.loginname, log.code, log.bookletname, log.timestamp, log.unitname, key, parameter)
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
                        Dim unitData As UnitLineData = UnitLineData.fromTestcenterAPI(responseData, ParentDlg.replaceBigdata)
                        If unitData.responses IsNot Nothing AndAlso unitData.responses.Count > 0 AndAlso
                            unitData.responses.First.responses.Count > 0 Then globalOutputStore.personData.AddUnitData(unitData)
                    Next
                End If
                progressValue += 1
            Next
            myBW.ReportProgress(0.0#, "beendet.")
            If Not myBW.CancellationPending AndAlso
                ParentDlg.write Then WriteOutputToXlsx.Write(myTemplate, myBW, e, ParentDlg.AllVariables, targetXlsxFilename)
        End If
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
