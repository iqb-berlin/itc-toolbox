Imports Newtonsoft.Json
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Packaging
Imports iqb.lib.openxml
Imports System.ComponentModel

Public Class LoadDataFromTestcenterPage4Results
    Private WithEvents myBackgroundWorker As BackgroundWorker = Nothing

    Private Sub Me_Loaded() Handles Me.Loaded
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
        Me.BtnCancelClose.IsEnabled = False

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
    End Sub

    Private Sub myBackgroundWorker_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles myBackgroundWorker.DoWork
        Dim myBW As BackgroundWorker = sender
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
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
            myBW.ReportProgress(3.0#, "Lese Booklets")
            Dim booklets As List(Of BookletDTO) = ParentDlg.itcConnection.getBooklets()
            Dim bookletSizes As Dictionary(Of String, Long) = (From b As BookletDTO In booklets).ToDictionary(Of String, Long)(Function(b) b.id, Function(b) b.info.totalSize)

            Dim myTestPersonList As New TestPersonList
            Dim AllPeople As New Dictionary(Of String, Dictionary(Of String, List(Of UnitLineData))) 'id -> booklet -> entries
            Dim AllVariables As New List(Of String)
            Dim AllUnitsWithResponses As New List(Of String)
            Dim multiplePersonAndUnits As New List(Of String)
            Dim LogEntryCount As Long = 0

            Dim maxProgressValue As Integer = ParentDlg.selectedDataGroups.Count * 2
            Dim progressValue As Integer = 0
            For Each dataGroupId As String In ParentDlg.selectedDataGroups
                myBW.ReportProgress(progressValue * 100 / maxProgressValue, "Lese '" + dataGroupId + "': ")
                Dim logData As List(Of LogEntryDTO) = ParentDlg.itcConnection.getLogs(dataGroupId)
                If Not String.IsNullOrEmpty(ParentDlg.itcConnection.lastErrorMsgText) Then
                    myBW.ReportProgress(progressValue * 100 / maxProgressValue, "e: Problem bei Logingruppe '" + dataGroupId + "': " +
                        ParentDlg.itcConnection.lastErrorMsgText + " (Logs)")
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

                        If key = "LOADCOMPLETE" Then
                            Dim sysdata As Dictionary(Of String, String) = Nothing
                            parameter = parameter.Replace("\""", """")
                            Try
                                sysdata = JsonConvert.DeserializeObject(parameter, GetType(Dictionary(Of String, String)))
                            Catch ex As Exception
                                sysdata = Nothing
                                Debug.Print("sysdata json convert failed: " + ex.Message)
                            End Try
                            myTestPersonList.SetSysdata(log.timestamp, log.groupname, log.loginname, log.code, log.bookletname, sysdata)
                        End If
                        myTestPersonList.AddLogEvent(log.groupname, log.loginname, log.code, log.bookletname, log.timestamp, log.unitname, key, parameter)
                    Next
                End If
                progressValue += 1

                myBW.ReportProgress(progressValue * 100 / maxProgressValue)
                Dim responseDataList As List(Of ResponseDTO) = ParentDlg.itcConnection.getResponses(dataGroupId)
                If Not String.IsNullOrEmpty(ParentDlg.itcConnection.lastErrorMsgText) Then
                    myBW.ReportProgress(progressValue * 100 / maxProgressValue, "e: Problem bei Logingruppe '" + dataGroupId + "': " +
                    ParentDlg.itcConnection.lastErrorMsgText + " (Responses)")
                Else
                    For Each responseData As ResponseDTO In responseDataList
                        Dim unitData As UnitLineData = UnitLineData.fromTestcenterAPI(responseData)
                        If unitData.hasResponses Then
                            If Not AllUnitsWithResponses.Contains(unitData.unitname) Then AllUnitsWithResponses.Add(unitData.unitname)
                            For Each entry As KeyValuePair(Of String, List(Of ResponseData)) In unitData.responses
                                For Each respData As ResponseData In entry.Value
                                    If Not AllVariables.Contains(unitData.unitname + "##" + respData.variableId) Then AllVariables.Add(unitData.unitname + "##" + respData.variableId)
                                Next
                            Next
                            If Not AllPeople.ContainsKey(unitData.personKey) Then AllPeople.Add(unitData.personKey, New Dictionary(Of String, List(Of UnitLineData)))
                            Dim myPerson As Dictionary(Of String, List(Of UnitLineData)) = AllPeople.Item(unitData.personKey)
                            If Not myPerson.ContainsKey(unitData.bookletname) Then myPerson.Add(unitData.bookletname, New List(Of UnitLineData))
                            Dim myBooklet As List(Of UnitLineData) = myPerson.Item(unitData.bookletname)
                            Dim myUnit As UnitLineData = (From u As UnitLineData In myBooklet Where u.unitname = unitData.unitname).FirstOrDefault
                            If myUnit Is Nothing Then
                                myBooklet.Add(unitData)
                            Else
                                multiplePersonAndUnits.Add(myUnit.groupname + " / " + myUnit.loginname + " / " + myUnit.code + " / " + myUnit.unitname)
                            End If
                        End If
                    Next
                End If
                progressValue += 1
            Next
            If multiplePersonAndUnits.Count > 0 Then
                Dim warningMessage As String = "w: Achtung: In " + multiplePersonAndUnits.Count.ToString +
                        " Fällen wurden mehrere Einträge pro Person und Unit gefunden. Nur der jeweils erste Eintrag wurde übernommen (ignoriere Zeilen "
                If multiplePersonAndUnits.Count > 20 Then
                    warningMessage += String.Join(", ", multiplePersonAndUnits.GetRange(0, 19)) + ", ... )."
                Else
                    warningMessage += String.Join(", ", multiplePersonAndUnits) + ")."
                End If
                myBW.ReportProgress(progressValue * 100 / maxProgressValue, warningMessage)
            End If

            myBW.ReportProgress(0.0#, "Daten für " + AllPeople.Count.ToString("#,##0") + " Testpersonen und " + AllVariables.Count.ToString("#,##0") + " Variablen gelesen.")
            myBW.ReportProgress(0.0#, LogEntryCount.ToString("#,##0") + " Log-Einträge gelesen.")

            If Not myBW.CancellationPending Then WriteOutputToXlsx.Write(myTemplate, myBW, e, AllVariables, AllPeople, myTestPersonList,
                                                                         bookletSizes, targetXlsxFilename)
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
        ParentDlg.DialogResult = False
    End Sub

End Class
