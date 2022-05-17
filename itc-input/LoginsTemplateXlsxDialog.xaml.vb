Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Packaging
Imports iqb.lib.openxml
Imports iqb.lib.components

Public Class LoginsTemplateXlsxDialog
    Private ErrorMessages As Dictionary(Of String, List(Of String))
    Private Warnings As Dictionary(Of String, List(Of String))

    Private Testgroups As Dictionary(Of String, groupdata)
    Private loginCount As Integer
    Private renewLogins As Boolean = False
    Private generateXml As Boolean = False
    Private Shared loginLength As Integer = 5
    Private Shared passwordLength As Integer = 0
    Private Shared addProctor As Boolean = False
    Private Shared addPrefixTestee As Boolean = False
    Private Shared addPrefixReview As Boolean = False
    Private Shared addPrefixPlus As Boolean = False

#Region "Vorspann"
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Me_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        BtnClose.Visibility = Windows.Visibility.Collapsed
        BtnEditor.Visibility = Windows.Visibility.Collapsed
        BtnContinue.Visibility = Windows.Visibility.Collapsed
        DPParameters.Visibility = Windows.Visibility.Collapsed

        ErrorMessages = New Dictionary(Of String, List(Of String))
        Warnings = New Dictionary(Of String, List(Of String))
        Testgroups = New Dictionary(Of String, groupdata)

        Process1_bw = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
        Process1_bw.RunWorkerAsync()
    End Sub

    Private WithEvents Process1_bw As ComponentModel.BackgroundWorker = Nothing
    Private WithEvents Process2_bw As ComponentModel.BackgroundWorker = Nothing

    Private Sub BtnCancel_Click() Handles BtnCancel.Click
        If Process1_bw IsNot Nothing AndAlso Process1_bw.IsBusy Then
            Process1_bw.CancelAsync()
            BtnCancel.IsEnabled = False
        ElseIf Process2_bw IsNot Nothing AndAlso Process2_bw.IsBusy Then
            Process2_bw.CancelAsync()
            BtnCancel.IsEnabled = False
        Else
            DialogResult = False
        End If
    End Sub

    Private Sub BtnClose_Click() Handles BtnClose.Click
        DialogResult = False
    End Sub

    Private Sub bw_ProgressChanged(ByVal sender As Object, ByVal e As ComponentModel.ProgressChangedEventArgs) Handles Process1_bw.ProgressChanged, Process2_bw.ProgressChanged
        Me.APBUC.UpdateProgressState(e.ProgressPercentage)
        If Not String.IsNullOrEmpty(e.UserState) Then MBUC.AddMessage(e.UserState)
    End Sub
    Private Sub AddWarningMessage(MessageStr As String, MessageParameter As String)
        If Not Warnings.ContainsKey(MessageStr) Then Warnings.Add(MessageStr, New List(Of String))
        Warnings.Item(MessageStr).Add(MessageParameter)
    End Sub
    Private Sub AddErrorMessage(MessageStr As String, MessageParameter As String)
        If Not ErrorMessages.ContainsKey(MessageStr) Then ErrorMessages.Add(MessageStr, New List(Of String))
        ErrorMessages.Item(MessageStr).Add(MessageParameter)
    End Sub

    Private Sub BtnEditor_Click() Handles BtnEditor.Click
        Try
            Dim txtFN As String = IO.Path.GetTempPath + IO.Path.DirectorySeparatorChar + "TestCenter" + Guid.NewGuid.ToString + ".txt"
            IO.File.WriteAllBytes(txtFN, System.Text.Encoding.Unicode.GetBytes(MBUC.Text))

            Dim proc As New Process
            With proc.StartInfo
                .FileName = txtFN
                .WindowStyle = ProcessWindowStyle.Normal
            End With
            proc.Start()

            Me.DialogResult = True
        Catch ex As Exception
            Dim msg As String = ex.Message
            If ex.InnerException IsNot Nothing Then msg += vbNewLine + ex.InnerException.Message
            iqb.lib.components.DialogFactory.MsgError(iqb.lib.components.DialogFactory.GetParentWindow(Me), "Übertragen Meldungen in Texteditor", msg)
        End Try
    End Sub

    Private Sub Process1_bw_RunWorkerCompleted(ByVal sender As Object, ByVal e As ComponentModel.RunWorkerCompletedEventArgs) Handles Process1_bw.RunWorkerCompleted
        APBUC.Value = 0.0#

        If ErrorMessages.Count > 0 Then
            MBUC.AddMessage("h: Fehler")
            For Each eMsg As KeyValuePair(Of String, List(Of String)) In ErrorMessages
                MBUC.AddMessage("e: " + eMsg.Key)
                For Each m As String In From s As String In eMsg.Value Order By s
                    MBUC.AddMessage("e: " + vbTab + m)
                Next
            Next
        End If

        If Warnings.Count > 0 Then
            MBUC.AddMessage("h: Warnungen")
            For Each wMsg As KeyValuePair(Of String, List(Of String)) In Warnings
                MBUC.AddMessage("w: " + wMsg.Key)
                For Each m As String In From s As String In wMsg.Value Order By s
                    MBUC.AddMessage("w: " + vbTab + m)
                Next
            Next
        End If


        MBUC.AddMessage("beendet")
        BtnCancel.Visibility = Windows.Visibility.Collapsed

        If e.Cancelled Then MBUC.AddMessage("durch Nutzer abgebrochen.")

        BtnClose.Visibility = Windows.Visibility.Visible
        BtnEditor.Visibility = Windows.Visibility.Visible

        If Me.Testgroups.Count > 0 Then
            BtnContinue.Visibility = Windows.Visibility.Visible
            DPParameters.Visibility = Windows.Visibility.Visible
            ChBprefixRs.IsChecked = addPrefixPlus
            ChBprefixRv.IsChecked = addPrefixReview
            ChBproctor.IsChecked = addProctor
            ChBprefixT.IsChecked = addPrefixTestee
            TBCharNumberLogin.Text = loginLength.ToString
            TBCharNumberPassword.Text = passwordLength.ToString
            If loginCount = 0 Then ChBnew.IsChecked = True
        End If
    End Sub

    Private Sub BtnContinue_Click() Handles BtnContinue.Click
        If ChBnew.IsChecked OrElse ChBxml.IsChecked Then
            addPrefixPlus = ChBprefixRs.IsChecked
            addPrefixReview = ChBprefixRv.IsChecked
            addProctor = ChBproctor.IsChecked
            addPrefixTestee = ChBprefixT.IsChecked
            Integer.TryParse(TBCharNumberLogin.Text, loginLength)
            Integer.TryParse(TBCharNumberPassword.Text, passwordLength)
            renewLogins = ChBnew.IsChecked
            generateXml = ChBxml.IsChecked
            If renewLogins And loginLength < 5 Then
                DialogFactory.MsgError(Me, "Erzeugen Logins", "Die Länge des Benutzernamens muss mindestens 5 sein.")
            Else
                Testgroups.Clear()
                ErrorMessages.Clear()
                Warnings.Clear()
                BtnContinue.Visibility = Windows.Visibility.Collapsed
                BtnCancel.Visibility = Windows.Visibility.Visible
                DPParameters.IsEnabled = False
                Process2_bw = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
                Process2_bw.RunWorkerAsync()
            End If
        Else
            DialogResult = False
        End If
    End Sub

    Private Sub Process2_bw_RunWorkerCompleted(ByVal sender As Object, ByVal e As ComponentModel.RunWorkerCompletedEventArgs) Handles Process2_bw.RunWorkerCompleted
        APBUC.Value = 0.0#

        If ErrorMessages.Count > 0 Then
            MBUC.AddMessage("h: Fehler")
            For Each eMsg As KeyValuePair(Of String, List(Of String)) In ErrorMessages
                MBUC.AddMessage("e: " + eMsg.Key)
                For Each m As String In From s As String In eMsg.Value Order By s
                    MBUC.AddMessage("e: " + vbTab + m)
                Next
            Next
        End If

        If Warnings.Count > 0 Then
            MBUC.AddMessage("h: Warnungen")
            For Each wMsg As KeyValuePair(Of String, List(Of String)) In Warnings
                MBUC.AddMessage("w: " + wMsg.Key)
                For Each m As String In From s As String In wMsg.Value Order By s
                    MBUC.AddMessage("w: " + vbTab + m)
                Next
            Next
        End If


        MBUC.AddMessage("beendet")
        BtnCancel.Visibility = Windows.Visibility.Collapsed

        If e.Cancelled Then MBUC.AddMessage("durch Nutzer abgebrochen.")

        BtnClose.Visibility = Windows.Visibility.Visible
        BtnEditor.Visibility = Windows.Visibility.Visible
    End Sub
#End Region

    '######################################################################################
    '######################################################################################
    Private Sub Process1_bw_DoWork(ByVal sender As Object, ByVal e As ComponentModel.DoWorkEventArgs) Handles Process1_bw.DoWork
        Dim myworker As ComponentModel.BackgroundWorker = sender

        myworker.ReportProgress(20.0#, "Öffne Datei")
        Dim sourceFileName = My.Settings.lastfile_LoginXlsx

        Dim sourceFile As Byte()
        Try
            sourceFile = IO.File.ReadAllBytes(sourceFileName)
        Catch ex As Exception
            myworker.ReportProgress(20.0#, "e:Konnte Datei " + sourceFileName + " nicht lesen (noch geöffnet?)")
            sourceFile = Nothing
        End Try

        If sourceFile IsNot Nothing Then
            Using MemStream As New IO.MemoryStream()
                MemStream.Write(sourceFile, 0, sourceFile.Length)
                Using sourceXLS As SpreadsheetDocument = SpreadsheetDocument.Open(MemStream, False)
                    myworker.ReportProgress(20.0#, "Lese Datei " + sourceFileName)

                    Dim groupName1Ref As String = xlsxFactory.GetDefinedNameValue(sourceXLS, "groupsCol.Name1")
                    Dim groupSheetName As String = xlsxFactory.GetWorksheetNameFromRefStr(groupName1Ref)
                    Dim groupName1Col As String = xlsxFactory.GetColumnFromRefStr(groupName1Ref)
                    Dim groupsFirstRow As Integer = xlsxFactory.GetRowFromRefStr(groupName1Ref) + 1
                    Dim groupName2Col As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "groupsCol.Name2"))
                    Dim groupNumberTestCol As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "groupsCol.NumberTest"))
                    Dim groupNumberPlusCol As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "groupsCol.NumberPlus"))
                    Dim groupNumberReviewCol As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "groupsCol.NumberReview"))
                    Dim groupNumberIDCol As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "groupsCol.ID"))

                    Dim loginsName1Ref As String = xlsxFactory.GetDefinedNameValue(sourceXLS, "loginsCol.GroupName1")
                    Dim loginsSheetName As String = xlsxFactory.GetWorksheetNameFromRefStr(loginsName1Ref)
                    Dim loginsName1Col As String = xlsxFactory.GetColumnFromRefStr(loginsName1Ref)
                    Dim loginsFirstRow As Integer = xlsxFactory.GetRowFromRefStr(loginsName1Ref) + 1
                    Dim loginsName2Col As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "loginsCol.GroupName2"))
                    Dim loginsGroupIdCol As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "loginsCol.GroupID"))
                    Dim loginsLoginCol As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "loginsCol.Name"))
                    Dim loginsPasswordCol As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "loginsCol.Password"))
                    Dim loginsModeCol As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "loginsCol.Mode"))

                    '-----------------------------------------------------------------------------
                    '-----------------------------------------------------------------------------
                    myworker.ReportProgress(20.0#, "h:Lese Gruppen")

                    Dim Zeile As Integer = groupsFirstRow
                    Dim groupName1 As String = ""
                    Dim fatal_error As Boolean = False
                    Do
                        If myworker.CancellationPending() Then Exit Do
                        groupName1 = xlsxFactory.GetCellValue(sourceXLS, groupSheetName, groupName1Col + Zeile.ToString)
                        If Not String.IsNullOrEmpty(groupName1) Then
                            Dim groupName2 As String = xlsxFactory.GetCellValue(sourceXLS, groupSheetName, groupName2Col + Zeile.ToString)
                            Dim numberTest As Integer = 0
                            Dim numberTestStr As String = xlsxFactory.GetCellValue(sourceXLS, groupSheetName, groupNumberTestCol + Zeile.ToString)
                            If Not String.IsNullOrEmpty(numberTestStr) Then Integer.TryParse(numberTestStr, numberTest)
                            Dim numberTestPlus As Integer = 0
                            Dim numberTestPlusStr As String = xlsxFactory.GetCellValue(sourceXLS, groupSheetName, groupNumberPlusCol + Zeile.ToString)
                            If Not String.IsNullOrEmpty(numberTestPlusStr) Then Integer.TryParse(numberTestPlusStr, numberTestPlus)
                            Dim numberReview As Integer = 0
                            Dim numberReviewStr As String = xlsxFactory.GetCellValue(sourceXLS, groupSheetName, groupNumberReviewCol + Zeile.ToString)
                            If Not String.IsNullOrEmpty(numberReviewStr) Then Integer.TryParse(numberReviewStr, numberReview)
                            Dim givenGroupID As String = xlsxFactory.GetCellValue(sourceXLS, groupSheetName, groupNumberIDCol + Zeile.ToString)
                            If String.IsNullOrEmpty(givenGroupID) OrElse Testgroups.ContainsKey(givenGroupID) Then
                                Testgroups.Add(GetGroupId(groupName1, ""), New groupdata With {
                                    .id = givenGroupID, .name1 = groupName1, .name2 = groupName2, .numberLogins = numberTest,
                                    .numberLoginsPlus = numberTestPlus, .numberReviews = numberReview
                                })
                            Else
                                Testgroups.Add(givenGroupID, New groupdata With {
                                    .id = givenGroupID, .name1 = groupName1, .name2 = groupName2, .numberLogins = numberTest,
                                    .numberLoginsPlus = numberTestPlus, .numberReviews = numberReview
                                })
                            End If
                        End If
                        Zeile += 1
                    Loop Until String.IsNullOrEmpty(groupName1)
                    myworker.ReportProgress(20.0#, Testgroups.Count.ToString + " Gruppen gefunden")

                    Zeile = loginsFirstRow
                    Dim groupID As String = ""
                    Dim login As String = ""
                    loginCount = 0
                    fatal_error = False
                    Do
                        If myworker.CancellationPending() Then Exit Do
                        Try
                            groupID = xlsxFactory.GetCellValue(sourceXLS, loginsSheetName, loginsGroupIdCol + Zeile.ToString)
                            login = xlsxFactory.GetCellValue(sourceXLS, loginsSheetName, loginsLoginCol + Zeile.ToString)
                        Catch ex As Exception
                            groupID = ""
                            login = ""
                        End Try
                        If Not String.IsNullOrEmpty(groupID) AndAlso Not String.IsNullOrEmpty(login) Then
                            loginCount += 1
                        End If
                        Zeile += 1
                    Loop Until String.IsNullOrEmpty(groupID) OrElse String.IsNullOrEmpty(login)
                    myworker.ReportProgress(20.0#, loginCount.ToString + " Logins gefunden")
                End Using
            End Using
        End If
    End Sub

    '######################################################################################
    '######################################################################################
    Private Sub Process2_bw_DoWork(ByVal sender As Object, ByVal e As ComponentModel.DoWorkEventArgs) Handles Process2_bw.DoWork
        Dim myworker As ComponentModel.BackgroundWorker = sender
        myworker.ReportProgress(20.0#, "Öffne Datei")
        Dim sourceFileName = My.Settings.lastfile_LoginXlsx

        Dim sourceFile As Byte()
        Try
            sourceFile = IO.File.ReadAllBytes(sourceFileName)
        Catch ex As Exception
            myworker.ReportProgress(20.0#, "e:Konnte Datei " + sourceFileName + " nicht lesen (noch geöffnet?)")
            sourceFile = Nothing
        End Try

        If sourceFile IsNot Nothing Then
            Using MemStream As New IO.MemoryStream()
                MemStream.Write(sourceFile, 0, sourceFile.Length)
                Using sourceXLS As SpreadsheetDocument = SpreadsheetDocument.Open(MemStream, True)
                    myworker.ReportProgress(5.0#, "Lese Datei " + sourceFileName)

                    Dim groupName1Ref As String = xlsxFactory.GetDefinedNameValue(sourceXLS, "groupsCol.Name1")
                    Dim groupSheetName As String = xlsxFactory.GetWorksheetNameFromRefStr(groupName1Ref)
                    Dim groupSheet As WorksheetPart = xlsxFactory.GetWorksheetPart(sourceXLS, groupSheetName)
                    Dim groupName1Col As String = xlsxFactory.GetColumnFromRefStr(groupName1Ref)
                    Dim groupsFirstRow As Integer = xlsxFactory.GetRowFromRefStr(groupName1Ref) + 1
                    Dim groupName2Col As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "groupsCol.Name2"))
                    Dim groupNumberTestCol As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "groupsCol.NumberTest"))
                    Dim groupNumberPlusCol As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "groupsCol.NumberPlus"))
                    Dim groupNumberReviewCol As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "groupsCol.NumberReview"))
                    Dim groupNumberIDCol As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "groupsCol.ID"))

                    Dim loginsName1Ref As String = xlsxFactory.GetDefinedNameValue(sourceXLS, "loginsCol.GroupName1")
                    Dim loginsSheetName As String = xlsxFactory.GetWorksheetNameFromRefStr(loginsName1Ref)
                    Dim loginsName1Col As String = xlsxFactory.GetColumnFromRefStr(loginsName1Ref)
                    Dim loginsFirstRow As Integer = xlsxFactory.GetRowFromRefStr(loginsName1Ref) + 1
                    Dim loginsName2Col As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "loginsCol.GroupName2"))
                    Dim loginsGroupIdCol As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "loginsCol.GroupID"))
                    Dim loginsLoginCol As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "loginsCol.Name"))
                    Dim loginsPasswordCol As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "loginsCol.Password"))
                    Dim loginsModeCol As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "loginsCol.Mode"))
                    Dim xlsxChanged As Boolean = False
                    '-----------------------------------------------------------------------------
                    '-----------------------------------------------------------------------------
                    myworker.ReportProgress(10.0#, "h:Lese Gruppen")

                    Dim groupIdSuffix As String = ":" + LoginsXlsxDialog.GetNewCode(4)
                    Dim Zeile As Integer = groupsFirstRow
                    Dim groupName1 As String = ""
                    Dim groupName2 As String = ""
                    Dim fatal_error As Boolean = False
                    Dim loginSum As Integer = 0
                    Do
                        If myworker.CancellationPending() Then Exit Do
                        groupName1 = xlsxFactory.GetCellValue(sourceXLS, groupSheetName, groupName1Col + Zeile.ToString)
                        If Not String.IsNullOrEmpty(groupName1) Then
                            groupName2 = xlsxFactory.GetCellValue(sourceXLS, groupSheetName, groupName2Col + Zeile.ToString)
                            Dim numberTest As Integer = 0
                            Dim numberTestStr As String = xlsxFactory.GetCellValue(sourceXLS, groupSheetName, groupNumberTestCol + Zeile.ToString)
                            If Not String.IsNullOrEmpty(numberTestStr) Then Integer.TryParse(numberTestStr, numberTest)
                            Dim numberTestPlus As Integer = 0
                            Dim numberTestPlusStr As String = xlsxFactory.GetCellValue(sourceXLS, groupSheetName, groupNumberPlusCol + Zeile.ToString)
                            If Not String.IsNullOrEmpty(numberTestPlusStr) Then Integer.TryParse(numberTestPlusStr, numberTestPlus)
                            Dim numberReview As Integer = 0
                            Dim numberReviewStr As String = xlsxFactory.GetCellValue(sourceXLS, groupSheetName, groupNumberReviewCol + Zeile.ToString)
                            If Not String.IsNullOrEmpty(numberReviewStr) Then Integer.TryParse(numberReviewStr, numberReview)
                            Dim givenGroupID As String = xlsxFactory.GetCellValue(sourceXLS, groupSheetName, groupNumberIDCol + Zeile.ToString)
                            If String.IsNullOrEmpty(givenGroupID) OrElse Testgroups.ContainsKey(givenGroupID) Then
                                Dim newGroupId As String = GetGroupId(groupName1, groupIdSuffix)
                                xlsxFactory.SetCellValueString(groupNumberIDCol, Zeile, groupSheet, newGroupId)
                                xlsxChanged = True
                                Testgroups.Add(newGroupId, New groupdata With {
                                    .id = newGroupId, .name1 = groupName1, .name2 = groupName2, .numberLogins = numberTest,
                                    .numberLoginsPlus = numberTestPlus, .numberReviews = numberReview
                                })
                            Else
                                Testgroups.Add(givenGroupID, New groupdata With {
                                    .id = givenGroupID, .name1 = groupName1, .name2 = groupName2, .numberLogins = numberTest,
                                    .numberLoginsPlus = numberTestPlus, .numberReviews = numberReview
                                })
                            End If
                            loginSum += numberTest
                            loginSum += numberReview
                            loginSum += numberTestPlus
                        End If
                        Zeile += 1
                    Loop Until String.IsNullOrEmpty(groupName1)
                    If xlsxChanged Then groupSheet.Worksheet.Save()

                    If renewLogins Then
                        myworker.ReportProgress(10.0#, "Generiere Logins")
                        If addProctor Then loginSum += Testgroups.Count
                        Dim allLogins As List(Of String) = LoginsXlsxDialog.GetNewCodeList(loginLength, loginSum)
                        Dim loginIndex As Integer = 0
                        Dim allPasswords As New List(Of String)
                        Dim passwordIndex As Integer = 0
                        If passwordLength > 0 Then
                            Dim pwMax As Integer = 150
                            If passwordLength > 2 Then pwMax = 1500
                            If passwordLength > 3 Then pwMax = loginSum
                            allPasswords = LoginsXlsxDialog.GetNewCodeList(passwordLength, pwMax)
                        End If

                        Dim usedColumns As New List(Of String)
                        usedColumns.Add(loginsName1Col)
                        usedColumns.Add(loginsName2Col)
                        usedColumns.Add(loginsGroupIdCol)
                        usedColumns.Add(loginsLoginCol)
                        usedColumns.Add(loginsPasswordCol)
                        usedColumns.Add(loginsModeCol)
                        Dim otherColumns As New List(Of String)
                        Dim columnToAdd As String = "A"
                        Do While columnToAdd <> "M"
                            If Not usedColumns.Contains(columnToAdd) Then otherColumns.Add(columnToAdd)
                            columnToAdd = xlsxFactory.GetNextColumn(columnToAdd)
                        Loop

                        Zeile = loginsFirstRow
                        Dim loginsSheet As WorksheetPart = xlsxFactory.GetWorksheetPart(sourceXLS, loginsSheetName)
                        Dim progressMax As Integer = Testgroups.Count
                        Dim progressCurrent As Integer = 0
                        For Each group As KeyValuePair(Of String, groupdata) In Testgroups
                            If myworker.CancellationPending() Then Exit For
                            myworker.ReportProgress(10.0# + 90 * progressCurrent / progressMax)
                            progressCurrent += 1
                            For index = 1 To group.Value.numberLogins
                                xlsxFactory.SetCellValueString(loginsName1Col, Zeile, loginsSheet, group.Value.name1)
                                xlsxFactory.SetCellValueString(loginsName2Col, Zeile, loginsSheet, group.Value.name2)
                                xlsxFactory.SetCellValueString(loginsGroupIdCol, Zeile, loginsSheet, group.Value.id)
                                xlsxFactory.SetCellValueString(loginsLoginCol, Zeile, loginsSheet, IIf(addPrefixTestee, ":" + index.ToString("00") + ":", "") + allLogins(loginIndex))
                                loginIndex += 1
                                If passwordLength > 0 Then
                                    xlsxFactory.SetCellValueString(loginsPasswordCol, Zeile, loginsSheet, allPasswords(passwordIndex))
                                    passwordIndex += 1
                                    If passwordIndex >= allPasswords.Count Then passwordIndex = 0
                                Else
                                    xlsxFactory.SetCellValueString(loginsPasswordCol, Zeile, loginsSheet, "")
                                End If
                                xlsxFactory.SetCellValueString(loginsModeCol, Zeile, loginsSheet, "run-hot-return")
                                For Each c As String In otherColumns
                                    xlsxFactory.SetCellValueString(c, Zeile, loginsSheet, "")
                                Next
                                Zeile += 1
                            Next
                            For index = 1 To group.Value.numberLoginsPlus
                                xlsxFactory.SetCellValueString(loginsName1Col, Zeile, loginsSheet, group.Value.name1)
                                xlsxFactory.SetCellValueString(loginsName2Col, Zeile, loginsSheet, group.Value.name2)
                                xlsxFactory.SetCellValueString(loginsGroupIdCol, Zeile, loginsSheet, group.Value.id)
                                xlsxFactory.SetCellValueString(loginsLoginCol, Zeile, loginsSheet, IIf(addPrefixPlus, ":RS:", "") + allLogins(loginIndex))
                                loginIndex += 1
                                If passwordLength > 0 Then
                                    xlsxFactory.SetCellValueString(loginsPasswordCol, Zeile, loginsSheet, allPasswords(passwordIndex))
                                    passwordIndex += 1
                                    If passwordIndex >= allPasswords.Count Then passwordIndex = 0
                                Else
                                    xlsxFactory.SetCellValueString(loginsPasswordCol, Zeile, loginsSheet, "")
                                End If
                                xlsxFactory.SetCellValueString(loginsModeCol, Zeile, loginsSheet, "run-hot-return")
                                For Each c As String In otherColumns
                                    xlsxFactory.SetCellValueString(c, Zeile, loginsSheet, "")
                                Next
                                Zeile += 1
                            Next
                            For index = 1 To group.Value.numberReviews
                                xlsxFactory.SetCellValueString(loginsName1Col, Zeile, loginsSheet, group.Value.name1)
                                xlsxFactory.SetCellValueString(loginsName2Col, Zeile, loginsSheet, group.Value.name2)
                                xlsxFactory.SetCellValueString(loginsGroupIdCol, Zeile, loginsSheet, group.Value.id)
                                xlsxFactory.SetCellValueString(loginsLoginCol, Zeile, loginsSheet, IIf(addPrefixReview, ":RV:", "") + allLogins(loginIndex))
                                loginIndex += 1
                                If passwordLength > 0 Then
                                    xlsxFactory.SetCellValueString(loginsPasswordCol, Zeile, loginsSheet, allPasswords(passwordIndex))
                                    passwordIndex += 1
                                    If passwordIndex >= allPasswords.Count Then passwordIndex = 0
                                Else
                                    xlsxFactory.SetCellValueString(loginsPasswordCol, Zeile, loginsSheet, "")
                                End If
                                xlsxFactory.SetCellValueString(loginsModeCol, Zeile, loginsSheet, "run-review")
                                For Each c As String In otherColumns
                                    xlsxFactory.SetCellValueString(c, Zeile, loginsSheet, "")
                                Next
                                Zeile += 1
                            Next
                            If addProctor Then
                                xlsxFactory.SetCellValueString(loginsName1Col, Zeile, loginsSheet, group.Value.name1)
                                xlsxFactory.SetCellValueString(loginsName2Col, Zeile, loginsSheet, group.Value.name2)
                                xlsxFactory.SetCellValueString(loginsGroupIdCol, Zeile, loginsSheet, group.Value.id)
                                xlsxFactory.SetCellValueString(loginsLoginCol, Zeile, loginsSheet, ":TL:" + allLogins(loginIndex))
                                loginIndex += 1
                                If passwordLength > 0 Then
                                    xlsxFactory.SetCellValueString(loginsPasswordCol, Zeile, loginsSheet, allPasswords(passwordIndex))
                                    passwordIndex += 1
                                    If passwordIndex >= allPasswords.Count Then passwordIndex = 0
                                Else
                                    xlsxFactory.SetCellValueString(loginsPasswordCol, Zeile, loginsSheet, "")
                                End If
                                xlsxFactory.SetCellValueString(loginsModeCol, Zeile, loginsSheet, "monitor-group")
                                For Each c As String In otherColumns
                                    xlsxFactory.SetCellValueString(c, Zeile, loginsSheet, "")
                                Next
                                Zeile += 1
                            End If
                        Next
                        Dim groupID As String = ""
                        Dim login As String = ""
                        Do
                            If myworker.CancellationPending() Then Exit Do
                            groupID = xlsxFactory.GetCellValue(sourceXLS, loginsSheetName, loginsGroupIdCol + Zeile.ToString)
                            login = xlsxFactory.GetCellValue(sourceXLS, loginsSheetName, loginsLoginCol + Zeile.ToString)
                            If Not String.IsNullOrEmpty(groupID) OrElse Not String.IsNullOrEmpty(login) Then
                                For Each c As String In otherColumns
                                    xlsxFactory.SetCellValueString(c, Zeile, loginsSheet, "") 'maybe this produces invalid cell content???
                                Next
                                For Each c As String In usedColumns
                                    xlsxFactory.SetCellValueString(c, Zeile, loginsSheet, "") 'maybe this produces invalid cell content???
                                Next
                            End If
                            Zeile += 1
                        Loop Until String.IsNullOrEmpty(groupID) AndAlso String.IsNullOrEmpty(login)
                        loginsSheet.Worksheet.Save()
                    End If
                    '#####################
                    If generateXml Then
                        Dim customTexts As New Dictionary(Of String, String)
                        Dim customTextsKeyRef As String = xlsxFactory.GetDefinedNameValue(sourceXLS, "customText.Key")
                        Dim customTextsSheetName As String = xlsxFactory.GetWorksheetNameFromRefStr(customTextsKeyRef)
                        Dim customTextsKeyCol As String = xlsxFactory.GetColumnFromRefStr(customTextsKeyRef)
                        Dim customTextsFirstRow As Integer = xlsxFactory.GetRowFromRefStr(customTextsKeyRef) + 1
                        Dim customTextsValueCol As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "customText.Value"))
                        Zeile = customTextsFirstRow
                        Dim customTextKey As String = ""
                        Do
                            If myworker.CancellationPending() Then Exit Do
                            Try
                                customTextKey = xlsxFactory.GetCellValue(sourceXLS, customTextsSheetName, customTextsKeyCol + Zeile.ToString)
                            Catch ex As Exception
                                customTextKey = ""
                            End Try
                            If Not String.IsNullOrEmpty(customTextKey) Then
                                Dim customTextValue = xlsxFactory.GetCellValue(sourceXLS, customTextsSheetName, customTextsValueCol + Zeile.ToString)
                                If Not String.IsNullOrEmpty(customTextValue) AndAlso Not customTexts.ContainsKey(customTextKey) Then
                                    customTexts.Add(customTextKey, customTextValue)
                                End If
                            End If
                            Zeile += 1
                        Loop Until String.IsNullOrEmpty(customTextKey)
                        Dim LoginXmlFile As XDocument = <?xml version="1.0" encoding="utf-8"?>
                                                        <Testtakers>
                                                            <Metadata>
                                                            </Metadata>
                                                        </Testtakers>
                        If customTexts.Count > 0 Then
                            Dim customTextsElement As XElement = <CustomTexts></CustomTexts>
                            For Each customText As KeyValuePair(Of String, String) In customTexts
                                customTextsElement.Add(<CustomText key=<%= customText.Key %>><%= customText.Value %></CustomText>)
                            Next
                            LoginXmlFile.Root.Add(customTextsElement)
                        End If

                        Zeile = loginsFirstRow
                        Dim groupID As String = ""
                        Dim login As String = ""
                        Dim password As String = ""
                        Dim loginMode As String = ""

                        loginCount = 0
                        fatal_error = False
                        Do
                            If myworker.CancellationPending() Then Exit Do
                            Try
                                groupID = xlsxFactory.GetCellValue(sourceXLS, loginsSheetName, loginsGroupIdCol + Zeile.ToString)
                                login = xlsxFactory.GetCellValue(sourceXLS, loginsSheetName, loginsLoginCol + Zeile.ToString)
                            Catch ex As Exception
                                groupID = ""
                                login = ""
                            End Try
                            If Not String.IsNullOrEmpty(groupID) AndAlso Not String.IsNullOrEmpty(login) Then
                                If Not Testgroups.ContainsKey(groupID) Then
                                    Try
                                        groupName1 = xlsxFactory.GetCellValue(sourceXLS, loginsSheetName, loginsGroupIdCol + Zeile.ToString)
                                        groupName2 = xlsxFactory.GetCellValue(sourceXLS, loginsSheetName, loginsLoginCol + Zeile.ToString)
                                    Catch ex As Exception
                                        groupName1 = ""
                                        groupName2 = ""
                                    End Try
                                    Testgroups.Add(groupID, New groupdata With {.id = groupID, .name1 = groupName1, .name2 = groupName2})
                                End If
                                Dim myGroup As groupdata = Testgroups.Item(groupID)
                                Try
                                    password = xlsxFactory.GetCellValue(sourceXLS, loginsSheetName, loginsPasswordCol + Zeile.ToString)
                                    loginMode = xlsxFactory.GetCellValue(sourceXLS, loginsSheetName, loginsModeCol + Zeile.ToString)
                                Catch ex As Exception
                                    password = ""
                                    loginMode = ""
                                End Try
                                myGroup.logins.Add(New logindata With {.login = login, .password = password, .mode = loginMode})
                            End If
                            Zeile += 1
                        Loop Until String.IsNullOrEmpty(groupID) OrElse String.IsNullOrEmpty(login)
                        For Each testGroup As KeyValuePair(Of String, groupdata) In Testgroups
                            LoginXmlFile.Root.Add(testGroup.Value.toXml())
                        Next
                        Dim targetFilePath As String = IO.Path.GetDirectoryName(sourceFileName) + IO.Path.DirectorySeparatorChar
                        Dim targetFileName As String = IO.Path.GetFileNameWithoutExtension(sourceFileName) + ".xml"
                        Try
                            LoginXmlFile.Save(targetFilePath + targetFileName)
                            myworker.ReportProgress(0.0#, "Habe Datei '" + targetFileName + "' gespeichert")
                        Catch ex As Exception
                            Dim msg As String = ex.Message
                            If ex.InnerException IsNot Nothing Then msg += vbNewLine + ex.InnerException.Message
                            myworker.ReportProgress(0.0#, "e: Konnte Datei '" + targetFileName + "' nicht schreiben (" + msg + ")")
                        End Try
                    End If
                End Using
                Try
                    Using fs As New IO.FileStream(My.Settings.lastfile_LoginXlsx, IO.FileMode.Create)
                        MemStream.WriteTo(fs)
                    End Using
                Catch ex As Exception
                    myworker.ReportProgress(0.0#, "e: Konnte Datei nicht schreiben: " + ex.Message)
                End Try
            End Using
        End If
    End Sub

    Private Function GetGroupId(name1 As String, suffix As String) As String
        Dim rx As New System.Text.RegularExpressions.Regex("[A-Z]")
        Dim UpperCases As String = ""
        For Each rxMatch As System.Text.RegularExpressions.Match In rx.Matches(name1)
            UpperCases += rxMatch.Value
        Next
        Dim index As Integer = 1
        Dim checkString As String = UpperCases
        Do While Testgroups.ContainsKey(checkString + suffix)
            index += 1
            checkString = UpperCases + index.ToString
        Loop
        Return checkString + suffix
    End Function
End Class

'#############################################################################
Public Class groupdata
    Public id As String = ""
    Public name1 As String = ""
    Public name2 As String = ""
    Public numberLogins As Integer = 0
    Public numberLoginsPlus As Integer = 0
    Public numberReviews As Integer = 0
    Public logins As New List(Of logindata)
    Public Function toXml() As XElement
        Dim myreturn As XElement = <Group id=<%= id %> label=<%= name1 + " - " + name2 %>></Group>
        For Each login As logindata In logins
            myreturn.Add(login.toXml)
        Next
        Return myreturn
    End Function
End Class

'#############################################################################
Public Class logindata
    Public login As String
    Public password As String = ""
    Public mode As String = "run-hot-return"

    Public Function toXml(Optional bookletName As String = "DummyBooklet") As XElement
        Dim myreturn As XElement = <Login mode=<%= mode %> name=<%= login %>>
                                       <Booklet><%= bookletName %></Booklet>
                                   </Login>
        If Not String.IsNullOrEmpty(password) Then myreturn.SetAttributeValue("pw", password)
        Return myreturn
    End Function

End Class