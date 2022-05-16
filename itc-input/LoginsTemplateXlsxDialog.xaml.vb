Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Packaging
Imports iqb.lib.openxml

Public Class LoginsTemplateXlsxDialog
    Private ErrorMessages As Dictionary(Of String, List(Of String))
    Private Warnings As Dictionary(Of String, List(Of String))

    Private Testgroups As Dictionary(Of String, groupdata)

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
        End If
    End Sub

    Private Sub BtnContinue_Click() Handles BtnContinue.Click
        ErrorMessages.Clear()
        Warnings.Clear()
        BtnContinue.Visibility = Windows.Visibility.Collapsed
        BtnCancel.Visibility = Windows.Visibility.Visible
        DPParameters.IsEnabled = False
        Process2_bw = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
        Process2_bw.RunWorkerAsync()
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
                            Dim groupID As String = xlsxFactory.GetCellValue(sourceXLS, groupSheetName, groupNumberIDCol + Zeile.ToString)
                            If String.IsNullOrEmpty(groupID) OrElse Testgroups.ContainsKey(groupID) Then
                                Testgroups.Add(GetGroupId(groupName1), New groupdata With {
                                    .id = groupID, .name1 = groupName1, .name2 = groupName2, .numberLogins = numberTest,
                                    .numberLoginsPlus = numberTestPlus, .numberReviews = numberReview
                                })
                            Else
                                Testgroups.Add(groupID, New groupdata With {
                                    .id = groupID, .name1 = groupName1, .name2 = groupName2, .numberLogins = numberTest,
                                    .numberLoginsPlus = numberTestPlus, .numberReviews = numberReview
                                })
                            End If
                        End If
                        Zeile += 1
                    Loop Until String.IsNullOrEmpty(groupName1)
                    myworker.ReportProgress(20.0#, Testgroups.Count.ToString + " Gruppen gefunden")
                End Using
            End Using
        End If
    End Sub

    '######################################################################################
    '######################################################################################
    Private Sub Process2_bw_DoWork(ByVal sender As Object, ByVal e As ComponentModel.DoWorkEventArgs) Handles Process2_bw.DoWork
        Dim myworker As ComponentModel.BackgroundWorker = sender

    End Sub

    Private Function GetGroupId(name1 As String) As String
        Dim rx As New System.Text.RegularExpressions.Regex("[A-Z]")
        Dim UpperCases As String = ""
        For Each rxMatch As System.Text.RegularExpressions.Match In rx.Matches(name1)
            UpperCases += rxMatch.Value
        Next
        Dim index As Integer = 1
        Dim checkString As String = UpperCases
        Do While Testgroups.ContainsKey(checkString)
            index += 1
            checkString = UpperCases + index.ToString
        Loop
        Return checkString
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
    Public Function toXml(titlePrefix As String, labelShort As String) As XElement
        If String.IsNullOrEmpty(labelShort) Then
            Return <Unit id=<%= id %> label=<%= titlePrefix + name1 %>/>
        Else
            Return <Unit id=<%= id %> label=<%= titlePrefix + name1 %> labelshort=<%= labelShort %>/>
        End If
    End Function
End Class

'#############################################################################
Public Class logindata
    Public login As String
    Public password As String = ""
    Public mode As String = "run-hot-return"

    Public Function toXml(codeToEnter As String, codeToEnterPrompt As String, timeMax As String) As XElement
        Dim myreturn As XElement = <Testlet id=<%= login %> label=<%= password %>/>

        Return myreturn
    End Function

End Class