Imports System.Globalization
Public Enum CsvContent
    SESSIONS
    RESPONSES
End Enum
Public Class ToCsvDialog
    Private DB As SQLiteConnector
    Private deCulture As CultureInfo
    Private csvContent As CsvContent
#Region "Vorspann"
    Public Sub New(SqliteDB As SQLiteConnector, mode As CsvContent)
        InitializeComponent()
        DB = SqliteDB
        deCulture = CultureInfo.CreateSpecificCulture("de-DE")
        csvContent = mode
    End Sub

    Private Sub Me_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If DB Is Nothing Then
            BtnCancelClose.Content = "Schließen"
            MBUC.AddMessage("Keine Datenbank geöffnet.")
        Else
            Process1_bw = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
            Process1_bw.RunWorkerAsync()
        End If
    End Sub

    Private WithEvents Process1_bw As ComponentModel.BackgroundWorker = Nothing

    Private Sub BtnCancelClose_Click() Handles BtnCancelClose.Click
        If Process1_bw IsNot Nothing AndAlso Process1_bw.IsBusy Then
            Process1_bw.CancelAsync()
            BtnCancelClose.IsEnabled = False
        Else
            DialogResult = False
        End If
    End Sub

    Private Sub bw_ProgressChanged(ByVal sender As Object, ByVal e As ComponentModel.ProgressChangedEventArgs) Handles Process1_bw.ProgressChanged
        Me.APBUC.UpdateProgressState(e.ProgressPercentage)
        If Not String.IsNullOrEmpty(e.UserState) Then MBUC.AddMessage(e.UserState)
    End Sub

    Private Sub Process1_bw_RunWorkerCompleted(ByVal sender As Object, ByVal e As ComponentModel.RunWorkerCompletedEventArgs) Handles Process1_bw.RunWorkerCompleted
        MBUC.AddMessage("beendet")
        If e.Cancelled Then MBUC.AddMessage("durch Nutzer abgebrochen.")
        Me.APBUC.UpdateProgressState(0.0#)
        BtnCancelClose.IsEnabled = True
        BtnCancelClose.Content = "Schließen"
    End Sub
#End Region

    '######################################################################################
    '######################################################################################
    Private Sub Process1_bw_DoWork(ByVal sender As Object, ByVal e As ComponentModel.DoWorkEventArgs) Handles Process1_bw.DoWork
        Dim worker As ComponentModel.BackgroundWorker = sender
        If Not worker.CancellationPending Then
            worker.ReportProgress(5.0#, "h:Schreibe " + IO.Path.GetFileName(My.Settings.lastfile_OutputTargetCsv) + " - bitte warten")
            Dim peopleListAll As Dictionary(Of Long, String) = DB.getPeopleListAll()
            Dim progressMax As Long = peopleListAll.Count
            Dim progressStep As Long = 1
            Dim totalOfDoubleUnits As Integer = 0
            Dim totalOfPeopleWithDoubleUnits As Integer = 0

            Dim file As System.IO.StreamWriter
            file = My.Computer.FileSystem.OpenTextFileWriter(My.Settings.lastfile_OutputTargetCsv, False)
            file.WriteLine(IIf(csvContent = CsvContent.RESPONSES, firstCsvLineResponses, firstCsvLineSessions))
            For Each person As KeyValuePair(Of Long, String) In From p As KeyValuePair(Of Long, String) In peopleListAll Order By p.Value
                If worker.CancellationPending Then Exit For
                Dim progressValue As Double = 5 + 95 * progressStep / progressMax
                progressStep += 1

                If csvContent = CsvContent.RESPONSES Then
                    For Each r As PersonResponseLong In DB.getPersonResponsesLong(person.Key)
                        file.WriteLine(PersonResponseLongAsCsvString(r))
                    Next
                Else
                    Dim sessions As List(Of SessionReport) = DB.getSessionReports(person.Key)
                    For Each s As SessionReport In sessions
                        file.WriteLine(PersonSessionAsCsvString(s, person.Value, sessions.Count))
                    Next
                End If
                worker.ReportProgress(progressValue)
            Next
            file.Close()
        End If
    End Sub

    Private Const firstCsvLineResponses = "Group;Login;Code;Booklet;Unit-Name;Unit-Alias;Response-Id;Response-Status;Response-Value;Response-Subform;Response-Code;Response-Score;Response-Timestamp;Response-Timestamp-String"
    Private Const firstCsvLineSessions = "Person;Booklet;Session-No;Sessions-Total;Units-w-Value;Start-At-TS;Start-At-DT;First-Responses-After-MS;First-Responses-After-MMSS;Last-Responses-After-MS;Last-Responses-After-MMSS;Load-Speed-MBitPerSec;Browser;OS;Screen;Screen-Width;Screen-Height"

    Private Function PersonResponseLongAsCsvString(r As PersonResponseLong, Optional separator As String = ";") As String
        Dim returnStr As String = """" + r.group + """"
        returnStr += separator + """" + r.login + """"
        returnStr += separator + """" + r.code + """"
        returnStr += separator + """" + r.booklet + """"
        returnStr += separator + """" + r.unitName + """"
        returnStr += separator + """" + r.unitAlias + """"
        returnStr += separator + """" + r.responseId + """"
        returnStr += separator + """" + r.responseStatus + """"
        returnStr += separator + """" + r.responseValue + """"
        returnStr += separator + """" + r.responseSubform + """"
        returnStr += separator + r.responseCode.ToString
        returnStr += separator + r.responseScore.ToString
        returnStr += separator + r.ts
        Dim tsLong As Long
        If Long.TryParse(r.ts, tsLong) Then
            Dim tsDt As New DateTime(1970, 1, 1, 0, 0, 0, 0)
            tsDt = tsDt.AddMilliseconds(tsLong)
            returnStr += separator + """" + tsDt.ToString(deCulture) + """"
        Else
            returnStr += separator + """?"""
        End If

        Return returnStr
    End Function

    Private Function PersonSessionAsCsvString(s As SessionReport, personId As String, sessionsCount As Integer, Optional separator As String = ";") As String
        Dim returnStr As String = """" + personId + """"
        returnStr += separator + """" + s.booklet + """"
        returnStr += separator + s.sessionNumber.ToString
        returnStr += separator + sessionsCount.ToString
        returnStr += separator + """" + String.Join(" ", s.unitsWithResponse) + """"
        returnStr += separator + s.sessionStartTs.ToString
        Dim sessionStart As New DateTime(1970, 1, 1, 0, 0, 0, 0)
        sessionStart = sessionStart.AddMilliseconds(s.sessionStartTs)
        returnStr += separator + """" + sessionStart.ToString(deCulture) + """"
        Dim timespan As Long = Int((s.firstUnitTS - s.sessionStartTs) / 1000)
        returnStr += separator + timespan.ToString
        returnStr += separator + """'" + Fix(timespan / 60).ToString + ":" + (timespan Mod 60).ToString("00") + """"
        timespan = Int((s.lastUnitTS - s.sessionStartTs) / 1000)
        returnStr += separator + timespan.ToString
        returnStr += separator + """'" + Fix(timespan / 60).ToString + ":" + (timespan Mod 60).ToString("00") + """"
        returnStr += separator + """" + s.contentLoadSpeed.ToString(deCulture) + """"
        returnStr += separator + """" + s.browser + """"
        returnStr += separator + """" + s.os + """"
        returnStr += separator + """" + s.screen + """"
        Dim nextString As String = separator + separator
        If Not String.IsNullOrEmpty(s.screen) AndAlso s.screen.IndexOf(" x ") > 0 Then
            Dim numericScreen As String() = s.screen.Split(" ")
            If numericScreen.Length > 2 Then
                Dim firstValue As Integer = 0
                Dim secondValue As Integer = 0
                If Integer.TryParse(numericScreen(0), firstValue) AndAlso Integer.TryParse(numericScreen(2), secondValue) Then
                    If firstValue > secondValue Then
                        nextString = separator + firstValue.ToString + separator + secondValue.ToString
                    Else
                        nextString = separator + secondValue.ToString + separator + firstValue.ToString
                    End If
                End If
            End If
        End If
        returnStr += nextString

        Return returnStr
    End Function
End Class
