Imports System.Globalization
Public Class ToCsvDialog
    Private DB As SQLiteConnector
    Private deCulture As CultureInfo
#Region "Vorspann"
    Public Sub New(SqliteDB As SQLiteConnector)
        InitializeComponent()
        DB = SqliteDB
        deCulture = CultureInfo.CreateSpecificCulture("de-DE")
    End Sub

    Private Sub Me_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If DB Is Nothing Then
            BtnCancelClose.Content = "Schließen"
            MBUC.AddMessage("Keine Datenbank geöffnet.")
        Else
            MBUC.AddMessage("Prüfe Datenbank - bitte warten")
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
            worker.ReportProgress(0.0#, "h:Prüfe auf Personen-Dopplungen")
            Dim peopleListAll As Dictionary(Of Long, String) = DB.getPeopleListAll()
            Dim progressMax As Long = peopleListAll.Count
            Dim progressStep As Long = 1
            Dim totalOfDoubleUnits As Integer = 0
            Dim totalOfPeopleWithDoubleUnits As Integer = 0

            Dim file As System.IO.StreamWriter
            file = My.Computer.FileSystem.OpenTextFileWriter(My.Settings.lastfile_OutputTargetCsv, True)
            file.WriteLine(firstCsvLine)
            For Each p As KeyValuePair(Of Long, String) In peopleListAll
                If worker.CancellationPending Then Exit For
                Dim progressValue As Double = 100 * progressStep / progressMax
                progressStep += 1
                Dim responses As List(Of PersonResponseLong) = DB.getPersonResponsesLong(p.Key)
                For Each r As PersonResponseLong In responses
                    file.WriteLine(PersonResponseLongAsCsvString(r))
                    worker.ReportProgress(progressValue)
                Next
            Next
            file.Close()
        End If
    End Sub

    Private Const firstCsvLine = "Group;Login;Code;Booklet;Unit-Name;Unit-Alias;Response-Id;Response-Status;Response-Value;Response-Subform;Response-Code;Response-Score;Response-Timestamp;Response-Timestamp-String"

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
        Dim tsDt As New DateTime(1970, 1, 1, 0, 0, 0, 0)
        tsDt = tsDt.AddMilliseconds(r.ts)
        returnStr += separator + """" + tsDt.ToString(deCulture) + """"

        Return returnStr
    End Function
End Class
