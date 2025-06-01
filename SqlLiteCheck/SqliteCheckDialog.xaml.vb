Public Class SqliteCheckDialog
    Private DB As SQLiteConnector
#Region "Vorspann"
    Public Sub New(SqliteDB As SQLiteConnector)
        InitializeComponent()
        DB = SqliteDB
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
            Dim peopleCounter As New Dictionary(Of String, Integer)
            For Each p As KeyValuePair(Of Long, String) In peopleListAll
                If Not peopleCounter.ContainsKey(p.Value) Then
                    peopleCounter.Add(p.Value, 1)
                Else
                    peopleCounter.Item(p.Value) = peopleCounter.Item(p.Value) + 1
                End If
            Next
            Dim moreThanOnePerson As List(Of String) = (From p As KeyValuePair(Of String, Integer) In peopleCounter
                                                        Where p.Value > 1 Let s As String = p.Key + " (" + p.Value.ToString + ")"
                                                        Order By p.Key Select s).ToList
            If moreThanOnePerson.Count > 0 Then
                worker.ReportProgress(0.0#, String.Join(vbNewLine, moreThanOnePerson))
            Else
                worker.ReportProgress(0.0#, "keine Personen-Dopplungen gefunden.")
            End If
            If Not worker.CancellationPending Then
                worker.ReportProgress(0.0#, "h:Prüfe auf Unit-Dopplungen")
                Dim progressMax As Long = peopleListAll.Count
                Dim progressStep As Long = 1
                Dim totalOfDoubleUnits As Integer = 0
                Dim totalOfPeopleWithDoubleUnits As Integer = 0
                For Each p As KeyValuePair(Of Long, String) In peopleListAll
                    If worker.CancellationPending Then Exit For
                    Dim progressValue As Double = 100 * progressStep / progressMax
                    progressStep += 1
                    worker.ReportProgress(progressValue)
                    Dim unitCounter As Dictionary(Of String, Dictionary(Of String, Integer)) = DB.getUnitList(p.Key)
                    Dim moreThanOneUnit As List(Of String) = (From b As KeyValuePair(Of String, Dictionary(Of String, Integer)) In unitCounter
                                                              From u As KeyValuePair(Of String, Integer) In b.Value
                                                              Where u.Value > 1 Let s As String = vbTab + b.Key + " / " + u.Key + ": " + u.Value.ToString
                                                              Order By b.Key, u.Key Select s).ToList
                    If moreThanOneUnit.Count > 0 Then
                        totalOfDoubleUnits += moreThanOneUnit.Count
                        totalOfPeopleWithDoubleUnits += 1
                        worker.ReportProgress(progressValue, p.Value + vbNewLine + String.Join(vbNewLine, moreThanOneUnit))
                    End If
                Next
                If totalOfPeopleWithDoubleUnits > 0 Then
                    worker.ReportProgress(0.0#, "Anzahl Personen mit Mehrfach-Units: " + totalOfPeopleWithDoubleUnits.ToString)
                    worker.ReportProgress(0.0#, "Anzahl Mehrfach-Units insgesamt: " + totalOfDoubleUnits.ToString)
                Else
                    worker.ReportProgress(0.0#, "keine Unit-Dopplungen gefunden.")
                End If
            End If
        End If
    End Sub
End Class
