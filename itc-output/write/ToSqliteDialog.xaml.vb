
Imports System.Data.Common

Public Class ToSqliteDialog

#Region "Vorspann"
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Me_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        BtnClose.Visibility = Windows.Visibility.Collapsed

        If String.IsNullOrEmpty(My.Settings.lastfile_OutputTargetSqlite) Then
            BtnContinue.Visibility = Windows.Visibility.Collapsed
            MBUC.AddMessage("Keine Zieldatei gewählt.")
        Else
            MBUC.AddMessage("Zieldatei: " + IO.Path.GetFileName(My.Settings.lastfile_OutputTargetSqlite))
            MBUC.AddMessage("Bitte Optionen wählen")
        End If
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
        DialogResult = True
    End Sub

    Private Sub bw_ProgressChanged(ByVal sender As Object, ByVal e As ComponentModel.ProgressChangedEventArgs) Handles Process1_bw.ProgressChanged, Process2_bw.ProgressChanged
        Me.APBUC.UpdateProgressState(e.ProgressPercentage)
        If Not String.IsNullOrEmpty(e.UserState) Then MBUC.AddMessage(e.UserState)
    End Sub

    Private Sub Process1_bw_RunWorkerCompleted(ByVal sender As Object, ByVal e As ComponentModel.RunWorkerCompletedEventArgs) Handles Process1_bw.RunWorkerCompleted
        APBUC.Value = 0.0#
        MBUC.AddMessage("beendet")
        BtnCancel.Visibility = Windows.Visibility.Collapsed
        If e.Cancelled Then MBUC.AddMessage("durch Nutzer abgebrochen.")

        BtnClose.Visibility = Windows.Visibility.Visible
    End Sub

    Private Sub BtnContinue_Click() Handles BtnContinue.Click
        If ChBResonses.IsChecked Then
            DPParameters.IsEnabled = False
            BtnClose.Visibility = Windows.Visibility.Collapsed
            BtnContinue.Visibility = Windows.Visibility.Collapsed
            BtnCancel.Visibility = Windows.Visibility.Visible

            Process1_bw = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
            Process1_bw.RunWorkerAsync()
        Else
            MBUC.AddMessage("Bitte Optionen wählen")
        End If
    End Sub

#End Region

    '######################################################################################
    '######################################################################################
    Private Sub Process1_bw_DoWork(ByVal sender As Object, ByVal e As ComponentModel.DoWorkEventArgs) Handles Process1_bw.DoWork
        Dim worker As ComponentModel.BackgroundWorker = sender
        Dim targetSqliteFilename As String = My.Settings.lastfile_OutputTargetSqlite
        Dim myTemplate As Byte() = Nothing

        Dim AllVariables As New List(Of String)
        worker.ReportProgress(0.0#, "Ermittle Variablen")

        For Each p As KeyValuePair(Of String, Person) In globalOutputStore.personDataFull
            For Each b As Booklet In p.Value.booklets
                For Each u As Unit In b.units
                    For Each rSub As SubForm In u.subforms
                        Dim varPrefix As String = u.alias
                        If Not String.IsNullOrEmpty(rSub.id) Then varPrefix += "##" + rSub.id
                        For Each r As ResponseData In rSub.responses
                            If r.status = "VALUE_CHANGED" AndAlso Not AllVariables.Contains(varPrefix + "##" + r.id) Then AllVariables.Add(varPrefix + "##" + r.id)
                        Next
                    Next
                Next
            Next
        Next
        worker.ReportProgress(0.0#, AllVariables.Count.ToString + " Variablen gefunden.")

        If AllVariables.Count > 0 Then
            Dim fact As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.SQLite")
            Using sqliteConnection As DbConnection = fact.CreateConnection()
                sqliteConnection.ConnectionString = "Data Source=" + targetSqliteFilename
                sqliteConnection.Open()
                For Each person As Person In
                            From kvp As KeyValuePair(Of String, Person) In globalOutputStore.personDataFull Order By kvp.Key Select kvp.Value
                    If worker.CancellationPending Then
                        e.Cancel = True
                        Exit For
                    End If
                    Using cmd As DbCommand = sqliteConnection.CreateCommand()
                        cmd.CommandText = "INSERT INTO person ([group], [login], [code]) VALUES ('" + person.group + "', '" + person.login + "', '" + person.code + "');"
                        cmd.ExecuteNonQuery()
                    End Using
                Next
            End Using
        End If

    End Sub
End Class
