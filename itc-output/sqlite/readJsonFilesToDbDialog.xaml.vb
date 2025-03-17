Imports Newtonsoft.Json

Public Class readJsonFilesToDbDialog
    Private files As String() = Nothing
    Public SqliteDB As SQLiteConnector = Nothing

#Region "Vorspann"
    Public Sub New(fileNameList As String())
        InitializeComponent()
        Me.files = fileNameList
    End Sub

    Private Sub Me_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If files.Length > 0 Then
            Process1_bw = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
            Process1_bw.RunWorkerAsync()
        Else
            APBUC.Value = 0.0#
            MBUC.AddMessage("Keine Datenbank zugewiesen")
            BtnCancel.Content = "Schließen"
        End If
    End Sub

    Private WithEvents Process1_bw As ComponentModel.BackgroundWorker = Nothing

    Private Sub BtnCancel_Click() Handles BtnCancel.Click
        If Process1_bw IsNot Nothing AndAlso Process1_bw.IsBusy Then
            Process1_bw.CancelAsync()
            BtnCancel.IsEnabled = False
            MBUC.AddMessage("Abbruch - bitte warten")
        Else
            DialogResult = False
        End If
    End Sub

    Private Sub bw_ProgressChanged(ByVal sender As Object, ByVal e As ComponentModel.ProgressChangedEventArgs) Handles Process1_bw.ProgressChanged
        Me.APBUC.UpdateProgressState(e.ProgressPercentage)
        If Not String.IsNullOrEmpty(e.UserState) Then MBUC.AddMessage(e.UserState)
    End Sub

    Private Sub Process1_bw_RunWorkerCompleted(ByVal sender As Object, ByVal e As ComponentModel.RunWorkerCompletedEventArgs) Handles Process1_bw.RunWorkerCompleted
        APBUC.Value = 0.0#
        MBUC.AddMessage("beendet")
        BtnCancel.Content = "Schließen"
        BtnCancel.IsEnabled = True
        If e.Cancelled Then MBUC.AddMessage("durch Nutzer abgebrochen.")
    End Sub
#End Region

    '######################################################################################
    '######################################################################################
    Private Sub Process1_bw_DoWork(ByVal sender As Object, ByVal e As ComponentModel.DoWorkEventArgs) Handles Process1_bw.DoWork
        Dim myworker As ComponentModel.BackgroundWorker = sender
        Dim progressMax As Integer = files.Length
        Dim progressCount As Integer = 0
        Dim progressValue As Double
        For Each fn In files
            myworker.ReportProgress(progressValue, "Lese " + IO.Path.GetFileName(fn))
            Try
                Using file As New IO.StreamReader(fn)
                    Dim js As New JsonSerializer()
                    Dim groupData As List(Of Person) = js.Deserialize(file, GetType(List(Of Person)))
                    Dim plusValuePerPerson As Double = 100 / progressMax
                    Dim fileMax As Integer = groupData.Count
                    Dim fileCount As Integer = 0
                    For Each p As Person In groupData
                        fileCount += 1
                        If myworker.CancellationPending Then Exit For
                        progressValue = progressCount * plusValuePerPerson + fileCount * (plusValuePerPerson / fileMax)
                        myworker.ReportProgress(progressValue)
                        SqliteDB.addPerson(p)
                    Next
                End Using
            Catch ex As Exception
                myworker.ReportProgress(progressValue, "Fehler " + IO.Path.GetFileName(fn) + ": " + ex.Message)
            End Try
            progressCount += 1
            If myworker.CancellationPending Then Exit For
        Next
    End Sub
End Class
