Imports Newtonsoft.Json

Public Class readJsonBookletFilesToDbDialog
    Private fileName As String = Nothing
    Public SqliteDB As SQLiteConnector = Nothing

#Region "Vorspann"
    Public Sub New(sourceFileName As String)
        InitializeComponent()
        Me.fileName = sourceFileName
    End Sub

    Private Sub Me_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If fileName IsNot Nothing AndAlso SqliteDB IsNot Nothing Then
            Process1_bw = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
            Process1_bw.RunWorkerAsync()
        Else
            MBUC.AddMessage("Keine Datenbank oder Datei zugewiesen")
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
        If Not String.IsNullOrEmpty(e.UserState) Then MBUC.AddMessage(e.UserState)
    End Sub

    Private Sub Process1_bw_RunWorkerCompleted(ByVal sender As Object, ByVal e As ComponentModel.RunWorkerCompletedEventArgs) Handles Process1_bw.RunWorkerCompleted
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
        myworker.ReportProgress(0.0#, "Lese " + IO.Path.GetFileName(fileName))
        Using file As New IO.StreamReader(fileName)
            Dim js As New JsonSerializer()
            Dim bookletSizes As Dictionary(Of String, Long) = Nothing
            Try
                bookletSizes = js.Deserialize(file, GetType(Dictionary(Of String, Long)))
            Catch ex As Exception
                myworker.ReportProgress(0.0#, "Datenfehler: " + ex.Message)
                bookletSizes = Nothing
            End Try
            If bookletSizes IsNot Nothing Then SqliteDB.addBooklets(bookletSizes)
        End Using
    End Sub
End Class
