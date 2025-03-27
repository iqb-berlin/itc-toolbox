Public Class LoadDataFromCsvDialog
    Public WriteToXls As Boolean
    Public AllVariables As List(Of String)
    Public sqliteConnection As SQLiteConnector

    Public Sub New(writeToXls As Boolean, Optional sqliteConnection As SQLiteConnector = Nothing)
        InitializeComponent()
        Me.WriteToXls = writeToXls
        Me.sqliteConnection = sqliteConnection
    End Sub
End Class
