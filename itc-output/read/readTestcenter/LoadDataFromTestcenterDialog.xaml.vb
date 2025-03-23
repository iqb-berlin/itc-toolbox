Public Enum TestcenterReadMode
    Reviews
    Responses
    SystemCheck
End Enum

Public Enum DataTarget
    Standard
    JsonFiles
    Datastore
    Sqlite
End Enum
Class LoadDataFromTestcenterDialog
    Public selectedDataGroups As List(Of String)
    Public target As DataTarget
    Public segregateBigdata As Boolean
    Public readMode As TestcenterReadMode
    Public itcConnection As ITCConnection
    Public sqliteConnection As SQLiteConnector

    Public Sub New(testcenterConnection As ITCConnection,
                   mode As TestcenterReadMode,
                   target As DataTarget,
                   Optional sqliteConnection As SQLiteConnector = Nothing)
        InitializeComponent()
        Me.itcConnection = testcenterConnection
        Me.readMode = mode
        Me.target = target
        Me.sqliteConnection = sqliteConnection
        If Me.target = DataTarget.Sqlite AndAlso Me.sqliteConnection Is Nothing Then Me.target = DataTarget.Datastore
    End Sub
End Class
