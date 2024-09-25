Public Class OutputDialog
    Friend Const LogFileFirstLine = "groupname;loginname;code;bookletname;unitname;timestamp;logentry"
    Friend Const ResponsesFileFirstLine = "groupname;loginname;code;bookletname;unitname;responses;laststate"
    Friend Const ResponsesFileFirstLineLegacy = "groupname;loginname;code;bookletname;unitname;responses;restorePoint;responseType;response-ts;restorePoint-ts;laststate"

    Friend outputConfig As New OutputConfig With {.bookletSizes = Nothing, .omitUnits = Nothing, .variables = Nothing}

    Public ResponsesOnly As Boolean
    Public WriteToXls As Boolean
    Public bookletSizes As Dictionary(Of String, Long)
    Public myTestPersonList As TestPersonList
    Public AllPeople As Dictionary(Of String, Dictionary(Of String, List(Of UnitLineData))) 'id -> booklet -> entries
    Public AllVariables As List(Of String)

    Public Sub New(Optional writeToXls As Boolean = True)
        InitializeComponent()
        Me.ResponsesOnly = ResponsesOnly
        Me.WriteToXls = writeToXls
    End Sub
End Class
