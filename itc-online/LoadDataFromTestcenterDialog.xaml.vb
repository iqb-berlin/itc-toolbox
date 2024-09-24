
 Class LoadDataFromTestcenterDialog
    Public itcConnection As ITCConnection
    Public creds As Net.NetworkCredential
    Public selectedDataGroups As List(Of String)
    Public ResponsesOnly As Boolean
    Public WriteToXls As Boolean
    Public bookletSizes As Dictionary(Of String, Long)
    Public myTestPersonList As TestPersonList
    Public AllPeople As Dictionary(Of String, Dictionary(Of String, List(Of UnitLineData))) 'id -> booklet -> entries
    Public AllVariables As List(Of String)

    Public Sub New(itcConnection As ITCConnection, Optional responsesOnly As Boolean = False, Optional writeToXls As Boolean = True)
        InitializeComponent()
        Me.itcConnection = itcConnection
        Me.ResponsesOnly = responsesOnly
        Me.WriteToXls = writeToXls
    End Sub

End Class
