Class LoadDataFromTestcenterToCsvDialog
    Public selectedDataGroups As List(Of String)
    Public itcConnection As ITCConnection
    Public csvType As CsvType

    Public Sub New(testcenterConnection As ITCConnection, type As CsvType)
        InitializeComponent()
        Me.itcConnection = testcenterConnection
        csvType = type
    End Sub
End Class
