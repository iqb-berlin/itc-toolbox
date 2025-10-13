Class LoadDataFromTestcenterToCsvDialog
    Public selectedDataGroups As List(Of String)
    Public itcConnection As ITCConnection

    Public Sub New(testcenterConnection As ITCConnection)
        InitializeComponent()
        Me.itcConnection = testcenterConnection
    End Sub
End Class
