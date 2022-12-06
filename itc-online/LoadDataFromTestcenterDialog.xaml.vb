
 Class LoadDataFromTestcenterDialog
    Public itcConnection As ITCConnection
    Public creds As Net.NetworkCredential
    Public selectedDataGroups As List(Of String)

    Public Sub New(itcConnection As ITCConnection)
        InitializeComponent()
        Me.itcConnection = itcConnection
    End Sub

End Class
