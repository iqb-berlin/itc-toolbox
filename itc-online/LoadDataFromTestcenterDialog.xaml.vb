
 Class LoadDataFromTestcenterDialog
    Public creds As Net.NetworkCredential
    Public selectedDataGroups As List(Of String)
    Public ResponsesOnly As Boolean
    Public WriteToXls As Boolean
    Public AllVariables As List(Of String)
    Public replaceBigdata As Boolean

    Public Sub New(Optional responsesOnly As Boolean = False, Optional writeToXls As Boolean = True)
        InitializeComponent()
        Me.ResponsesOnly = responsesOnly
        Me.WriteToXls = writeToXls
    End Sub

End Class
