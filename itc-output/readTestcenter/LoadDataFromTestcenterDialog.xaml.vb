Public Enum TestcenterReadMode
    Reviews
    Responses
    SystemCheck
End Enum
Class LoadDataFromTestcenterDialog
    Public selectedDataGroups As List(Of String)
    Public write As Boolean
    Public segregateBigdata As Boolean
    Public readMode As TestcenterReadMode

    Public Sub New(mode As TestcenterReadMode, Optional instantWrite As Boolean = True)
        InitializeComponent()
        Me.readMode = mode
        Me.write = instantWrite
    End Sub

End Class
