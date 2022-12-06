
Public Class LoadDataFromTestcenterDialog
    Public itcConnection As ITCConnection
    Public creds As Net.NetworkCredential

    Public Sub New(itcConnection As ITCConnection)
        InitializeComponent()
        Me.itcConnection = itcConnection
    End Sub

End Class
