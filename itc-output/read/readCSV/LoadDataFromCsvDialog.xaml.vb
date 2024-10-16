Public Class OutputDialog
    Friend outputConfig As New OutputConfig With {.bookletSizes = Nothing, .omitUnits = Nothing, .variables = Nothing}

    Public WriteToXls As Boolean
    Public bookletSizes As Dictionary(Of String, Long)
    Public AllVariables As List(Of String)
    Public segregateBigdata As Boolean

    Public Sub New(Optional writeToXls As Boolean = True)
        InitializeComponent()
        Me.WriteToXls = writeToXls
    End Sub
End Class
