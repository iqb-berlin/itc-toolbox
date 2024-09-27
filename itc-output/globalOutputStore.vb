Public Class globalOutputStore
    Public Shared itcConnection As ITCConnection = Nothing
    Public Shared personData As New PersonList
    Public Shared bookletSizes As New Dictionary(Of String, Long)

    Public Shared Sub clear()
        itcConnection = Nothing
        personData = New PersonList
        bookletSizes = New Dictionary(Of String, Long)
    End Sub
End Class
