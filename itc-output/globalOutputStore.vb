Public Class globalOutputStore
    Public Shared personDataFull As New PersonList
    Public Shared bookletSizes As New Dictionary(Of String, Long)
    Public Shared personResponses As New List(Of PersonResponses)

    Public Shared Sub clear()
        personDataFull = New PersonList
        bookletSizes = New Dictionary(Of String, Long)
        personResponses = New List(Of PersonResponses)
    End Sub
End Class
