﻿Public Class globalOutputStore
    Public Shared itcConnection As ITCConnection = Nothing
    Public Shared personData As New PersonList
    Public Shared bookletSizes As New Dictionary(Of String, Long)
    Public Shared bigData As New Dictionary(Of String, String)
    Public Shared personResponses As New List(Of PersonResponses)

    Public Shared Sub clear()
        personData = New PersonList
        bookletSizes = New Dictionary(Of String, Long)
        bigData = New Dictionary(Of String, String)
        personResponses = New List(Of PersonResponses)
    End Sub
End Class
