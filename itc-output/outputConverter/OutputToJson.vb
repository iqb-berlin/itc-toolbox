Imports Newtonsoft.Json
Public Class OutputToJson
    Public Shared Sub Write(
                           data As Dictionary(Of String, List(Of UnitLineData)),
                           targetJsonFilename As String
                           )
        Using file As New IO.StreamWriter(targetJsonFilename)
            Dim js As New JsonSerializer()
            js.Formatting = Formatting.Indented
            js.Serialize(file,
                         From group As KeyValuePair(Of String, List(Of UnitLineData)) In data
                         From u As UnitLineData In group.Value
                         Select u)
        End Using
    End Sub

    Public Shared Function Read(sourceJsonFilenames As String()) As List(Of UnitLineData)
        Dim returnData As List(Of UnitLineData) = Nothing
        Try
            For Each fn In sourceJsonFilenames
                Using file As New IO.StreamReader(fn)
                    Dim js As New JsonSerializer()
                    returnData = js.Deserialize(file, GetType(List(Of UnitLineData)))
                End Using
            Next
        Catch ex As Exception
            returnData = Nothing
        End Try
        Return returnData
    End Function
End Class
