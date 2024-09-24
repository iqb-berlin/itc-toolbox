Imports Newtonsoft.Json
Public Class WriteOutputToJson
    Public Shared Sub Write(
                           AllVariables As List(Of String),
                           AllPeople As Dictionary(Of String, Dictionary(Of String, List(Of UnitLineData))),
                           myTestPersonList As TestPersonList,
                           bookletSizes As Dictionary(Of String, Long),
                           targetJsonFilename As String
                           )
        Using file As New IO.StreamWriter(targetJsonFilename)
            Dim js As New JsonSerializer()
            js.Serialize(file,
                         From p As KeyValuePair(Of String, Dictionary(Of String, List(Of UnitLineData))) In AllPeople
                         From b As KeyValuePair(Of String, List(Of UnitLineData)) In p.Value
                         From u As UnitLineData In b.Value
                         Select u)
        End Using
    End Sub

End Class
