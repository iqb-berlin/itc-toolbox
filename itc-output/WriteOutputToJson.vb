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
            js.Serialize(file, AllVariables)
            js.Serialize(file, AllPeople)
            js.Serialize(file, myTestPersonList)
            js.Serialize(file, bookletSizes)
        End Using
    End Sub

End Class
