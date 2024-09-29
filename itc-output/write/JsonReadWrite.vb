Imports Newtonsoft.Json
Public Class JsonReadWrite
    Public Shared Sub Write(targetJsonFilename As String)
        Using file As New IO.StreamWriter(targetJsonFilename)
            Dim js As New JsonSerializer()
            js.Formatting = Formatting.Indented
            js.Serialize(file,
                         From p As KeyValuePair(Of String, Person) In globalOutputStore.personData
                         Select p.Value)
        End Using
    End Sub

    Public Shared Sub WriteByGroup(targetFoldername As String)
        Dim groups As List(Of String) = (From p As KeyValuePair(Of String, Person) In globalOutputStore.personData Select p.Value.group).Distinct.ToList
        For Each g As String In groups
            Using file As New IO.StreamWriter(targetFoldername + IO.Path.DirectorySeparatorChar + g + ".json")
                Dim js As New JsonSerializer()
                js.Formatting = Formatting.Indented
                js.Serialize(file,
                         From p As KeyValuePair(Of String, Person) In globalOutputStore.personData Where p.Value.group = g
                         Select p.Value)
            End Using
        Next
    End Sub

    Public Shared Sub WriteBigData(targetFoldername As String)
        For Each big As KeyValuePair(Of String, String) In globalOutputStore.bigData
            Using file As New IO.StreamWriter(targetFoldername + IO.Path.DirectorySeparatorChar + big.Key)
                file.Write(big.Value)
            End Using
        Next
    End Sub

    Public Shared Sub WriteBooklets(targetJsonFilename As String)
        Using file As New IO.StreamWriter(targetJsonFilename)
            Dim js As New JsonSerializer()
            js.Formatting = Formatting.Indented
            js.Serialize(file, globalOutputStore.bookletSizes)
        End Using
    End Sub

    Public Shared Sub Read(sourceJsonFilenames As String())
        Try
            For Each fn In sourceJsonFilenames
                Using file As New IO.StreamReader(fn)
                    Dim js As New JsonSerializer()
                    Dim groupData As List(Of Person) = js.Deserialize(file, GetType(List(Of Person)))
                    For Each p As Person In groupData
                        globalOutputStore.personData.Add(p.group + p.login + p.code, p)
                    Next
                End Using
            Next
        Catch ex As Exception
            Debug.Print(sourceJsonFilenames.ToString, ex.Message)
        End Try
    End Sub
End Class
