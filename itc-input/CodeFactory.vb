Public Class CodeFactory
    Private Const codeCharacters = "abcdefghprqstuvxyz"
    Private Const codeNumbers = "2345679"
    Public Shared Function GetNewCode(codeLen As Integer) As String
        Dim newCode As String = ""
        Dim isNumber As Boolean = False
        Randomize()
        Do
            newCode = newCode & IIf(isNumber, Mid(codeNumbers, Int(codeNumbers.Length * Rnd() + 1), 1), Mid(codeCharacters, Int(codeCharacters.Length * Rnd() + 1), 1))
            isNumber = Not isNumber
        Loop Until newCode.Length = codeLen
        Return newCode
    End Function

    Public Shared Function GetNewCodeList(codeLen As Integer, codeCount As Integer) As List(Of String)
        Dim codeList As New List(Of String)
        Randomize()
        For i As Integer = 1 To codeCount
            Dim newCode As String
            Do
                newCode = ""
                Dim isNumber As Boolean = False
                Do
                    newCode = newCode & IIf(isNumber, Mid(codeNumbers, Int(codeNumbers.Length * Rnd() + 1), 1), Mid(codeCharacters, Int(codeCharacters.Length * Rnd() + 1), 1))
                    isNumber = Not isNumber
                Loop Until newCode.Length = codeLen
            Loop While codeList.Contains(newCode)
            codeList.Add(newCode)
        Next
        Return codeList
    End Function
End Class
