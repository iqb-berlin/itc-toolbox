Imports Newtonsoft.Json

Public Class SessionReport
    Inherits Session
    Public booklet As String
    Public sessionNumber As Integer
    Public sessionStartTs As Long
    Public contentLoadSpeed As Double
    Public firstUnitTS As Long
    Public lastUnitTS As Long
    Public unitsWithResponse As List(Of String)
End Class
