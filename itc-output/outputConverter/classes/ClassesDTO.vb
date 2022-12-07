Public Class GroupDataDTO
    Public groupName As String
    Public bookletsStarted As Integer
    Public numUnitsMin As Integer
    Public numUnitsMax As Integer
    Public numUnitsTotal As Integer
    Public numUnitsAvg As Double
    Public lastChange As Long
End Class

Public Class LogEntryDTO
    Public groupname As String
    Public loginname As String
    Public code As String
    Public bookletname As String
    Public unitname As String
    Public timestamp As Long
    Public logentry As String
End Class

Public Class ResponseDataDTO
    Public id As String
    Public content As String
    Public ts As Long
    Public responseType As String
End Class

Public Class ResponseDTO
    Public groupname As String
    Public loginname As String
    Public code As String
    Public bookletname As String
    Public unitname As String
    Public laststate As String
    Public responses As List(Of ResponseDataDTO)
End Class

Public class BookletInfoDTO
    Public description As String
    Public label As String
    Public totalSize As long
End Class

Public Class BookletDTO
    Public name As String
    Public size As Integer
    Public modificationTime As Long
    Public type As String
    Public id As String
    'Public report
    Public info As BookletInfoDTO
End Class