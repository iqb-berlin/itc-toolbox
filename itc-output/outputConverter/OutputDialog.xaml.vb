Public Class OutputDialog
    Friend Const LogFileFirstLine = "groupname;loginname;code;bookletname;unitname;timestamp;logentry"
    Friend Const ResponsesFileFirstLine = "groupname;loginname;code;bookletname;unitname;responses;restorePoint;responseType;response-ts;restorePoint-ts;laststate"

    Friend bookletSize As New Dictionary(Of String, Long)

End Class
