Public Class OutputDialog
    Friend Const LogFileFirstLine = "groupname;loginname;code;bookletname;unitname;timestamp;logentry"
    Friend Const ResponsesFileFirstLine = "groupname;loginname;code;bookletname;unitname;responses;responseType;response-ts;laststate"
    Friend Const ResponsesFileFirstLineLegacy = "groupname;loginname;code;bookletname;unitname;responses;restorePoint;responseType;response-ts;restorePoint-ts;laststate"

    Friend outputConfig As New OutputConfig With {.bookletSizes = Nothing, .omitUnits = Nothing, .variables = Nothing}

End Class
