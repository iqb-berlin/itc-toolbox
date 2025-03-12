Imports System.Data.Common
Imports iqb.lib.windows

Public Class SQLiteConnector
    Const currentDbVersion = 1
    Private fileName As String
    Public ReadOnly dbVersion As Integer
    Public ReadOnly dbCreator As String
    Public ReadOnly dbCreatedDateTime As String
    Public Sub New(dbFileName As String)
        fileName = dbFileName
        Dim addFullSchema As Boolean = Not IO.File.Exists(fileName)
        Using sqliteConnection As DbConnection = GetOpenConnection()
            If addFullSchema Then
                Using cmd As DbCommand = sqliteConnection.CreateCommand()
                    cmd.CommandText = "
                        CREATE TABLE [db_info] ([key] TEXT NOT NULL, [value] TEXT NOT NULL);
                        CREATE TABLE [person] ([group] text NOT NULL, [login] text NOT NULL, [code] text NULL);
                    "
                    cmd.ExecuteNonQuery()
                End Using
                dbVersion = 1
                Dim now As DateTime = DateTime.Now
                dbCreatedDateTime = now.ToShortDateString + " " + now.ToShortTimeString
                dbCreator = ADFactory.GetMyNameLong
                Using cmd As DbCommand = sqliteConnection.CreateCommand()
                    cmd.CommandText = "
                        INSERT INTO [db_info] ([key],[value]) VALUES ('name', 'IQB-Testcenter-Output');
                        INSERT INTO [db_info] ([key],[value]) VALUES ('dbVersion', '" + dbVersion.ToString + "');
                        INSERT INTO [db_info] ([key],[value]) VALUES ('dbCreator', '" + dbCreator + "');
                        INSERT INTO [db_info] ([key],[value]) VALUES ('dbCreatedDateTime', '" + dbCreatedDateTime + "');
                    "
                    cmd.ExecuteNonQuery()
                End Using
            Else
                Using cmd As DbCommand = sqliteConnection.CreateCommand()
                    cmd.CommandText = "SELECT * FROM [db_info];"
                    Dim dbReader As DbDataReader = cmd.ExecuteReader()
                    While dbReader.Read()
                        Dim key As String = dbReader.GetString(0)
                        Dim value As String = dbReader.GetString(1)
                        Select Case key
                            Case "dbVersion" : dbVersion = Long.Parse(value)
                            Case "dbCreator" : dbCreator = value
                            Case "dbCreatedDateTime" : dbCreatedDateTime = value
                        End Select
                    End While
                End Using
            End If
        End Using
    End Sub

    Public Function GetOpenConnection() As DbConnection
        Dim fact As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.SQLite")
        Dim sqliteConnection As DbConnection = fact.CreateConnection()
        sqliteConnection.ConnectionString = "Data Source=" + fileName
        sqliteConnection.Open()
        Return sqliteConnection
    End Function
End Class
