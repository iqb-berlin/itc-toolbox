Imports System.Data.Common
Imports System.Globalization
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
        Using sqliteConnection As DbConnection = GetOpenConnection(False)
            If addFullSchema Then
                Using cmd As DbCommand = sqliteConnection.CreateCommand()
                    cmd.CommandText = "
SELECT 1;
PRAGMA foreign_keys=OFF;
BEGIN TRANSACTION;
CREATE TABLE [db_info] (
  [key] text NOT NULL
, [value] text NOT NULL
);
CREATE TABLE [person] (
  [id] INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL
, [group] text NOT NULL
, [login]  NOT NULL
, [code]  NULL
);
CREATE TABLE [booklet] (
  [id] INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL
, [name] text NOT NULL
, [personId] bigint NOT NULL
, [lastTs] bigint DEFAULT (0) NOT NULL
, [firstTs] bigint DEFAULT (0) NOT NULL
, CONSTRAINT [FK_booklet_0_0] FOREIGN KEY ([personId]) REFERENCES [person] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [session] (
  [bookletId] bigint NOT NULL
, [browser] text NULL
, [os] text NULL
, [screen] text NULL
, [ts] bigint NULL
, [loadCompleteMS] bigint NULL
, CONSTRAINT [FK_session_0_0] FOREIGN KEY ([bookletId]) REFERENCES [booklet] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [bookletLog] (
  [bookletId] bigint NOT NULL
, [key] text NOT NULL
, [parameter] text NULL
, [ts] bigint NULL
, CONSTRAINT [FK_bookletLog_0_0] FOREIGN KEY ([bookletId]) REFERENCES [booklet] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [unit] (
  [id] INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL
, [bookletId] bigint NOT NULL
, [name] text NOT NULL
, [alias] text NULL
, CONSTRAINT [FK_unit_0_0] FOREIGN KEY ([bookletId]) REFERENCES [booklet] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [unitLog] (
  [unitId] bigint NOT NULL
, [key] text NOT NULL
, [parameter] text NULL
, [ts] bigint NULL
, CONSTRAINT [FK_unitLog_0_0] FOREIGN KEY ([unitId]) REFERENCES [unit] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [unitLastState] (
  [unitId] bigint NOT NULL
, [key] text NOT NULL
, [value] text NULL
, CONSTRAINT [FK_unitLastState_0_0] FOREIGN KEY ([unitId]) REFERENCES [unit] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [chunk] (
  [unitId] bigint NOT NULL
, [key] text NOT NULL
, [type] text NULL
, [variables] text NULL
, [ts] bigint NULL
, CONSTRAINT [FK_chunk_0_0] FOREIGN KEY ([unitId]) REFERENCES [unit] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [subform] (
  [id] INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL
, [unitId] bigint NOT NULL
, [key] text NULL
, CONSTRAINT [FK_subform_0_0] FOREIGN KEY ([unitId]) REFERENCES [unit] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [response] (
  [subformId] bigint NOT NULL
, [variableId] text NOT NULL
, [value] text NULL
, [status] text NOT NULL
, [code] integer DEFAULT (0) NOT NULL
, [score] integer DEFAULT (0) NOT NULL
, CONSTRAINT [FK_response_0_0] FOREIGN KEY ([subformId]) REFERENCES [subform] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TRIGGER [fki_booklet_personId_person_id] BEFORE Insert ON [booklet] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table booklet violates foreign key constraint FK_booklet_0_0') WHERE (SELECT id FROM person WHERE  id = NEW.personId) IS NULL; END;
CREATE TRIGGER [fku_booklet_personId_person_id] BEFORE Update ON [booklet] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table booklet violates foreign key constraint FK_booklet_0_0') WHERE (SELECT id FROM person WHERE  id = NEW.personId) IS NULL; END;
CREATE TRIGGER [fki_session_bookletId_booklet_id] BEFORE Insert ON [session] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table session violates foreign key constraint FK_session_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
CREATE TRIGGER [fku_session_bookletId_booklet_id] BEFORE Update ON [session] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table session violates foreign key constraint FK_session_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
CREATE TRIGGER [fki_bookletLog_bookletId_booklet_id] BEFORE Insert ON [bookletLog] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table bookletLog violates foreign key constraint FK_bookletLog_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
CREATE TRIGGER [fku_bookletLog_bookletId_booklet_id] BEFORE Update ON [bookletLog] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table bookletLog violates foreign key constraint FK_bookletLog_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
CREATE TRIGGER [fki_unit_bookletId_booklet_id] BEFORE Insert ON [unit] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table unit violates foreign key constraint FK_unit_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
CREATE TRIGGER [fku_unit_bookletId_booklet_id] BEFORE Update ON [unit] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table unit violates foreign key constraint FK_unit_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
CREATE TRIGGER [fki_unitLog_unitId_unit_id] BEFORE Insert ON [unitLog] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table unitLog violates foreign key constraint FK_unitLog_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fku_unitLog_unitId_unit_id] BEFORE Update ON [unitLog] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table unitLog violates foreign key constraint FK_unitLog_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fki_unitLastState_unitId_unit_id] BEFORE Insert ON [unitLastState] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table unitLastState violates foreign key constraint FK_unitLastState_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fku_unitLastState_unitId_unit_id] BEFORE Update ON [unitLastState] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table unitLastState violates foreign key constraint FK_unitLastState_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fki_chunk_unitId_unit_id] BEFORE Insert ON [chunk] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table chunk violates foreign key constraint FK_chunk_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fku_chunk_unitId_unit_id] BEFORE Update ON [chunk] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table chunk violates foreign key constraint FK_chunk_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fki_subform_unitId_unit_id] BEFORE Insert ON [subform] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table subform violates foreign key constraint FK_subform_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fku_subform_unitId_unit_id] BEFORE Update ON [subform] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table subform violates foreign key constraint FK_subform_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fki_response_subformId_subform_id] BEFORE Insert ON [response] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table response violates foreign key constraint FK_response_0_0') WHERE (SELECT id FROM subform WHERE  id = NEW.subformId) IS NULL; END;
CREATE TRIGGER [fku_response_subformId_subform_id] BEFORE Update ON [response] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table response violates foreign key constraint FK_response_0_0') WHERE (SELECT id FROM subform WHERE  id = NEW.subformId) IS NULL; END;
COMMIT;"
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

    Public Function GetOpenConnection(ReadOnlyMode As Boolean) As DbConnection
        Dim fact As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.SQLite")
        Dim sqliteConnection As DbConnection = fact.CreateConnection()
        sqliteConnection.ConnectionString = "Data Source=" + fileName + IIf(ReadOnlyMode, ";Read Only=True;", "")
        sqliteConnection.Open()
        Return sqliteConnection
    End Function

    Public Function GetCoreData(closeConnection As Boolean) As String
        Dim returnText As String = ""
        Using sqliteConnection As DbConnection = GetOpenConnection(True)
            Using cmd As DbCommand = sqliteConnection.CreateCommand()
                cmd.CommandText = "SELECT COUNT(*) FROM [person];"
                Dim personCount As Long = cmd.ExecuteScalar()
                cmd.CommandText = "SELECT COUNT(*) FROM [response];"
                Dim responseCount As Long = cmd.ExecuteScalar()
                Dim deCulture = CultureInfo.CreateSpecificCulture("de-DE")
                returnText = "Anzahl Personen: " + personCount.ToString("N0", deCulture) +
                    ", Anzahl Antwortdaten: " + responseCount.ToString("N0", deCulture)
            End Using
        End Using
        If closeConnection Then
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End If
        Return returnText
    End Function

    Public Sub addPerson(p As Person)
        Using sqliteConnection As DbConnection = GetOpenConnection(False)
            Using cmd As DbCommand = sqliteConnection.CreateCommand()
                'Dim personDbId As Long = sqliteConnection.

                cmd.CommandText = "
BEGIN TRANSACTION;
INSERT INTO [person] ([group],[login],[code]) VALUES ('" + p.group + "', '" + p.login + "', '" + p.code + "');
SELECT last_insert_rowid();
COMMIT;"
                Dim lastInsert_PersonId As Long = cmd.ExecuteScalar()
                For Each b As Booklet In p.booklets
                    'If b.lastTS = 0 OrElse b.firstTS = 0 Then b.setTimestamps()
                    'cmd.CommandText = "select * from booklet where personId = @personId AND name = @bookletName"
                    'cmd.Parameters.Add("@personId", lastInsert_PersonId)

                    '" + lastInsert_PersonId.ToString + "' "


                    cmd.CommandText = "
BEGIN TRANSACTION;
INSERT INTO [booklet] ([personId],[name]) VALUES (" + lastInsert_PersonId.ToString + ", '" + b.id + "');
SELECT last_insert_rowid();
COMMIT;"
                    Dim lastInsert_BookletId As Long = cmd.ExecuteScalar()
                    cmd.CommandText = ""
                    For Each bLog As LogEntry In b.logs
                        cmd.CommandText += "
INSERT INTO [bookletLog] ([bookletId],[key],[parameter],[ts]) VALUES (" +
lastInsert_BookletId.ToString + ", '" + bLog.key + "','" + bLog.parameter + "'," + bLog.ts.ToString + ");"
                    Next
                    For Each s As Session In b.sessions
                        cmd.CommandText += "
INSERT INTO [session] ([bookletId],[browser],[os],[ts],[screen],[loadCompleteMS]) VALUES (" +
lastInsert_BookletId.ToString + ", '" + s.browser + "','" + s.os + "'," + s.ts.ToString + ",'" + s.screen + "'," + s.loadCompleteMS.ToString + ");"
                    Next
                    If Not String.IsNullOrEmpty(cmd.CommandText) Then
                        cmd.CommandText = "BEGIN TRANSACTION;" + cmd.CommandText + "COMMIT;"
                        cmd.ExecuteNonQuery()
                    End If

                    For Each u As Unit In b.units
                        cmd.CommandText = "
BEGIN TRANSACTION;
INSERT INTO [unit] ([bookletId],[name],[alias]) VALUES (" + lastInsert_BookletId.ToString + ", '" + u.id + "', '" + u.alias + "');
SELECT last_insert_rowid();
COMMIT;"
                        Dim lastInsert_UnitId As Long = cmd.ExecuteScalar()

                        cmd.CommandText = ""
                        For Each ls As LastStateEntry In u.laststate
                            cmd.CommandText += "
INSERT INTO [unitLastState] ([unitId],[key],[value]) VALUES (" +
lastInsert_UnitId.ToString + ", '" + ls.key + "','" + ls.value + "');"
                        Next
                        For Each l As LogEntry In u.logs
                            cmd.CommandText += "
INSERT INTO [unitLog] ([unitId],[key],[parameter],[ts]) VALUES (" +
lastInsert_UnitId.ToString + ", '" + l.key + "','" + l.parameter + "'," + l.ts.ToString + ");"
                        Next
                        For Each r As ResponseChunkData In u.chunks
                            cmd.CommandText += "
INSERT INTO [chunk] ([unitId],[key],[type],[variables],[ts]) VALUES (" +
lastInsert_UnitId.ToString + ", '" + r.id + "','" + r.type + "','" + String.Join(" ", r.variables) + "'," + r.ts.ToString + ");"
                        Next
                        If Not String.IsNullOrEmpty(cmd.CommandText) Then
                            cmd.CommandText = "BEGIN TRANSACTION;" + cmd.CommandText + "COMMIT;"
                            cmd.ExecuteNonQuery()
                        End If

                        For Each sf As SubForm In u.subforms
                            cmd.CommandText = "
BEGIN TRANSACTION;
INSERT INTO [subform] ([unitId],[key]) VALUES (" + lastInsert_UnitId.ToString + ", '" + sf.id + "');
SELECT last_insert_rowid();
COMMIT;"
                            Dim lastInsert_SubformId As Long = cmd.ExecuteScalar()
                            cmd.CommandText = ""
                            For Each resp As ResponseData In sf.responses
                                cmd.CommandText += "
INSERT INTO [response] ([subformId],[variableId],[value],[status],[code],[score]) VALUES (" +
lastInsert_SubformId.ToString + ", '" + resp.id + "','" + resp.value.Replace("'", "''") + "','" + resp.status + "'," +
                                resp.code.ToString + "," + resp.score.ToString + ");"
                            Next
                            If Not String.IsNullOrEmpty(cmd.CommandText) Then
                                cmd.CommandText = "BEGIN TRANSACTION;" + cmd.CommandText + "COMMIT;"
                                cmd.ExecuteNonQuery()
                            End If
                        Next
                    Next
                Next
            End Using
        End Using
    End Sub
End Class
