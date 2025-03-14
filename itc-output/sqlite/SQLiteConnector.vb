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
SELECT 1;
PRAGMA foreign_keys=OFF;
BEGIN TRANSACTION;
CREATE TABLE [db_info] (
  [key] text NOT NULL
, [value] text NOT NULL
);
CREATE TABLE [person] (
  [id] bigint NOT NULL
, [group] text NOT NULL
, [login]  NOT NULL
, [code]  NULL
, CONSTRAINT [sqlite_autoindex_person_1] PRIMARY KEY ([id])
);
CREATE TABLE [booklet] (
  [id] bigint NOT NULL
, [name] text NOT NULL
, [personId] bigint NOT NULL
, CONSTRAINT [sqlite_autoindex_booklet_1] PRIMARY KEY ([id])
, CONSTRAINT [FK_booklet_0_0] FOREIGN KEY ([personId]) REFERENCES [person] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [unit] (
  [id] bigint NOT NULL
, [bookletId] bigint NOT NULL
, [name] text NOT NULL
, [alias] text NULL
, CONSTRAINT [sqlite_autoindex_unit_1] PRIMARY KEY ([id])
, CONSTRAINT [FK_unit_0_0] FOREIGN KEY ([bookletId]) REFERENCES [booklet] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [subform] (
  [id] bigint NOT NULL
, [unitId] bigint NOT NULL
, [key] text NULL
, CONSTRAINT [sqlite_autoindex_subform_1] PRIMARY KEY ([id])
, CONSTRAINT [FK_subform_0_0] FOREIGN KEY ([unitId]) REFERENCES [unit] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [response] (
  [subformId] bigint NOT NULL
, [variableId] text NOT NULL
, [value] text NULL
, [status] text NOT NULL
, [code] bigint DEFAULT (0) NOT NULL
, [score] bigint DEFAULT (0) NOT NULL
, CONSTRAINT [FK_response_0_0] FOREIGN KEY ([subformId]) REFERENCES [subform] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [chunk] (
  [unitId] bigint NOT NULL
, [key] text NOT NULL
, [type] text NULL
, [variables] text NULL
, [ts] bigint NULL
, CONSTRAINT [FK_chunk_0_0] FOREIGN KEY ([unitId]) REFERENCES [unit] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [unitLastState] (
  [unitId] bigint NOT NULL
, [key] text NOT NULL
, [value] text NULL
, CONSTRAINT [FK_unitLastState_0_0] FOREIGN KEY ([unitId]) REFERENCES [unit] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [unitLog] (
  [unitId] bigint NOT NULL
, [key] text NOT NULL
, [paremeter] text NULL
, [ts] bigint NULL
, CONSTRAINT [FK_unitLog_0_0] FOREIGN KEY ([unitId]) REFERENCES [unit] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [bookletLog] (
  [bookletId] bigint NOT NULL
, [key] text NOT NULL
, [paremeter] text NULL
, [ts] bigint NULL
, CONSTRAINT [FK_bookletLog_0_0] FOREIGN KEY ([bookletId]) REFERENCES [booklet] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
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
CREATE TRIGGER [fki_booklet_personId_person_id] BEFORE Insert ON [booklet] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table booklet violates foreign key constraint FK_booklet_0_0') WHERE (SELECT id FROM person WHERE  id = NEW.personId) IS NULL; END;
CREATE TRIGGER [fku_booklet_personId_person_id] BEFORE Update ON [booklet] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table booklet violates foreign key constraint FK_booklet_0_0') WHERE (SELECT id FROM person WHERE  id = NEW.personId) IS NULL; END;
CREATE TRIGGER [fki_unit_bookletId_booklet_id] BEFORE Insert ON [unit] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table unit violates foreign key constraint FK_unit_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
CREATE TRIGGER [fku_unit_bookletId_booklet_id] BEFORE Update ON [unit] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table unit violates foreign key constraint FK_unit_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
CREATE TRIGGER [fki_subform_unitId_unit_id] BEFORE Insert ON [subform] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table subform violates foreign key constraint FK_subform_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fku_subform_unitId_unit_id] BEFORE Update ON [subform] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table subform violates foreign key constraint FK_subform_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fki_response_subformId_subform_id] BEFORE Insert ON [response] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table response violates foreign key constraint FK_response_0_0') WHERE (SELECT id FROM subform WHERE  id = NEW.subformId) IS NULL; END;
CREATE TRIGGER [fku_response_subformId_subform_id] BEFORE Update ON [response] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table response violates foreign key constraint FK_response_0_0') WHERE (SELECT id FROM subform WHERE  id = NEW.subformId) IS NULL; END;
CREATE TRIGGER [fki_chunk_unitId_unit_id] BEFORE Insert ON [chunk] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table chunk violates foreign key constraint FK_chunk_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fku_chunk_unitId_unit_id] BEFORE Update ON [chunk] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table chunk violates foreign key constraint FK_chunk_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fki_unitLastState_unitId_unit_id] BEFORE Insert ON [unitLastState] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table unitLastState violates foreign key constraint FK_unitLastState_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fku_unitLastState_unitId_unit_id] BEFORE Update ON [unitLastState] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table unitLastState violates foreign key constraint FK_unitLastState_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fki_unitLog_unitId_unit_id] BEFORE Insert ON [unitLog] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table unitLog violates foreign key constraint FK_unitLog_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fku_unitLog_unitId_unit_id] BEFORE Update ON [unitLog] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table unitLog violates foreign key constraint FK_unitLog_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fki_bookletLog_bookletId_booklet_id] BEFORE Insert ON [bookletLog] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table bookletLog violates foreign key constraint FK_bookletLog_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
CREATE TRIGGER [fku_bookletLog_bookletId_booklet_id] BEFORE Update ON [bookletLog] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table bookletLog violates foreign key constraint FK_bookletLog_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
CREATE TRIGGER [fki_session_bookletId_booklet_id] BEFORE Insert ON [session] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table session violates foreign key constraint FK_session_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
CREATE TRIGGER [fku_session_bookletId_booklet_id] BEFORE Update ON [session] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table session violates foreign key constraint FK_session_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
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

    Public Function GetOpenConnection() As DbConnection
        Dim fact As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.SQLite")
        Dim sqliteConnection As DbConnection = fact.CreateConnection()
        sqliteConnection.ConnectionString = "Data Source=" + fileName
        sqliteConnection.Open()
        Return sqliteConnection
    End Function
End Class
