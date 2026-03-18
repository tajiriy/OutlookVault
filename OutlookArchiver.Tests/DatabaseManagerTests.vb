Option Explicit On
Option Strict On
Option Infer Off

Imports NUnit.Framework
Imports OutlookArchiver.Data
Imports System.Data.SQLite
Imports System.IO

Namespace Tests

    <TestFixture>
    Public Class DatabaseManagerTests

        Private _testDbPath As String
        Private _dbManager As DatabaseManager

        <SetUp>
        Public Sub SetUp()
            _testDbPath = Path.Combine(Path.GetTempPath(), "OutlookArchiverTest_" & Guid.NewGuid().ToString("N") & ".db")
            _dbManager = New DatabaseManager(_testDbPath)
        End Sub

        <TearDown>
        Public Sub TearDown()
            ' WAL ファイルも含めて削除
            If File.Exists(_testDbPath) Then File.Delete(_testDbPath)
            If File.Exists(_testDbPath & "-wal") Then File.Delete(_testDbPath & "-wal")
            If File.Exists(_testDbPath & "-shm") Then File.Delete(_testDbPath & "-shm")
        End Sub

        ' ── スキーマ初期化 ────────────────────────────────────────────

        <Test>
        Public Sub Initialize_CreatesDbFile()
            _dbManager.Initialize()

            Assert.IsTrue(File.Exists(_testDbPath))
        End Sub

        <Test>
        Public Sub Initialize_CreatesEmailsTable()
            _dbManager.Initialize()

            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Assert.IsTrue(TableExists(conn, "emails"))
            End Using
        End Sub

        <Test>
        Public Sub Initialize_CreatesAttachmentsTable()
            _dbManager.Initialize()

            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Assert.IsTrue(TableExists(conn, "attachments"))
            End Using
        End Sub

        <Test>
        Public Sub Initialize_CreatesDeletedMessageIdsTable()
            _dbManager.Initialize()

            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Assert.IsTrue(TableExists(conn, "deleted_message_ids"))
            End Using
        End Sub

        ' ── FTS テーブル ──────────────────────────────────────────────

        <Test>
        Public Sub Initialize_CreatesEmailsFtsTable()
            _dbManager.Initialize()

            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Assert.IsTrue(TableExists(conn, "emails_fts"))
            End Using
        End Sub

        <Test>
        Public Sub Initialize_CreatesAttachmentsFtsTable()
            _dbManager.Initialize()

            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Assert.IsTrue(TableExists(conn, "attachments_fts"))
            End Using
        End Sub

        ' ── インデックス ──────────────────────────────────────────────

        <Test>
        Public Sub Initialize_CreatesIndices()
            _dbManager.Initialize()

            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Assert.IsTrue(IndexExists(conn, "idx_emails_received_at"))
                Assert.IsTrue(IndexExists(conn, "idx_emails_normalized_subject"))
                Assert.IsTrue(IndexExists(conn, "idx_emails_message_id"))
                Assert.IsTrue(IndexExists(conn, "idx_emails_thread_id"))
                Assert.IsTrue(IndexExists(conn, "idx_emails_in_reply_to"))
                Assert.IsTrue(IndexExists(conn, "idx_emails_folder"))
                Assert.IsTrue(IndexExists(conn, "idx_attachments_email_id"))
            End Using
        End Sub

        ' ── トリガー ──────────────────────────────────────────────────

        <Test>
        Public Sub Initialize_CreatesFtsTriggers()
            _dbManager.Initialize()

            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Assert.IsTrue(TriggerExists(conn, "emails_ai"))
                Assert.IsTrue(TriggerExists(conn, "emails_ad"))
                Assert.IsTrue(TriggerExists(conn, "emails_au"))
                Assert.IsTrue(TriggerExists(conn, "attachments_ai"))
                Assert.IsTrue(TriggerExists(conn, "attachments_ad"))
            End Using
        End Sub

        ' ── マイグレーション ──────────────────────────────────────────

        <Test>
        Public Sub Initialize_EmailsTable_HasEmailSizeColumn()
            _dbManager.Initialize()

            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Assert.IsTrue(ColumnExists(conn, "emails", "email_size"))
            End Using
        End Sub

        <Test>
        Public Sub Initialize_AttachmentsTable_HasContentIdColumn()
            _dbManager.Initialize()

            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Assert.IsTrue(ColumnExists(conn, "attachments", "content_id"))
            End Using
        End Sub

        <Test>
        Public Sub Initialize_AttachmentsTable_HasIsInlineColumn()
            _dbManager.Initialize()

            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Assert.IsTrue(ColumnExists(conn, "attachments", "is_inline"))
            End Using
        End Sub

        ' ── 冪等性 ────────────────────────────────────────────────────

        <Test>
        Public Sub Initialize_CalledTwice_NoError()
            _dbManager.Initialize()
            Assert.DoesNotThrow(Sub() _dbManager.Initialize())
        End Sub

        ' ── WAL モード ────────────────────────────────────────────────

        <Test>
        Public Sub Initialize_SetsWalMode()
            _dbManager.Initialize()

            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand("PRAGMA journal_mode;", conn)
                    Dim mode As String = CStr(cmd.ExecuteScalar())
                    Assert.AreEqual("wal", mode.ToLower())
                End Using
            End Using
        End Sub

        ' ── ヘルパーメソッド ──────────────────────────────────────────

        Private Shared Function TableExists(conn As SQLiteConnection, tableName As String) As Boolean
            Dim sql As String = "SELECT COUNT(1) FROM sqlite_master WHERE type IN ('table', 'view') AND name = @name"
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@name", tableName)
                Return Convert.ToInt32(cmd.ExecuteScalar()) > 0
            End Using
        End Function

        Private Shared Function IndexExists(conn As SQLiteConnection, indexName As String) As Boolean
            Dim sql As String = "SELECT COUNT(1) FROM sqlite_master WHERE type = 'index' AND name = @name"
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@name", indexName)
                Return Convert.ToInt32(cmd.ExecuteScalar()) > 0
            End Using
        End Function

        Private Shared Function TriggerExists(conn As SQLiteConnection, triggerName As String) As Boolean
            Dim sql As String = "SELECT COUNT(1) FROM sqlite_master WHERE type = 'trigger' AND name = @name"
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@name", triggerName)
                Return Convert.ToInt32(cmd.ExecuteScalar()) > 0
            End Using
        End Function

        Private Shared Function ColumnExists(conn As SQLiteConnection, tableName As String, columnName As String) As Boolean
            Dim sql As String = String.Format("PRAGMA table_info({0});", tableName)
            Using cmd As New SQLiteCommand(sql, conn)
                Using reader As SQLiteDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        If CStr(reader("name")) = columnName Then Return True
                    End While
                End Using
            End Using
            Return False
        End Function

    End Class

End Namespace
