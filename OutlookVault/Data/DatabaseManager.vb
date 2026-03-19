Option Explicit On
Option Strict On
Option Infer Off

Imports System.Data.SQLite
Imports System.IO

Namespace Data

    ''' <summary>SQLite の synchronous プラグマ設定値。</summary>
    Public Enum SynchronousMode
        Off
        Normal
        Full
    End Enum

    ''' <summary>
    ''' SQLite 接続の管理・スキーマ初期化を担当するクラス。
    ''' </summary>
    Public Class DatabaseManager

        Private ReadOnly _dbPath As String

        Public Sub New(dbPath As String)
            _dbPath = dbPath
        End Sub

        ''' <summary>開いた状態の SQLiteConnection を返す。呼び出し側で Using して閉じること。</summary>
        Public Function GetConnection() As SQLiteConnection
            Dim conn As New SQLiteConnection("Data Source=" & _dbPath & ";Version=3;foreign keys=True;")
            conn.Open()
            Return conn
        End Function

        ''' <summary>
        ''' DB ファイルを作成し、テーブル・インデックスを初期化する。
        ''' 既に存在する場合はスキップ（べき等）。
        ''' </summary>
        Public Sub Initialize()
            Dim dir As String = Path.GetDirectoryName(_dbPath)
            If Not String.IsNullOrEmpty(dir) AndAlso Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            Using conn As SQLiteConnection = GetConnection()
                ' WAL モードで書き込みパフォーマンスを向上
                ExecuteNonQuery(conn, "PRAGMA journal_mode=WAL;")

                CreateTables(conn)
                RunMigrations(conn)
                CreateIndices(conn)
            End Using
        End Sub

        ' ── テーブル ──────────────────────────────────────────────

        Private Sub CreateTables(conn As SQLiteConnection)
            Dim sql As String = "
CREATE TABLE IF NOT EXISTS emails (
    id                 INTEGER PRIMARY KEY AUTOINCREMENT,
    message_id         TEXT,
    in_reply_to        TEXT,
    [references]       TEXT,
    thread_id          TEXT,
    entry_id           TEXT,
    subject            TEXT,
    normalized_subject TEXT,
    sender_name        TEXT,
    sender_email       TEXT,
    to_recipients      TEXT,
    cc_recipients      TEXT,
    bcc_recipients     TEXT,
    body_text          TEXT,
    body_html          TEXT,
    received_at        TEXT NOT NULL,
    sent_at            TEXT,
    folder_name        TEXT,
    has_attachments    INTEGER DEFAULT 0,
    email_size         INTEGER DEFAULT 0,
    created_at         TEXT DEFAULT (datetime('now', 'localtime')),
    updated_at         TEXT DEFAULT (datetime('now', 'localtime')),
    deleted_at         TEXT,
    UNIQUE(message_id, folder_name)
);

CREATE TABLE IF NOT EXISTS attachments (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    email_id   INTEGER NOT NULL REFERENCES emails(id) ON DELETE CASCADE,
    file_name  TEXT NOT NULL,
    file_path  TEXT NOT NULL,
    file_size  INTEGER,
    mime_type  TEXT,
    content_id TEXT,
    is_inline  INTEGER DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now', 'localtime'))
);

CREATE TABLE IF NOT EXISTS deleted_message_ids (
    message_id TEXT PRIMARY KEY,
    deleted_at TEXT DEFAULT (datetime('now', 'localtime'))
);

CREATE TABLE IF NOT EXISTS exchange_address_cache (
    ex_address   TEXT PRIMARY KEY,
    smtp_address TEXT NOT NULL,
    cached_at    TEXT DEFAULT (datetime('now', 'localtime'))
);

CREATE TABLE IF NOT EXISTS error_message_ids (
    message_id    TEXT PRIMARY KEY,
    folder_name   TEXT,
    subject       TEXT,
    error_message TEXT,
    received_date TEXT,
    sender_name   TEXT,
    error_date    TEXT DEFAULT (datetime('now', 'localtime'))
);

CREATE TABLE IF NOT EXISTS folder_sync_state (
    folder_name    TEXT PRIMARY KEY,
    last_sync_time TEXT NOT NULL,
    full_sync_done INTEGER DEFAULT 0,
    updated_at     TEXT DEFAULT (datetime('now', 'localtime'))
);

CREATE TABLE IF NOT EXISTS auto_delete_rules (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    name              TEXT NOT NULL,
    filter_expression TEXT NOT NULL,
    enabled           INTEGER DEFAULT 1,
    sort_order        INTEGER DEFAULT 0,
    created_at        TEXT DEFAULT (datetime('now', 'localtime')),
    updated_at        TEXT DEFAULT (datetime('now', 'localtime'))
);"
            ExecuteNonQuery(conn, sql)
        End Sub


        ' ── インデックス ──────────────────────────────────────────

        Private Sub CreateIndices(conn As SQLiteConnection)
            Dim sql As String = "
CREATE INDEX IF NOT EXISTS idx_emails_received_at        ON emails(received_at DESC);
CREATE INDEX IF NOT EXISTS idx_emails_normalized_subject ON emails(normalized_subject);
CREATE INDEX IF NOT EXISTS idx_emails_message_id         ON emails(message_id);
CREATE INDEX IF NOT EXISTS idx_emails_thread_id          ON emails(thread_id);
CREATE INDEX IF NOT EXISTS idx_emails_in_reply_to        ON emails(in_reply_to);
CREATE INDEX IF NOT EXISTS idx_emails_folder             ON emails(folder_name);
CREATE INDEX IF NOT EXISTS idx_attachments_email_id      ON attachments(email_id);
CREATE INDEX IF NOT EXISTS idx_emails_deleted_at         ON emails(deleted_at);"
            ExecuteNonQuery(conn, sql)
        End Sub

        ' ── マイグレーション ────────────────────────────────────────

        ''' <summary>既存 DB に対するスキーマ変更を適用する（べき等）。</summary>
        Private Sub RunMigrations(conn As SQLiteConnection)
            ' v1.x → ゴミ箱機能: emails.deleted_at カラム追加
            If Not ColumnExists(conn, "emails", "deleted_at") Then
                ExecuteNonQuery(conn, "ALTER TABLE emails ADD COLUMN deleted_at TEXT;")
            End If

            ' v1.x → 同一 MessageID のフォルダ別保存: UNIQUE(message_id) → UNIQUE(message_id, folder_name)
            ' SQLite はカラム制約の UNIQUE を直接削除できないため、テーブルを再作成する
            If HasUniqueConstraintOnMessageIdOnly(conn) Then
                MigrateMessageIdUnique(conn)
            End If
        End Sub

        ''' <summary>emails テーブルの message_id カラムに単独 UNIQUE 制約があるか確認する。</summary>
        Private Shared Function HasUniqueConstraintOnMessageIdOnly(conn As SQLiteConnection) As Boolean
            ' index_list で UNIQUE インデックスを列挙し、message_id 単独のものがあるか確認
            Using cmd As New SQLiteCommand("PRAGMA index_list(emails);", conn)
                Using reader As SQLiteDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim indexName As String = reader("name").ToString()
                        Dim isUnique As Boolean = Convert.ToInt32(reader("unique")) = 1
                        If Not isUnique Then Continue While

                        ' このインデックスの列を確認
                        Dim columns As New List(Of String)()
                        Using cmd2 As New SQLiteCommand(String.Format("PRAGMA index_info({0});", indexName), conn)
                            Using reader2 As SQLiteDataReader = cmd2.ExecuteReader()
                                While reader2.Read()
                                    columns.Add(reader2("name").ToString())
                                End While
                            End Using
                        End Using

                        ' message_id 単独の UNIQUE インデックスか
                        If columns.Count = 1 AndAlso
                           String.Equals(columns(0), "message_id", StringComparison.OrdinalIgnoreCase) Then
                            Return True
                        End If
                    End While
                End Using
            End Using
            Return False
        End Function

        ''' <summary>emails テーブルを再作成して UNIQUE(message_id, folder_name) に変更する。</summary>
        Private Sub MigrateMessageIdUnique(conn As SQLiteConnection)
            Services.Logger.Info("マイグレーション: emails テーブルの UNIQUE 制約を (message_id, folder_name) に変更します")
            Dim sql As String = "
BEGIN TRANSACTION;
CREATE TABLE emails_new (
    id                 INTEGER PRIMARY KEY AUTOINCREMENT,
    message_id         TEXT,
    in_reply_to        TEXT,
    [references]       TEXT,
    thread_id          TEXT,
    entry_id           TEXT,
    subject            TEXT,
    normalized_subject TEXT,
    sender_name        TEXT,
    sender_email       TEXT,
    to_recipients      TEXT,
    cc_recipients      TEXT,
    bcc_recipients     TEXT,
    body_text          TEXT,
    body_html          TEXT,
    received_at        TEXT NOT NULL,
    sent_at            TEXT,
    folder_name        TEXT,
    has_attachments    INTEGER DEFAULT 0,
    email_size         INTEGER DEFAULT 0,
    created_at         TEXT DEFAULT (datetime('now', 'localtime')),
    updated_at         TEXT DEFAULT (datetime('now', 'localtime')),
    deleted_at         TEXT,
    UNIQUE(message_id, folder_name)
);
INSERT INTO emails_new SELECT * FROM emails;
DROP TABLE emails;
ALTER TABLE emails_new RENAME TO emails;
COMMIT;"
            ExecuteNonQuery(conn, sql)
            ' インデックスはテーブル再作成で消えるが CreateIndices で再作成される
            Services.Logger.Info("マイグレーション: UNIQUE 制約の変更が完了しました")
        End Sub

        ''' <summary>指定テーブルに指定カラムが存在するか確認する。</summary>
        Private Shared Function ColumnExists(conn As SQLiteConnection, tableName As String, columnName As String) As Boolean
            Using cmd As New SQLiteCommand(String.Format("PRAGMA table_info({0});", tableName), conn)
                Using reader As SQLiteDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        If String.Equals(reader("name").ToString(), columnName, StringComparison.OrdinalIgnoreCase) Then
                            Return True
                        End If
                    End While
                End Using
            End Using
            Return False
        End Function

        ' ── パフォーマンスチューニング ─────────────────────────────

        ''' <summary>
        ''' synchronous プラグマを設定する。
        ''' 取り込み中は Off にして fsync を省略し高速化、終了後に Normal に戻す。
        ''' </summary>
        Public Sub SetSynchronousMode(conn As SQLiteConnection, mode As SynchronousMode)
            ExecuteNonQuery(conn, "PRAGMA synchronous=" & mode.ToString().ToUpperInvariant() & ";")
        End Sub

        ' ── テーブルデータ取得 ────────────────────────────────────

        ''' <summary>指定テーブルの全行を DataTable で返す。長文列は先頭100文字に切り詰める。</summary>
        Public Function GetTableData(tableName As String) As System.Data.DataTable
            ' テーブル名をホワイトリストで検証（SQLインジェクション防止）
            Dim allowed() As String = {"emails", "attachments", "deleted_message_ids", "exchange_address_cache", "error_message_ids", "folder_sync_state", "auto_delete_rules"}
            If Array.IndexOf(allowed, tableName) < 0 Then
                Throw New ArgumentException("無効なテーブル名: " & tableName)
            End If

            Dim sql As String = BuildTableDataQuery(tableName)

            Dim dt As New System.Data.DataTable()
            Using conn As SQLiteConnection = GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using
            End Using
            Return dt
        End Function

        ''' <summary>テーブルビューア用の SELECT 文を生成する。長文列は SUBSTR で切り詰める。</summary>
        Private Shared Function BuildTableDataQuery(tableName As String) As String
            If tableName <> "emails" Then
                Return "SELECT * FROM " & tableName
            End If

            ' emails テーブル: 長文列を先頭100文字に切り詰め
            Const TruncLen As Integer = 100
            Return String.Format(
                "SELECT id, message_id, in_reply_to, [references], thread_id, entry_id, " &
                "subject, normalized_subject, sender_name, sender_email, " &
                "SUBSTR(to_recipients, 1, {0}) AS to_recipients, " &
                "SUBSTR(cc_recipients, 1, {0}) AS cc_recipients, " &
                "SUBSTR(bcc_recipients, 1, {0}) AS bcc_recipients, " &
                "SUBSTR(body_text, 1, {0}) AS body_text, " &
                "SUBSTR(body_html, 1, {0}) AS body_html, " &
                "received_at, sent_at, folder_name, has_attachments, email_size, " &
                "created_at, updated_at, deleted_at FROM emails", TruncLen)
        End Function

        ' ── ユーティリティ ────────────────────────────────────────

        Private Sub ExecuteNonQuery(conn As SQLiteConnection, sql As String)
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.ExecuteNonQuery()
            End Using
        End Sub

    End Class

End Namespace
