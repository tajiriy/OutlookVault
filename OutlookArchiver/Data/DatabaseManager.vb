Option Explicit On
Option Strict On
Option Infer Off

Imports System.Data.SQLite
Imports System.IO

Namespace Data

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
        ''' DB ファイルを作成し、テーブル・インデックス・FTS5仮想テーブル・トリガーを初期化する。
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
                CreateIndices(conn)
                CreateFtsTables(conn)
                CreateFtsTriggers(conn)
            End Using
        End Sub

        ' ── テーブル ──────────────────────────────────────────────

        Private Sub CreateTables(conn As SQLiteConnection)
            Dim sql As String = "
CREATE TABLE IF NOT EXISTS emails (
    id                 INTEGER PRIMARY KEY AUTOINCREMENT,
    message_id         TEXT UNIQUE,
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
    created_at         TEXT DEFAULT (datetime('now', 'localtime')),
    updated_at         TEXT DEFAULT (datetime('now', 'localtime'))
);

CREATE TABLE IF NOT EXISTS attachments (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    email_id   INTEGER NOT NULL REFERENCES emails(id) ON DELETE CASCADE,
    file_name  TEXT NOT NULL,
    file_path  TEXT NOT NULL,
    file_size  INTEGER,
    mime_type  TEXT,
    created_at TEXT DEFAULT (datetime('now', 'localtime'))
);

CREATE TABLE IF NOT EXISTS deleted_message_ids (
    message_id TEXT PRIMARY KEY,
    deleted_at TEXT DEFAULT (datetime('now', 'localtime'))
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
CREATE INDEX IF NOT EXISTS idx_attachments_email_id      ON attachments(email_id);"
            ExecuteNonQuery(conn, sql)
        End Sub

        ' ── FTS4 仮想テーブル ─────────────────────────────────────
        ' FTS5 は SQLite.Interop.dll のビルドによっては無効な場合があるため FTS4 を使用する。
        ' FTS4 は content= でコンテンツテーブルを指定でき、rowid で紐付けられる。

        Private Sub CreateFtsTables(conn As SQLiteConnection)
            Dim sql As String = "
CREATE VIRTUAL TABLE IF NOT EXISTS emails_fts USING fts4(
    content=""emails"",
    subject,
    body_text,
    sender_name,
    sender_email
);

CREATE VIRTUAL TABLE IF NOT EXISTS attachments_fts USING fts4(
    content=""attachments"",
    file_name
);"
            ExecuteNonQuery(conn, sql)
        End Sub

        ' ── FTS4 同期トリガー ─────────────────────────────────────
        ' FTS4 の削除は DELETE FROM ... WHERE rowid = を使う（FTS5 の 'delete' 特殊構文は不要）。

        Private Sub CreateFtsTriggers(conn As SQLiteConnection)
            ' emails_fts を emails テーブルと同期するトリガー
            Dim sqlInsert As String = "
CREATE TRIGGER IF NOT EXISTS emails_ai AFTER INSERT ON emails BEGIN
    INSERT INTO emails_fts(rowid, subject, body_text, sender_name, sender_email)
    VALUES (new.id, new.subject, new.body_text, new.sender_name, new.sender_email);
END;"
            Dim sqlDelete As String = "
CREATE TRIGGER IF NOT EXISTS emails_ad AFTER DELETE ON emails BEGIN
    DELETE FROM emails_fts WHERE rowid = old.id;
END;"
            Dim sqlUpdate As String = "
CREATE TRIGGER IF NOT EXISTS emails_au AFTER UPDATE ON emails BEGIN
    DELETE FROM emails_fts WHERE rowid = old.id;
    INSERT INTO emails_fts(rowid, subject, body_text, sender_name, sender_email)
    VALUES (new.id, new.subject, new.body_text, new.sender_name, new.sender_email);
END;"
            Dim sqlAttachInsert As String = "
CREATE TRIGGER IF NOT EXISTS attachments_ai AFTER INSERT ON attachments BEGIN
    INSERT INTO attachments_fts(rowid, file_name)
    VALUES (new.id, new.file_name);
END;"
            Dim sqlAttachDelete As String = "
CREATE TRIGGER IF NOT EXISTS attachments_ad AFTER DELETE ON attachments BEGIN
    DELETE FROM attachments_fts WHERE rowid = old.id;
END;"
            ExecuteNonQuery(conn, sqlInsert)
            ExecuteNonQuery(conn, sqlDelete)
            ExecuteNonQuery(conn, sqlUpdate)
            ExecuteNonQuery(conn, sqlAttachInsert)
            ExecuteNonQuery(conn, sqlAttachDelete)
        End Sub

        ' ── ユーティリティ ────────────────────────────────────────

        Private Sub ExecuteNonQuery(conn As SQLiteConnection, sql As String)
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.ExecuteNonQuery()
            End Using
        End Sub

    End Class

End Namespace
