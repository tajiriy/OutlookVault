Option Explicit On
Option Strict On
Option Infer Off

Imports System.Data.SQLite
Imports System.Text

Namespace Data

    ''' <summary>
    ''' emails / attachments / deleted_message_ids テーブルへのアクセスを担当するリポジトリクラス。
    ''' </summary>
    Public Class EmailRepository

        Private ReadOnly _dbManager As DatabaseManager

        Public Sub New(dbManager As DatabaseManager)
            _dbManager = dbManager
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  Email 挿入・更新
        ' ════════════════════════════════════════════════════════════

        ''' <summary>メールを DB に挿入し、採番された ID を返す。</summary>
        Public Function InsertEmail(email As Models.Email) As Integer
            Const sql As String = "
INSERT INTO emails (
    message_id, in_reply_to, [references], thread_id, entry_id,
    subject, normalized_subject, sender_name, sender_email,
    to_recipients, cc_recipients, bcc_recipients,
    body_text, body_html, received_at, sent_at, folder_name, has_attachments
) VALUES (
    @message_id, @in_reply_to, @references, @thread_id, @entry_id,
    @subject, @normalized_subject, @sender_name, @sender_email,
    @to_recipients, @cc_recipients, @bcc_recipients,
    @body_text, @body_html, @received_at, @sent_at, @folder_name, @has_attachments
);
SELECT last_insert_rowid();"

            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@message_id", NullableStr(email.MessageId))
                    cmd.Parameters.AddWithValue("@in_reply_to", NullableStr(email.InReplyTo))
                    cmd.Parameters.AddWithValue("@references", NullableStr(email.References))
                    cmd.Parameters.AddWithValue("@thread_id", NullableStr(email.ThreadId))
                    cmd.Parameters.AddWithValue("@entry_id", NullableStr(email.EntryId))
                    cmd.Parameters.AddWithValue("@subject", NullableStr(email.Subject))
                    cmd.Parameters.AddWithValue("@normalized_subject", NullableStr(email.NormalizedSubject))
                    cmd.Parameters.AddWithValue("@sender_name", NullableStr(email.SenderName))
                    cmd.Parameters.AddWithValue("@sender_email", NullableStr(email.SenderEmail))
                    cmd.Parameters.AddWithValue("@to_recipients", NullableStr(email.ToRecipients))
                    cmd.Parameters.AddWithValue("@cc_recipients", NullableStr(email.CcRecipients))
                    cmd.Parameters.AddWithValue("@bcc_recipients", NullableStr(email.BccRecipients))
                    cmd.Parameters.AddWithValue("@body_text", NullableStr(email.BodyText))
                    cmd.Parameters.AddWithValue("@body_html", NullableStr(email.BodyHtml))
                    cmd.Parameters.AddWithValue("@received_at", email.ReceivedAt.ToString("o"))
                    cmd.Parameters.AddWithValue("@sent_at",
                        If(email.SentAt.HasValue,
                           CType(email.SentAt.Value.ToString("o"), Object),
                           CType(DBNull.Value, Object)))
                    cmd.Parameters.AddWithValue("@folder_name", NullableStr(email.FolderName))
                    cmd.Parameters.AddWithValue("@has_attachments", CType(If(email.HasAttachments, 1, 0), Object))
                    Return Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            End Using
        End Function

        ' ════════════════════════════════════════════════════════════
        '  Attachment 挿入
        ' ════════════════════════════════════════════════════════════

        ''' <summary>添付ファイルメタデータを DB に挿入し、採番された ID を返す。</summary>
        Public Function InsertAttachment(attachment As Models.Attachment) As Integer
            Const sql As String = "
INSERT INTO attachments (email_id, file_name, file_path, file_size, mime_type)
VALUES (@email_id, @file_name, @file_path, @file_size, @mime_type);
SELECT last_insert_rowid();"

            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@email_id", CType(attachment.EmailId, Object))
                    cmd.Parameters.AddWithValue("@file_name", attachment.FileName)
                    cmd.Parameters.AddWithValue("@file_path", attachment.FilePath)
                    cmd.Parameters.AddWithValue("@file_size", CType(attachment.FileSize, Object))
                    cmd.Parameters.AddWithValue("@mime_type", NullableStr(attachment.MimeType))
                    Return Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            End Using
        End Function

        ' ════════════════════════════════════════════════════════════
        '  Email 取得
        ' ════════════════════════════════════════════════════════════

        ''' <summary>フォルダ指定・ページング付きでメール一覧を取得する（受信日時降順）。</summary>
        Public Function GetEmails(Optional folderName As String = Nothing,
                                  Optional pageIndex As Integer = 0,
                                  Optional pageSize As Integer = 100) As List(Of Models.Email)
            Dim sb As New StringBuilder("SELECT * FROM emails")
            If Not String.IsNullOrEmpty(folderName) Then
                sb.Append(" WHERE folder_name = @folder_name")
            End If
            sb.Append(" ORDER BY received_at DESC")
            sb.AppendFormat(" LIMIT {0} OFFSET {1}", pageSize, pageIndex * pageSize)

            Dim result As New List(Of Models.Email)
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sb.ToString(), conn)
                    If Not String.IsNullOrEmpty(folderName) Then
                        cmd.Parameters.AddWithValue("@folder_name", folderName)
                    End If
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            result.Add(MapEmail(reader))
                        End While
                    End Using
                End Using
            End Using
            Return result
        End Function

        ''' <summary>ID でメールを1件取得し、添付ファイル一覧もセットして返す。</summary>
        Public Function GetEmailById(id As Integer) As Models.Email
            Const sql As String = "SELECT * FROM emails WHERE id = @id"
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@id", CType(id, Object))
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Dim email As Models.Email = MapEmail(reader)
                            email.Attachments = GetAttachmentsByEmailId(id)
                            Return email
                        End If
                    End Using
                End Using
            End Using
            Return Nothing
        End Function

        ''' <summary>同一スレッドのメールを時系列昇順で取得する。</summary>
        Public Function GetEmailsByThreadId(threadId As String) As List(Of Models.Email)
            Const sql As String = "SELECT * FROM emails WHERE thread_id = @thread_id ORDER BY received_at ASC"
            Dim result As New List(Of Models.Email)
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@thread_id", threadId)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            result.Add(MapEmail(reader))
                        End While
                    End Using
                End Using
            End Using
            Return result
        End Function

        ''' <summary>DB に存在するフォルダ名の一覧を取得する。</summary>
        Public Function GetFolderNames() As List(Of String)
            Const sql As String = "SELECT DISTINCT folder_name FROM emails WHERE folder_name IS NOT NULL ORDER BY folder_name"
            Dim result As New List(Of String)
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            result.Add(reader.GetString(0))
                        End While
                    End Using
                End Using
            End Using
            Return result
        End Function

        ''' <summary>フォルダ指定可能なメール総数を返す。</summary>
        Public Function GetTotalCount(Optional folderName As String = Nothing) As Integer
            Dim sql As String
            If String.IsNullOrEmpty(folderName) Then
                sql = "SELECT COUNT(1) FROM emails"
            Else
                sql = "SELECT COUNT(1) FROM emails WHERE folder_name = @folder_name"
            End If
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    If Not String.IsNullOrEmpty(folderName) Then
                        cmd.Parameters.AddWithValue("@folder_name", folderName)
                    End If
                    Return Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            End Using
        End Function

        ' ════════════════════════════════════════════════════════════
        '  Attachment 取得
        ' ════════════════════════════════════════════════════════════

        ''' <summary>指定メール ID の添付ファイル一覧を返す。</summary>
        Public Function GetAttachmentsByEmailId(emailId As Integer) As List(Of Models.Attachment)
            Const sql As String = "SELECT * FROM attachments WHERE email_id = @email_id"
            Dim result As New List(Of Models.Attachment)
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@email_id", CType(emailId, Object))
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            result.Add(MapAttachment(reader))
                        End While
                    End Using
                End Using
            End Using
            Return result
        End Function

        ''' <summary>MessageID でメールを1件取得する。スレッド判定に使用。</summary>
        Public Function GetEmailByMessageId(messageId As String) As Models.Email
            Const sql As String = "SELECT * FROM emails WHERE message_id = @message_id LIMIT 1"
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@message_id", messageId)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        If reader.Read() Then Return MapEmail(reader)
                    End Using
                End Using
            End Using
            Return Nothing
        End Function

        ''' <summary>正規化済み件名が一致するメールを受信日時昇順で最大 limit 件取得する。</summary>
        Public Function GetEmailsByNormalizedSubject(normalizedSubject As String,
                                                     limit As Integer) As List(Of Models.Email)
            Const sql As String = "SELECT * FROM emails WHERE normalized_subject = @normalized_subject ORDER BY received_at ASC LIMIT @limit"
            Dim result As New List(Of Models.Email)
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@normalized_subject", normalizedSubject)
                    cmd.Parameters.AddWithValue("@limit", CType(limit, Object))
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            result.Add(MapEmail(reader))
                        End While
                    End Using
                End Using
            End Using
            Return result
        End Function

        ' ════════════════════════════════════════════════════════════
        '  全文検索（FTS4）
        ' ════════════════════════════════════════════════════════════

        ''' <summary>FTS4 でメールを全文検索する。結果は受信日時降順。</summary>
        Public Function SearchEmails(query As String,
                                     Optional folderName As String = Nothing) As List(Of Models.Email)
            Dim sb As New StringBuilder(
                "SELECT e.* FROM emails e INNER JOIN emails_fts f ON e.id = f.rowid WHERE emails_fts MATCH @query")
            If Not String.IsNullOrEmpty(folderName) Then
                sb.Append(" AND e.folder_name = @folder_name")
            End If
            sb.Append(" ORDER BY e.received_at DESC")

            Dim result As New List(Of Models.Email)
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sb.ToString(), conn)
                    cmd.Parameters.AddWithValue("@query", query)
                    If Not String.IsNullOrEmpty(folderName) Then
                        cmd.Parameters.AddWithValue("@folder_name", folderName)
                    End If
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            result.Add(MapEmail(reader))
                        End While
                    End Using
                End Using
            End Using
            Return result
        End Function

        ' ════════════════════════════════════════════════════════════
        '  削除・重複チェック
        ' ════════════════════════════════════════════════════════════

        ''' <summary>同一 MessageID がすでに DB に存在するか確認する（重複取り込み防止）。</summary>
        Public Function MessageIdExists(messageId As String) As Boolean
            Const sql As String = "SELECT COUNT(1) FROM emails WHERE message_id = @message_id"
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@message_id", messageId)
                    Return Convert.ToInt32(cmd.ExecuteScalar()) > 0
                End Using
            End Using
        End Function

        ''' <summary>MessageID がトゥームストーン（削除済み）テーブルに存在するか確認する。</summary>
        Public Function IsMessageIdDeleted(messageId As String) As Boolean
            Const sql As String = "SELECT COUNT(1) FROM deleted_message_ids WHERE message_id = @message_id"
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@message_id", messageId)
                    Return Convert.ToInt32(cmd.ExecuteScalar()) > 0
                End Using
            End Using
        End Function

        ''' <summary>
        ''' メールを DB から削除し、MessageID をトゥームストーンに記録する。
        ''' ON DELETE CASCADE で添付ファイルレコードも自動削除される。
        ''' </summary>
        Public Sub DeleteEmail(id As Integer)
            ' まず MessageID を取得してトゥームストーンに登録
            Dim email As Models.Email = GetEmailById(id)
            If email IsNot Nothing AndAlso Not String.IsNullOrEmpty(email.MessageId) Then
                MarkMessageIdAsDeleted(email.MessageId)
            End If

            Const sql As String = "DELETE FROM emails WHERE id = @id"
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@id", CType(id, Object))
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        ''' <summary>MessageID をトゥームストーンテーブルに記録する（再取り込み防止）。</summary>
        Public Sub MarkMessageIdAsDeleted(messageId As String)
            Const sql As String = "INSERT OR IGNORE INTO deleted_message_ids (message_id) VALUES (@message_id)"
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@message_id", messageId)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  マッピング・ユーティリティ
        ' ════════════════════════════════════════════════════════════

        Private Function MapEmail(reader As SQLiteDataReader) As Models.Email
            Dim email As New Models.Email()
            email.Id = reader.GetInt32(reader.GetOrdinal("id"))
            email.MessageId = GetStr(reader, "message_id")
            email.InReplyTo = GetStr(reader, "in_reply_to")
            email.References = GetStr(reader, "references")
            email.ThreadId = GetStr(reader, "thread_id")
            email.EntryId = GetStr(reader, "entry_id")
            email.Subject = GetStr(reader, "subject")
            email.NormalizedSubject = GetStr(reader, "normalized_subject")
            email.SenderName = GetStr(reader, "sender_name")
            email.SenderEmail = GetStr(reader, "sender_email")
            email.ToRecipients = GetStr(reader, "to_recipients")
            email.CcRecipients = GetStr(reader, "cc_recipients")
            email.BccRecipients = GetStr(reader, "bcc_recipients")
            email.BodyText = GetStr(reader, "body_text")
            email.BodyHtml = GetStr(reader, "body_html")
            email.ReceivedAt = DateTime.Parse(reader.GetString(reader.GetOrdinal("received_at")))
            Dim sentStr As String = GetStr(reader, "sent_at")
            If Not String.IsNullOrEmpty(sentStr) Then email.SentAt = DateTime.Parse(sentStr)
            email.FolderName = GetStr(reader, "folder_name")
            email.HasAttachments = reader.GetInt32(reader.GetOrdinal("has_attachments")) = 1
            Return email
        End Function

        Private Function MapAttachment(reader As SQLiteDataReader) As Models.Attachment
            Dim att As New Models.Attachment()
            att.Id = reader.GetInt32(reader.GetOrdinal("id"))
            att.EmailId = reader.GetInt32(reader.GetOrdinal("email_id"))
            att.FileName = reader.GetString(reader.GetOrdinal("file_name"))
            att.FilePath = reader.GetString(reader.GetOrdinal("file_path"))
            Dim sizeOrd As Integer = reader.GetOrdinal("file_size")
            If Not reader.IsDBNull(sizeOrd) Then att.FileSize = reader.GetInt64(sizeOrd)
            att.MimeType = GetStr(reader, "mime_type")
            Return att
        End Function

        ''' <summary>NULL の列を Nothing として返すヘルパー。</summary>
        Private Function GetStr(reader As SQLiteDataReader, column As String) As String
            Dim ord As Integer = reader.GetOrdinal(column)
            If reader.IsDBNull(ord) Then Return Nothing
            Return reader.GetString(ord)
        End Function

        ''' <summary>Nothing の文字列を DBNull.Value に変換するヘルパー。</summary>
        Private Function NullableStr(value As String) As Object
            If value Is Nothing Then Return CType(DBNull.Value, Object)
            Return CType(value, Object)
        End Function

    End Class

End Namespace
