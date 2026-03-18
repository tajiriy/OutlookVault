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

        ' ── バルク書き込み用（トランザクション共有） ──────────────────
        Private _bulkConn As SQLiteConnection = Nothing
        Private _bulkTx As SQLiteTransaction = Nothing

        ' ── Prepared Statement 再利用 ────────────────────────────────
        Private _bulkEmailCmd As SQLiteCommand = Nothing
        Private _bulkAttachCmd As SQLiteCommand = Nothing

        Public Sub New(dbManager As DatabaseManager)
            _dbManager = dbManager
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  バルク書き込みモード（トランザクション管理）
        ' ════════════════════════════════════════════════════════════

        ''' <summary>バルク書き込みモードを開始する。呼び出し後の Insert は共有トランザクション上で実行される。</summary>
        Public Sub BeginBulk()
            If _bulkConn IsNot Nothing Then Return  ' 既に開始済み
            _bulkConn = _dbManager.GetConnection()
            _bulkTx = _bulkConn.BeginTransaction()
            _bulkEmailCmd = CreateEmailInsertCommand(_bulkConn)
            _bulkAttachCmd = CreateAttachmentInsertCommand(_bulkConn)
        End Sub

        ''' <summary>バルクトランザクションをコミットして接続を閉じる。</summary>
        Public Sub CommitBulk()
            DisposeBulkCommands()
            If _bulkTx IsNot Nothing Then
                _bulkTx.Commit()
                _bulkTx.Dispose()
                _bulkTx = Nothing
            End If
            If _bulkConn IsNot Nothing Then
                _bulkConn.Dispose()
                _bulkConn = Nothing
            End If
        End Sub

        ''' <summary>バルクトランザクションをロールバックして接続を閉じる。</summary>
        Public Sub RollbackBulk()
            DisposeBulkCommands()
            If _bulkTx IsNot Nothing Then
                _bulkTx.Rollback()
                _bulkTx.Dispose()
                _bulkTx = Nothing
            End If
            If _bulkConn IsNot Nothing Then
                _bulkConn.Dispose()
                _bulkConn = Nothing
            End If
        End Sub

        ''' <summary>Prepared Statement を破棄する。</summary>
        Private Sub DisposeBulkCommands()
            If _bulkEmailCmd IsNot Nothing Then
                _bulkEmailCmd.Dispose()
                _bulkEmailCmd = Nothing
            End If
            If _bulkAttachCmd IsNot Nothing Then
                _bulkAttachCmd.Dispose()
                _bulkAttachCmd = Nothing
            End If
        End Sub

        ''' <summary>バルクモードが有効かどうかを返す。</summary>
        Public ReadOnly Property IsBulkActive As Boolean
            Get
                Return _bulkConn IsNot Nothing
            End Get
        End Property

        ' ════════════════════════════════════════════════════════════
        '  Email 挿入・更新
        ' ════════════════════════════════════════════════════════════

        Private Const EmailInsertSql As String = "
INSERT INTO emails (
    message_id, in_reply_to, [references], thread_id, entry_id,
    subject, normalized_subject, sender_name, sender_email,
    to_recipients, cc_recipients, bcc_recipients,
    body_text, body_html, received_at, sent_at, folder_name, has_attachments, email_size
) VALUES (
    @message_id, @in_reply_to, @references, @thread_id, @entry_id,
    @subject, @normalized_subject, @sender_name, @sender_email,
    @to_recipients, @cc_recipients, @bcc_recipients,
    @body_text, @body_html, @received_at, @sent_at, @folder_name, @has_attachments, @email_size
);
SELECT last_insert_rowid();"

        ''' <summary>Email INSERT 用 Prepared Statement を作成する。</summary>
        Private Shared Function CreateEmailInsertCommand(conn As SQLiteConnection) As SQLiteCommand
            Dim cmd As New SQLiteCommand(EmailInsertSql, conn)
            cmd.Parameters.Add("@message_id", System.Data.DbType.String)
            cmd.Parameters.Add("@in_reply_to", System.Data.DbType.String)
            cmd.Parameters.Add("@references", System.Data.DbType.String)
            cmd.Parameters.Add("@thread_id", System.Data.DbType.String)
            cmd.Parameters.Add("@entry_id", System.Data.DbType.String)
            cmd.Parameters.Add("@subject", System.Data.DbType.String)
            cmd.Parameters.Add("@normalized_subject", System.Data.DbType.String)
            cmd.Parameters.Add("@sender_name", System.Data.DbType.String)
            cmd.Parameters.Add("@sender_email", System.Data.DbType.String)
            cmd.Parameters.Add("@to_recipients", System.Data.DbType.String)
            cmd.Parameters.Add("@cc_recipients", System.Data.DbType.String)
            cmd.Parameters.Add("@bcc_recipients", System.Data.DbType.String)
            cmd.Parameters.Add("@body_text", System.Data.DbType.String)
            cmd.Parameters.Add("@body_html", System.Data.DbType.String)
            cmd.Parameters.Add("@received_at", System.Data.DbType.String)
            cmd.Parameters.Add("@sent_at", System.Data.DbType.String)
            cmd.Parameters.Add("@folder_name", System.Data.DbType.String)
            cmd.Parameters.Add("@has_attachments", System.Data.DbType.Int32)
            cmd.Parameters.Add("@email_size", System.Data.DbType.Int64)
            Return cmd
        End Function

        ''' <summary>メールを DB に挿入し、採番された ID を返す。バルクモード中は共有トランザクションを使用する。</summary>
        Public Function InsertEmail(email As Models.Email) As Integer
            If _bulkConn IsNot Nothing Then
                Return ExecuteEmailInsert(_bulkEmailCmd, email)
            End If
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As SQLiteCommand = CreateEmailInsertCommand(conn)
                    Return ExecuteEmailInsert(cmd, email)
                End Using
            End Using
        End Function

        ''' <summary>Email INSERT コマンドにパラメータをセットして実行する。</summary>
        Private Shared Function ExecuteEmailInsert(cmd As SQLiteCommand, email As Models.Email) As Integer
            cmd.Parameters("@message_id").Value = NullableStr(email.MessageId)
            cmd.Parameters("@in_reply_to").Value = NullableStr(email.InReplyTo)
            cmd.Parameters("@references").Value = NullableStr(email.References)
            cmd.Parameters("@thread_id").Value = NullableStr(email.ThreadId)
            cmd.Parameters("@entry_id").Value = NullableStr(email.EntryId)
            cmd.Parameters("@subject").Value = NullableStr(email.Subject)
            cmd.Parameters("@normalized_subject").Value = NullableStr(email.NormalizedSubject)
            cmd.Parameters("@sender_name").Value = NullableStr(email.SenderName)
            cmd.Parameters("@sender_email").Value = NullableStr(email.SenderEmail)
            cmd.Parameters("@to_recipients").Value = NullableStr(email.ToRecipients)
            cmd.Parameters("@cc_recipients").Value = NullableStr(email.CcRecipients)
            cmd.Parameters("@bcc_recipients").Value = NullableStr(email.BccRecipients)
            cmd.Parameters("@body_text").Value = NullableStr(email.BodyText)
            cmd.Parameters("@body_html").Value = NullableStr(email.BodyHtml)
            cmd.Parameters("@received_at").Value = email.ReceivedAt.ToString("o")
            cmd.Parameters("@sent_at").Value =
                If(email.SentAt.HasValue,
                   CType(email.SentAt.Value.ToString("o"), Object),
                   CType(DBNull.Value, Object))
            cmd.Parameters("@folder_name").Value = NullableStr(email.FolderName)
            cmd.Parameters("@has_attachments").Value = CType(If(email.HasAttachments, 1, 0), Object)
            cmd.Parameters("@email_size").Value = CType(email.EmailSize, Object)
            Return Convert.ToInt32(cmd.ExecuteScalar())
        End Function

        ' ════════════════════════════════════════════════════════════
        '  Attachment 挿入
        ' ════════════════════════════════════════════════════════════

        Private Const AttachmentInsertSql As String = "
INSERT INTO attachments (email_id, file_name, file_path, file_size, mime_type, content_id, is_inline)
VALUES (@email_id, @file_name, @file_path, @file_size, @mime_type, @content_id, @is_inline);
SELECT last_insert_rowid();"

        ''' <summary>Attachment INSERT 用 Prepared Statement を作成する。</summary>
        Private Shared Function CreateAttachmentInsertCommand(conn As SQLiteConnection) As SQLiteCommand
            Dim cmd As New SQLiteCommand(AttachmentInsertSql, conn)
            cmd.Parameters.Add("@email_id", System.Data.DbType.Int32)
            cmd.Parameters.Add("@file_name", System.Data.DbType.String)
            cmd.Parameters.Add("@file_path", System.Data.DbType.String)
            cmd.Parameters.Add("@file_size", System.Data.DbType.Int64)
            cmd.Parameters.Add("@mime_type", System.Data.DbType.String)
            cmd.Parameters.Add("@content_id", System.Data.DbType.String)
            cmd.Parameters.Add("@is_inline", System.Data.DbType.Int32)
            Return cmd
        End Function

        ''' <summary>添付ファイルメタデータを DB に挿入し、採番された ID を返す。バルクモード中は共有トランザクションを使用する。</summary>
        Public Function InsertAttachment(attachment As Models.Attachment) As Integer
            If _bulkConn IsNot Nothing Then
                Return ExecuteAttachmentInsert(_bulkAttachCmd, attachment)
            End If
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As SQLiteCommand = CreateAttachmentInsertCommand(conn)
                    Return ExecuteAttachmentInsert(cmd, attachment)
                End Using
            End Using
        End Function

        ''' <summary>Attachment INSERT コマンドにパラメータをセットして実行する。</summary>
        Private Shared Function ExecuteAttachmentInsert(cmd As SQLiteCommand, attachment As Models.Attachment) As Integer
            cmd.Parameters("@email_id").Value = CType(attachment.EmailId, Object)
            cmd.Parameters("@file_name").Value = attachment.FileName
            cmd.Parameters("@file_path").Value = attachment.FilePath
            cmd.Parameters("@file_size").Value = CType(attachment.FileSize, Object)
            cmd.Parameters("@mime_type").Value = NullableStr(attachment.MimeType)
            cmd.Parameters("@content_id").Value = NullableStr(attachment.ContentId)
            cmd.Parameters("@is_inline").Value = CType(If(attachment.IsInline, 1, 0), Object)
            Return Convert.ToInt32(cmd.ExecuteScalar())
        End Function

        ''' <summary>指定メールの has_attachments フラグを更新する。</summary>
        Public Sub UpdateHasAttachments(emailId As Integer, hasAttachments As Boolean)
            Dim sql As String = "UPDATE emails SET has_attachments = @val WHERE id = @id"
            If _bulkConn IsNot Nothing Then
                Using cmd As New SQLiteCommand(sql, _bulkConn)
                    cmd.Transaction = _bulkTx
                    cmd.Parameters.AddWithValue("@val", If(hasAttachments, 1, 0))
                    cmd.Parameters.AddWithValue("@id", emailId)
                    cmd.ExecuteNonQuery()
                End Using
            Else
                Using conn As SQLiteConnection = _dbManager.GetConnection()
                    Using cmd As New SQLiteCommand(sql, conn)
                        cmd.Parameters.AddWithValue("@val", If(hasAttachments, 1, 0))
                        cmd.Parameters.AddWithValue("@id", emailId)
                        cmd.ExecuteNonQuery()
                    End Using
                End Using
            End If
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  FTS トリガー制御（取り込み高速化）
        ' ════════════════════════════════════════════════════════════

        ''' <summary>FTS 同期トリガーを一時的に無効化する（取り込み中の INSERT を高速化）。</summary>
        Public Sub DisableFtsTriggers(conn As SQLiteConnection)
            Dim sql As String = "
DROP TRIGGER IF EXISTS emails_ai;
DROP TRIGGER IF EXISTS emails_ad;
DROP TRIGGER IF EXISTS emails_au;
DROP TRIGGER IF EXISTS attachments_ai;
DROP TRIGGER IF EXISTS attachments_ad;"
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        ''' <summary>FTS 同期トリガーを再作成する。RebuildFtsIndex の後に呼ぶこと。</summary>
        Public Sub EnableFtsTriggers(conn As SQLiteConnection)
            Dim sqlEmailsAi As String = "
CREATE TRIGGER IF NOT EXISTS emails_ai AFTER INSERT ON emails BEGIN
    INSERT INTO emails_fts(rowid, subject, body_text, sender_name, sender_email)
    VALUES (new.id, new.subject, new.body_text, new.sender_name, new.sender_email);
END;"
            Dim sqlEmailsAd As String = "
CREATE TRIGGER IF NOT EXISTS emails_ad AFTER DELETE ON emails BEGIN
    DELETE FROM emails_fts WHERE rowid = old.id;
END;"
            Dim sqlEmailsAu As String = "
CREATE TRIGGER IF NOT EXISTS emails_au AFTER UPDATE ON emails BEGIN
    DELETE FROM emails_fts WHERE rowid = old.id;
    INSERT INTO emails_fts(rowid, subject, body_text, sender_name, sender_email)
    VALUES (new.id, new.subject, new.body_text, new.sender_name, new.sender_email);
END;"
            Dim sqlAttachAi As String = "
CREATE TRIGGER IF NOT EXISTS attachments_ai AFTER INSERT ON attachments BEGIN
    INSERT INTO attachments_fts(rowid, file_name)
    VALUES (new.id, new.file_name);
END;"
            Dim sqlAttachAd As String = "
CREATE TRIGGER IF NOT EXISTS attachments_ad AFTER DELETE ON attachments BEGIN
    DELETE FROM attachments_fts WHERE rowid = old.id;
END;"
            Using cmd As New SQLiteCommand(sqlEmailsAi, conn)
                cmd.ExecuteNonQuery()
            End Using
            Using cmd As New SQLiteCommand(sqlEmailsAd, conn)
                cmd.ExecuteNonQuery()
            End Using
            Using cmd As New SQLiteCommand(sqlEmailsAu, conn)
                cmd.ExecuteNonQuery()
            End Using
            Using cmd As New SQLiteCommand(sqlAttachAi, conn)
                cmd.ExecuteNonQuery()
            End Using
            Using cmd As New SQLiteCommand(sqlAttachAd, conn)
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        ''' <summary>FTS インデックスをメインテーブルから一括再構築する。</summary>
        Public Sub RebuildFtsIndex(conn As SQLiteConnection)
            ' emails_fts を再構築
            Using cmd As New SQLiteCommand("INSERT INTO emails_fts(emails_fts) VALUES('rebuild');", conn)
                cmd.ExecuteNonQuery()
            End Using
            ' attachments_fts を再構築
            Using cmd As New SQLiteCommand("INSERT INTO attachments_fts(attachments_fts) VALUES('rebuild');", conn)
                cmd.ExecuteNonQuery()
            End Using
        End Sub

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

        ''' <summary>指定メール ID の添付ファイル一覧を返す。FilePath は絶対パスに変換済み。</summary>
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

            ' DB には相対パスが格納されているため絶対パスに変換する
            Dim baseDir As String = IO.Path.GetFullPath(Config.AppSettings.Instance.AttachmentDirectory)
            For Each att As Models.Attachment In result
                If Not IO.Path.IsPathRooted(att.FilePath) Then
                    att.FilePath = IO.Path.Combine(baseDir, att.FilePath)
                End If
            Next

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

        ''' <summary>
        ''' メールを全文検索する（件名・本文・差出人・添付ファイル名）。結果は受信日時降順。
        ''' ASCII のみのクエリは FTS4 インデックスを使用し、日本語等の非 ASCII 文字を含む場合は
        ''' LIKE による部分一致検索にフォールバックする（FTS4 の simple トークナイザーは
        ''' 日本語の単語境界を認識しないため）。
        ''' </summary>
        Public Function SearchEmails(query As String,
                                     Optional folderName As String = Nothing) As List(Of Models.Email)
            If ContainsNonAscii(query) Then
                Return SearchEmailsLike(query, folderName)
            End If

            ' ASCII のみ: FTS4 全文検索（添付ファイル名も含む）
            Dim sb As New StringBuilder(
                "SELECT DISTINCT e.* FROM emails e WHERE e.id IN (" &
                "SELECT rowid FROM emails_fts WHERE emails_fts MATCH @query1" &
                " UNION" &
                " SELECT a.email_id FROM attachments a" &
                " INNER JOIN attachments_fts af ON a.id = af.rowid" &
                " WHERE attachments_fts MATCH @query2)")
            If Not String.IsNullOrEmpty(folderName) Then
                sb.Append(" AND e.folder_name = @folder_name")
            End If
            sb.Append(" ORDER BY e.received_at DESC")

            Dim result As New List(Of Models.Email)
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sb.ToString(), conn)
                    cmd.Parameters.AddWithValue("@query1", query)
                    cmd.Parameters.AddWithValue("@query2", query)
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

        ''' <summary>
        ''' LIKE による部分一致検索（日本語等の非 ASCII クエリ用）。
        ''' FTS4 の simple トークナイザーは日本語の単語境界を扱えないため、
        ''' 非 ASCII 文字を含むクエリはこちらで処理する。
        ''' </summary>
        Private Function SearchEmailsLike(query As String,
                                          folderName As String) As List(Of Models.Email)
            Dim likeParam As String = "%" & query & "%"
            Dim sb As New StringBuilder(
                "SELECT e.* FROM emails e WHERE " &
                "(e.subject LIKE @q1 OR e.body_text LIKE @q2 OR e.sender_name LIKE @q3 OR e.sender_email LIKE @q4" &
                " OR e.id IN (SELECT a.email_id FROM attachments a WHERE a.file_name LIKE @q5))")
            If Not String.IsNullOrEmpty(folderName) Then
                sb.Append(" AND e.folder_name = @folder_name")
            End If
            sb.Append(" ORDER BY e.received_at DESC")

            Dim result As New List(Of Models.Email)
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sb.ToString(), conn)
                    cmd.Parameters.AddWithValue("@q1", likeParam)
                    cmd.Parameters.AddWithValue("@q2", likeParam)
                    cmd.Parameters.AddWithValue("@q3", likeParam)
                    cmd.Parameters.AddWithValue("@q4", likeParam)
                    cmd.Parameters.AddWithValue("@q5", likeParam)
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

        ''' <summary>文字列に非 ASCII 文字（日本語等）が含まれるか確認する。</summary>
        Private Function ContainsNonAscii(s As String) As Boolean
            For Each c As Char In s
                If AscW(c) > 127 Then Return True
            Next
            Return False
        End Function

        ' ════════════════════════════════════════════════════════════
        '  スレッド判定キャッシュ構築
        ' ════════════════════════════════════════════════════════════

        ''' <summary>
        ''' スレッド判定用インメモリキャッシュを構築する。
        ''' 取り込み開始前に呼び出し、ThreadingService.LoadCaches() に渡す。
        ''' </summary>
        Public Sub GetThreadIdCaches(ByRef messageIdMap As Dictionary(Of String, String),
                                     ByRef subjectMap As Dictionary(Of String, String))
            messageIdMap = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            subjectMap = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            Const sql As String =
                "SELECT message_id, normalized_subject, thread_id FROM emails WHERE thread_id IS NOT NULL"

            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim msgId As String = GetStr(reader, "message_id")
                            Dim subject As String = GetStr(reader, "normalized_subject")
                            Dim threadId As String = GetStr(reader, "thread_id")
                            If Not String.IsNullOrEmpty(msgId) AndAlso Not messageIdMap.ContainsKey(msgId) Then
                                messageIdMap.Add(msgId, threadId)
                            End If
                            If Not String.IsNullOrEmpty(subject) AndAlso Not subjectMap.ContainsKey(subject) Then
                                subjectMap.Add(subject, threadId)
                            End If
                        End While
                    End Using
                End Using
            End Using
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  削除・重複チェック
        ' ════════════════════════════════════════════════════════════

        ''' <summary>全メールの MessageID を一括取得する（取り込み時の高速重複チェック用）。</summary>
        Public Function GetAllMessageIds() As HashSet(Of String)
            Const sql As String = "SELECT message_id FROM emails WHERE message_id IS NOT NULL"
            Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
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

        ''' <summary>全削除済み MessageID を一括取得する（取り込み時の高速重複チェック用）。</summary>
        Public Function GetAllDeletedMessageIds() As HashSet(Of String)
            Const sql As String = "SELECT message_id FROM deleted_message_ids"
            Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
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
            Dim sizeOrd As Integer = reader.GetOrdinal("email_size")
            If Not reader.IsDBNull(sizeOrd) Then email.EmailSize = reader.GetInt64(sizeOrd)
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
            att.ContentId = GetStr(reader, "content_id")
            Dim isInlineOrd As Integer = reader.GetOrdinal("is_inline")
            If Not reader.IsDBNull(isInlineOrd) Then att.IsInline = reader.GetInt32(isInlineOrd) = 1
            Return att
        End Function

        ''' <summary>NULL の列を Nothing として返すヘルパー。</summary>
        Private Function GetStr(reader As SQLiteDataReader, column As String) As String
            Dim ord As Integer = reader.GetOrdinal(column)
            If reader.IsDBNull(ord) Then Return Nothing
            Return reader.GetString(ord)
        End Function

        ''' <summary>Nothing の文字列を DBNull.Value に変換するヘルパー。</summary>
        Private Shared Function NullableStr(value As String) As Object
            If value Is Nothing Then Return CType(DBNull.Value, Object)
            Return CType(value, Object)
        End Function

    End Class

End Namespace
