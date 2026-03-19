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
        Implements IDisposable

        Private _disposed As Boolean = False
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
        '  Email 取得
        ' ════════════════════════════════════════════════════════════

        ''' <summary>一覧表示用に軽量なカラムのみ取得する（本文を除外）。</summary>
        Public Function GetEmailsForList(Optional folderName As String = Nothing) As List(Of Models.Email)
            Dim sb As New StringBuilder(
                "SELECT id, subject, sender_name, sender_email, received_at, has_attachments, email_size, thread_id FROM emails")
            If Not String.IsNullOrEmpty(folderName) Then
                sb.Append(" WHERE folder_name = @folder_name")
            End If
            sb.Append(" ORDER BY received_at DESC")

            Dim result As New List(Of Models.Email)
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sb.ToString(), conn)
                    If Not String.IsNullOrEmpty(folderName) Then
                        cmd.Parameters.AddWithValue("@folder_name", folderName)
                    End If
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            result.Add(MapEmailSummary(reader))
                        End While
                    End Using
                End Using
            End Using
            Return result
        End Function

        ''' <summary>フォルダ指定でメール一覧を取得する（受信日時降順）。</summary>
        ''' <param name="folderName">フォルダ名。Nothing の場合はすべて。</param>
        ''' <param name="pageIndex">ページ番号（0始まり）。pageSize &lt;= 0 の場合は無視。</param>
        ''' <param name="pageSize">ページサイズ。0 以下の場合は全件取得。</param>
        Public Function GetEmails(Optional folderName As String = Nothing,
                                  Optional pageIndex As Integer = 0,
                                  Optional pageSize As Integer = 0) As List(Of Models.Email)
            Dim sb As New StringBuilder("SELECT * FROM emails")
            If Not String.IsNullOrEmpty(folderName) Then
                sb.Append(" WHERE folder_name = @folder_name")
            End If
            sb.Append(" ORDER BY received_at DESC")
            If pageSize > 0 Then
                sb.AppendFormat(" LIMIT {0} OFFSET {1}", pageSize, pageIndex * pageSize)
            End If

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

        ''' <summary>
        ''' フォルダ別の件数と全体件数を1クエリで取得する。
        ''' Item1: 全体件数、Item2: フォルダ名→件数の Dictionary。
        ''' </summary>
        Public Function GetFolderCounts() As Tuple(Of Integer, Dictionary(Of String, Integer))
            Dim folders As New Dictionary(Of String, Integer)()
            Dim total As Integer = 0
            Const sql As String = "SELECT folder_name, COUNT(1) FROM emails WHERE folder_name IS NOT NULL GROUP BY folder_name ORDER BY folder_name"
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim folderName As String = reader.GetString(0)
                            Dim count As Integer = reader.GetInt32(1)
                            folders(folderName) = count
                            total += count
                        End While
                    End Using
                End Using
                ' folder_name が NULL のレコード数を加算
                Using cmd2 As New SQLiteCommand("SELECT COUNT(1) FROM emails WHERE folder_name IS NULL", conn)
                    total += Convert.ToInt32(cmd2.ExecuteScalar())
                End Using
            End Using
            Return Tuple.Create(total, folders)
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

        ''' <summary>最終取り込み日時（emails.created_at の最大値）を返す。レコードがない場合は Nothing。</summary>
        Public Function GetLastImportDate() As DateTime?
            Const sql As String = "SELECT MAX(created_at) FROM emails"
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    Dim result As Object = cmd.ExecuteScalar()
                    If result Is Nothing OrElse TypeOf result Is DBNull Then Return Nothing
                    Dim dateStr As String = CType(result, String)
                    Dim dt As DateTime
                    If DateTime.TryParse(dateStr, dt) Then Return dt
                    Return Nothing
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

        ''' <summary>
        ''' EmailSearchFilter を使用してメールを検索する。
        ''' 列指定・AND/OR・添付有無などの高度なフィルタ構文に対応。
        ''' </summary>
        Public Function SearchEmailsFiltered(query As String,
                                             Optional folderName As String = Nothing) As List(Of Models.Email)
            Dim filter As New Filters.EmailSearchFilter()
            Dim sq As Filters.EmailSearchFilter.SearchQuery = filter.Parse(query, folderName)

            ' サブクエリで ID を絞り、外側では一覧表示用の軽量カラムのみ取得
            Dim sb As New StringBuilder(
                "SELECT e2.id, e2.subject, e2.sender_name, e2.sender_email, e2.received_at, e2.has_attachments, e2.email_size, e2.thread_id" &
                " FROM emails e2 INNER JOIN (SELECT DISTINCT e.id FROM emails e")
            If Not String.IsNullOrEmpty(sq.WhereClause) Then
                sb.Append(" WHERE ")
                sb.Append(sq.WhereClause)
            End If
            sb.Append(") matched ON e2.id = matched.id ORDER BY e2.received_at DESC")

            Dim result As New List(Of Models.Email)()
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sb.ToString(), conn)
                    For Each p As SQLiteParameter In sq.Parameters
                        cmd.Parameters.Add(p)
                    Next
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            result.Add(MapEmailSummary(reader))
                        End While
                    End Using
                End Using
            End Using
            Return result
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

        ''' <summary>
        ''' 指定フォルダの取り込み済みメールの message_id と id の Dictionary を返す。
        ''' 削除同期で Outlook 側と突合するために使用する。
        ''' </summary>
        Public Function GetMessageIdsByFolder(folderName As String) As Dictionary(Of String, Integer)
            Const sql As String = "SELECT message_id, id FROM emails WHERE folder_name = @folder_name AND message_id IS NOT NULL"
            Dim result As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@folder_name", folderName)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim msgId As String = reader.GetString(0)
                            Dim id As Integer = reader.GetInt32(1)
                            If Not result.ContainsKey(msgId) Then
                                result.Add(msgId, id)
                            End If
                        End While
                    End Using
                End Using
            End Using
            Return result
        End Function

        ''' <summary>
        ''' 複数メールを一括削除し、削除対象の添付ファイルパス一覧を返す。
        ''' トゥームストーンへの登録も行う。
        ''' ON DELETE CASCADE で attachments レコードは自動削除される。
        ''' </summary>
        Public Function DeleteEmailsByIds(ids As List(Of Integer)) As List(Of String)
            Dim attachmentPaths As New List(Of String)()
            If ids Is Nothing OrElse ids.Count = 0 Then Return attachmentPaths

            Dim baseDir As String = IO.Path.GetFullPath(Config.AppSettings.Instance.AttachmentDirectory)

            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using tx As SQLiteTransaction = conn.BeginTransaction()
                    ' 添付ファイルパスを収集
                    For Each emailId As Integer In ids
                        Using cmd As New SQLiteCommand("SELECT file_path FROM attachments WHERE email_id = @id", conn)
                            cmd.Parameters.AddWithValue("@id", CType(emailId, Object))
                            Using reader As SQLiteDataReader = cmd.ExecuteReader()
                                While reader.Read()
                                    Dim relativePath As String = reader.GetString(0)
                                    Dim fullPath As String = IO.Path.Combine(baseDir, relativePath)
                                    attachmentPaths.Add(fullPath)
                                End While
                            End Using
                        End Using
                    Next

                    ' message_id をトゥームストーンに登録
                    Using cmd As New SQLiteCommand(
                        "INSERT OR IGNORE INTO deleted_message_ids (message_id) " &
                        "SELECT message_id FROM emails WHERE id = @id AND message_id IS NOT NULL", conn)
                        cmd.Parameters.Add("@id", System.Data.DbType.Int32)
                        For Each emailId As Integer In ids
                            cmd.Parameters("@id").Value = CType(emailId, Object)
                            cmd.ExecuteNonQuery()
                        Next
                    End Using

                    ' メールを一括削除（CASCADE で attachments も削除）
                    Using cmd As New SQLiteCommand("DELETE FROM emails WHERE id = @id", conn)
                        cmd.Parameters.Add("@id", System.Data.DbType.Int32)
                        For Each emailId As Integer In ids
                            cmd.Parameters("@id").Value = CType(emailId, Object)
                            cmd.ExecuteNonQuery()
                        Next
                    End Using

                    tx.Commit()
                End Using
            End Using

            Return attachmentPaths
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
        '  エラー除外 MessageID
        ' ════════════════════════════════════════════════════════════

        ''' <summary>全エラー除外 MessageID を一括取得する（取り込み時の高速スキップ用）。</summary>
        Public Function GetAllErrorMessageIds() As HashSet(Of String)
            Const sql As String = "SELECT message_id FROM error_message_ids"
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

        ''' <summary>エラー MessageID を登録する（次回以降の取り込みでスキップされる）。</summary>
        Public Sub InsertErrorMessageId(messageId As String,
                                         folderName As String,
                                         subject As String,
                                         errorMessage As String,
                                         Optional receivedDate As DateTime? = Nothing,
                                         Optional senderName As String = "")
            If String.IsNullOrEmpty(messageId) Then Return
            Const sql As String =
                "INSERT OR REPLACE INTO error_message_ids (message_id, folder_name, subject, error_message, received_date, sender_name) " &
                "VALUES (@message_id, @folder_name, @subject, @error_message, @received_date, @sender_name)"
            If _bulkConn IsNot Nothing Then
                Using cmd As New SQLiteCommand(sql, _bulkConn)
                    cmd.Transaction = _bulkTx
                    SetErrorMessageIdParams(cmd, messageId, folderName, subject, errorMessage, receivedDate, senderName)
                    cmd.ExecuteNonQuery()
                End Using
            Else
                Using conn As SQLiteConnection = _dbManager.GetConnection()
                    Using cmd As New SQLiteCommand(sql, conn)
                        SetErrorMessageIdParams(cmd, messageId, folderName, subject, errorMessage, receivedDate, senderName)
                        cmd.ExecuteNonQuery()
                    End Using
                End Using
            End If
        End Sub

        ''' <summary>InsertErrorMessageId のパラメータセット用ヘルパー。</summary>
        Private Shared Sub SetErrorMessageIdParams(cmd As SQLiteCommand,
                                                    messageId As String,
                                                    folderName As String,
                                                    subject As String,
                                                    errorMessage As String,
                                                    receivedDate As DateTime?,
                                                    senderName As String)
            cmd.Parameters.AddWithValue("@message_id", messageId)
            cmd.Parameters.AddWithValue("@folder_name", NullableStr(folderName))
            cmd.Parameters.AddWithValue("@subject", NullableStr(subject))
            cmd.Parameters.AddWithValue("@error_message", NullableStr(errorMessage))
            cmd.Parameters.AddWithValue("@received_date",
                If(receivedDate.HasValue,
                   CType(receivedDate.Value.ToString("o"), Object),
                   CType(DBNull.Value, Object)))
            cmd.Parameters.AddWithValue("@sender_name", NullableStr(senderName))
        End Sub

        ''' <summary>指定 MessageID のエラー除外を解除する。</summary>
        Public Sub DeleteErrorMessageId(messageId As String)
            Const sql As String = "DELETE FROM error_message_ids WHERE message_id = @message_id"
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@message_id", messageId)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        ''' <summary>全エラー除外を解除する。</summary>
        Public Sub ClearAllErrorMessageIds()
            Const sql As String = "DELETE FROM error_message_ids"
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        ''' <summary>エラー除外リストの全件を取得する（UI 表示用）。</summary>
        Public Function GetErrorMessageEntries() As System.Data.DataTable
            Const sql As String =
                "SELECT message_id, folder_name, subject, error_message, received_date, sender_name, error_date " &
                "FROM error_message_ids ORDER BY error_date DESC"
            Dim dt As New System.Data.DataTable()
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using
            End Using
            Return dt
        End Function

        ''' <summary>エラー除外リストの件数を返す。</summary>
        Public Function GetErrorMessageIdCount() As Integer
            Const sql As String = "SELECT COUNT(1) FROM error_message_ids"
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    Return Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            End Using
        End Function

        ' ════════════════════════════════════════════════════════════
        '  Exchange アドレスキャッシュ
        ' ════════════════════════════════════════════════════════════

        ''' <summary>永続化された Exchange アドレスキャッシュを全件読み込む。</summary>
        Public Function LoadExchangeAddressCache() As Dictionary(Of String, String)
            Const sql As String = "SELECT ex_address, smtp_address FROM exchange_address_cache"
            Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim exAddr As String = reader.GetString(0)
                            Dim smtpAddr As String = reader.GetString(1)
                            If Not result.ContainsKey(exAddr) Then
                                result.Add(exAddr, smtpAddr)
                            End If
                        End While
                    End Using
                End Using
            End Using
            Return result
        End Function

        ''' <summary>Exchange アドレスキャッシュの新規分を DB に書き戻す。</summary>
        Public Sub SaveExchangeAddressCache(cache As Dictionary(Of String, String))
            Const sql As String = "INSERT OR IGNORE INTO exchange_address_cache (ex_address, smtp_address) VALUES (@ex, @smtp)"
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using tx As SQLiteTransaction = conn.BeginTransaction()
                    Using cmd As New SQLiteCommand(sql, conn)
                        cmd.Parameters.Add("@ex", System.Data.DbType.String)
                        cmd.Parameters.Add("@smtp", System.Data.DbType.String)
                        For Each kvp As KeyValuePair(Of String, String) In cache
                            cmd.Parameters("@ex").Value = kvp.Key
                            cmd.Parameters("@smtp").Value = kvp.Value
                            cmd.ExecuteNonQuery()
                        Next
                    End Using
                    tx.Commit()
                End Using
            End Using
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  マッピング・ユーティリティ
        ' ════════════════════════════════════════════════════════════

        ''' <summary>一覧表示用の軽量マッピング（本文系カラムなし）。</summary>
        Private Function MapEmailSummary(reader As SQLiteDataReader) As Models.Email
            Dim email As New Models.Email()
            email.Id = reader.GetInt32(reader.GetOrdinal("id"))
            email.Subject = GetStr(reader, "subject")
            email.SenderName = GetStr(reader, "sender_name")
            email.SenderEmail = GetStr(reader, "sender_email")
            email.ReceivedAt = DateTime.Parse(reader.GetString(reader.GetOrdinal("received_at")))
            email.HasAttachments = reader.GetInt32(reader.GetOrdinal("has_attachments")) = 1
            Dim sizeOrd As Integer = reader.GetOrdinal("email_size")
            If Not reader.IsDBNull(sizeOrd) Then email.EmailSize = reader.GetInt64(sizeOrd)
            email.ThreadId = GetStr(reader, "thread_id")
            Return email
        End Function

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

        ' ════════════════════════════════════════════════════════════
        '  フォルダ同期状態
        ' ════════════════════════════════════════════════════════════

        ''' <summary>指定フォルダの同期状態を取得する。未登録の場合は Nothing を返す。</summary>
        Public Function GetFolderSyncState(folderName As String) As Models.FolderSyncState
            Const sql As String =
                "SELECT folder_name, last_sync_time, full_sync_done FROM folder_sync_state WHERE folder_name = @folder_name"
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@folder_name", folderName)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Dim state As New Models.FolderSyncState()
                            state.FolderName = reader.GetString(0)
                            state.LastSyncTime = DateTime.Parse(reader.GetString(1))
                            state.FullSyncDone = reader.GetInt32(2) = 1
                            Return state
                        End If
                    End Using
                End Using
            End Using
            Return Nothing
        End Function

        ''' <summary>フォルダの同期状態を更新する（UPSERT）。</summary>
        Public Sub UpdateFolderSyncState(folderName As String, lastSyncTime As DateTime, fullSyncDone As Boolean)
            Const sql As String =
                "INSERT INTO folder_sync_state (folder_name, last_sync_time, full_sync_done, updated_at) " &
                "VALUES (@folder_name, @last_sync_time, @full_sync_done, datetime('now', 'localtime')) " &
                "ON CONFLICT(folder_name) DO UPDATE SET " &
                "last_sync_time = @last_sync_time, full_sync_done = @full_sync_done, updated_at = datetime('now', 'localtime')"
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@folder_name", folderName)
                    cmd.Parameters.AddWithValue("@last_sync_time", lastSyncTime.ToString("yyyy-MM-dd HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@full_sync_done", If(fullSyncDone, 1, 0))
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  IDisposable
        ' ════════════════════════════════════════════════════════════

        Protected Overridable Sub Dispose(disposing As Boolean)
            If _disposed Then Return
            If disposing Then
                ' バルクモードが残っていればロールバックしてリソース解放
                RollbackBulk()
            End If
            _disposed = True
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

    End Class

End Namespace
