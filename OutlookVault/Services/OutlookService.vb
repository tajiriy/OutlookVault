Option Explicit On
Option Strict On
Option Infer Off

Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text

Namespace Services

    ''' <summary>
    ''' Outlook COM API のラッパー。フォルダ一覧の取得・メールデータの抽出・添付ファイルの保存を担当する。
    ''' </summary>
    Public Class OutlookService
        Implements IDisposable

        ' ── MAPI プロパティタグ ────────────────────────────────────
        Private Const PropMessageId As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
        Private Const PropInReplyTo As String = "http://schemas.microsoft.com/mapi/proptag/0x1042001F"
        Private Const PropTransportHeaders As String = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
        ''' <summary>PR_ATTACH_CONTENT_ID: 添付ファイルの MIME Content-ID（インライン画像の cid: 参照に対応）</summary>
        Private Const PropAttachContentId As String = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
        ''' <summary>PR_ATTR_HIDDEN: フォルダの隠し属性（システムフォルダの判定に使用）</summary>
        Private Const PropAttrHidden As String = "http://schemas.microsoft.com/mapi/proptag/0x10F4000B"

        Private ReadOnly _app As Outlook.Application
        Private ReadOnly _ns As Outlook.NameSpace
        Private _disposed As Boolean

        ''' <summary>Exchange アドレス (EX) → SMTP アドレスのキャッシュ。GetExchangeUser() のネットワーク往復を削減する。</summary>
        Private ReadOnly _exchangeSmtpCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        Private Sub New(app As Outlook.Application)
            _app = app
            _ns = _app.GetNamespace("MAPI")
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  接続
        ' ════════════════════════════════════════════════════════════

        ''' <summary>起動済みの Outlook に接続する。起動していない場合は Nothing を返す。</summary>
        Public Shared Function TryConnect() As OutlookService
            Try
                Dim obj As Object = Marshal.GetActiveObject("Outlook.Application")
                Dim app As Outlook.Application = CType(obj, Outlook.Application)
                Return New OutlookService(app)
            Catch ex As COMException
                Return Nothing
            End Try
        End Function

        ''' <summary>起動済みの Outlook に接続する。起動していない場合は新規起動する。</summary>
        Public Shared Function Connect() As OutlookService
            Dim svc As OutlookService = TryConnect()
            If svc IsNot Nothing Then Return svc
            Dim app As New Outlook.Application()
            Return New OutlookService(app)
        End Function

        ' ════════════════════════════════════════════════════════════
        '  Exchange アドレスキャッシュ管理
        ' ════════════════════════════════════════════════════════════

        ''' <summary>永続化済みの Exchange アドレスキャッシュをメモリにロードする。取り込み開始前に呼ぶ。</summary>
        Public Sub LoadExchangeCache(cache As Dictionary(Of String, String))
            For Each kvp As KeyValuePair(Of String, String) In cache
                If Not _exchangeSmtpCache.ContainsKey(kvp.Key) Then
                    _exchangeSmtpCache.Add(kvp.Key, kvp.Value)
                End If
            Next
        End Sub

        ''' <summary>現在のメモリ上の Exchange アドレスキャッシュを返す。取り込み終了後の永続化に使用。</summary>
        Public Function GetExchangeCache() As Dictionary(Of String, String)
            Return _exchangeSmtpCache
        End Function

        ' ════════════════════════════════════════════════════════════
        '  フォルダ一覧
        ' ════════════════════════════════════════════════════════════

        ''' <summary>全ストアを検索してメールフォルダの表示名一覧を返す。</summary>
        Public Function GetAvailableFolderNames() As List(Of String)
            ' 除外対象のシステムフォルダの EntryID を収集
            Dim excludedIds As HashSet(Of String) = GetExcludedFolderEntryIds()

            Dim result As New List(Of String)()
            Dim stores As Outlook.Stores = _ns.Stores
            Try
                For i As Integer = 1 To stores.Count
                    Dim store As Outlook.Store = stores.Item(i)
                    Dim root As Outlook.MAPIFolder = Nothing
                    Try
                        root = store.GetRootFolder()
                        ' ルートフォルダ（ストア名）は除外し、サブフォルダのみ収集
                        Dim subFolders As Outlook.Folders = root.Folders
                        Try
                            For j As Integer = 1 To subFolders.Count
                                Dim childFolder As Outlook.MAPIFolder = CType(subFolders.Item(j), Outlook.MAPIFolder)
                                Try
                                    CollectFolderNames(childFolder, result, excludedIds)
                                Finally
                                    Marshal.ReleaseComObject(childFolder)
                                End Try
                            Next
                        Finally
                            Marshal.ReleaseComObject(subFolders)
                        End Try
                    Finally
                        If root IsNot Nothing Then Marshal.ReleaseComObject(root)
                        Marshal.ReleaseComObject(store)
                    End Try
                Next
            Finally
                Marshal.ReleaseComObject(stores)
            End Try
            Return result
        End Function

        ''' <summary>除外対象のシステムフォルダ（同期の失敗、RSS フィード等）の EntryID セットを返す。</summary>
        Private Function GetExcludedFolderEntryIds() As HashSet(Of String)
            Dim ids As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            ' 除外するデフォルトフォルダの種別
            Dim excludedTypes As Outlook.OlDefaultFolders() = New Outlook.OlDefaultFolders() {
                Outlook.OlDefaultFolders.olFolderSyncIssues,
                Outlook.OlDefaultFolders.olFolderConflicts,
                Outlook.OlDefaultFolders.olFolderLocalFailures,
                Outlook.OlDefaultFolders.olFolderServerFailures,
                Outlook.OlDefaultFolders.olFolderRssFeeds
            }
            For Each folderType As Outlook.OlDefaultFolders In excludedTypes
                Try
                    Dim f As Outlook.MAPIFolder = _ns.GetDefaultFolder(folderType)
                    If f IsNot Nothing Then
                        Try
                            ids.Add(f.EntryID)
                        Finally
                            Marshal.ReleaseComObject(f)
                        End Try
                    End If
                Catch ex As COMException
                    ' デフォルトフォルダが存在しない環境では例外が発生する（正常動作）
                End Try
            Next
            Return ids
        End Function

        Private Sub CollectFolderNames(folder As Outlook.MAPIFolder, result As List(Of String),
                                       excludedIds As HashSet(Of String))
            ' 除外対象フォルダとその配下はスキップ
            If excludedIds.Contains(folder.EntryID) Then Return

            ' メールフォルダのみ対象：ContainerClass が "IPF.Note" かつ隠しフォルダでないこと
            If GetContainerClass(folder) = "IPF.Note" AndAlso
               Not IsFolderHidden(folder) Then
                result.Add(folder.Name)
            End If
            Dim subFolders As Outlook.Folders = folder.Folders
            Try
                For i As Integer = 1 To subFolders.Count
                    Dim childFolder As Outlook.MAPIFolder = CType(subFolders.Item(i), Outlook.MAPIFolder)
                    Try
                        CollectFolderNames(childFolder, result, excludedIds)
                    Finally
                        Marshal.ReleaseComObject(childFolder)
                    End Try
                Next
            Finally
                Marshal.ReleaseComObject(subFolders)
            End Try
        End Sub

        ''' <summary>フォルダが隠し属性（PR_ATTR_HIDDEN）を持つかどうかを返す。</summary>
        Private Shared Function IsFolderHidden(folder As Outlook.MAPIFolder) As Boolean
            Try
                Dim pa As Outlook.PropertyAccessor = folder.PropertyAccessor
                Try
                    Dim val As Object = pa.GetProperty(PropAttrHidden)
                    If TypeOf val Is Boolean Then Return CBool(val)
                Finally
                    Marshal.ReleaseComObject(pa)
                End Try
            Catch ex As COMException
                ' PropertyAccessor が未サポートのフォルダでは例外が発生する（正常動作）
            End Try
            Return False
        End Function

        ''' <summary>フォルダの PR_CONTAINER_CLASS を返す。取得できない場合は空文字列。</summary>
        Private Shared Function GetContainerClass(folder As Outlook.MAPIFolder) As String
            Try
                Dim pa As Outlook.PropertyAccessor = folder.PropertyAccessor
                Try
                    Dim val As Object = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3613001F")
                    If val IsNot Nothing AndAlso TypeOf val Is String Then Return CStr(val)
                Finally
                    Marshal.ReleaseComObject(pa)
                End Try
            Catch ex As COMException
                ' ContainerClass が未設定のフォルダでは例外が発生する（正常動作）
            End Try
            Return String.Empty
        End Function

        ''' <summary>フォルダ名でフォルダを検索して返す。見つからない場合は Nothing。</summary>
        Public Function FindFolder(folderName As String) As Outlook.MAPIFolder
            Dim result As Outlook.MAPIFolder = Nothing
            Dim stores As Outlook.Stores = _ns.Stores
            Try
                For i As Integer = 1 To stores.Count
                    Dim store As Outlook.Store = stores.Item(i)
                    Dim root As Outlook.MAPIFolder = Nothing
                    Try
                        root = store.GetRootFolder()
                        Dim found As Outlook.MAPIFolder = SearchFolder(root, folderName)
                        If found IsNot Nothing Then
                            result = found
                            ' found が root 自身の場合は root を解放しない（呼び出し元で使用）
                            If found IsNot root Then
                                Marshal.ReleaseComObject(root)
                                root = Nothing  ' Finally での二重解放を防止
                            Else
                                root = Nothing  ' found = root なので解放しない
                            End If
                            Exit For
                        End If
                    Finally
                        If root IsNot Nothing Then Marshal.ReleaseComObject(root)
                        Marshal.ReleaseComObject(store)
                    End Try
                Next
            Finally
                Marshal.ReleaseComObject(stores)
            End Try
            Return result
        End Function

        Private Function SearchFolder(folder As Outlook.MAPIFolder, name As String) As Outlook.MAPIFolder
            If folder.Name = name Then Return folder
            Dim subFolders As Outlook.Folders = folder.Folders
            Try
                For i As Integer = 1 To subFolders.Count
                    Dim childFolder As Outlook.MAPIFolder = CType(subFolders.Item(i), Outlook.MAPIFolder)
                    Dim found As Outlook.MAPIFolder = SearchFolder(childFolder, name)
                    If found IsNot Nothing Then
                        ' found を返すので childFolder は解放しない（found 自体の可能性がある）
                        If found IsNot childFolder Then Marshal.ReleaseComObject(childFolder)
                        Return found
                    End If
                    Marshal.ReleaseComObject(childFolder)
                Next
            Finally
                Marshal.ReleaseComObject(subFolders)
            End Try
            Return Nothing
        End Function

        ''' <summary>フォルダ内のアイテム数を返す（MailItem のみカウント）。</summary>
        Public Function GetMailItemCount(folderName As String) As Integer
            Dim folder As Outlook.MAPIFolder = FindFolder(folderName)
            If folder Is Nothing Then Return 0
            Dim items As Outlook.Items = folder.Items
            Try
                Return items.Count
            Finally
                Marshal.ReleaseComObject(items)
                Marshal.ReleaseComObject(folder)
            End Try
        End Function

        ' ════════════════════════════════════════════════════════════
        '  フォルダ内 MessageID 一括取得（削除同期用）
        ' ════════════════════════════════════════════════════════════

        ''' <summary>
        ''' フォルダ内の全 MailItem から MessageID のみを軽量に取得して HashSet で返す。
        ''' Outlook 側で削除されたメールを検出するための突合に使用する。
        ''' </summary>
        Public Function GetFolderMessageIds(folder As Outlook.MAPIFolder,
                                            Optional progress As IProgress(Of Integer) = Nothing) As HashSet(Of String)
            Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim items As Outlook.Items = folder.Items
            Try
                Dim totalCount As Integer = items.Count

                For i As Integer = 1 To totalCount
                    Dim rawItem As Object = items.Item(i)
                    Try
                        If TypeOf rawItem Is Outlook.MailItem Then
                            Dim mailItem As Outlook.MailItem = CType(rawItem, Outlook.MailItem)
                            Try
                                Dim messageId As String = ExtractMessageId(mailItem)
                                If Not String.IsNullOrEmpty(messageId) Then
                                    result.Add(messageId)
                                End If
                            Catch
                                ' 個別メールの MessageID 取得エラーはスキップ
                            End Try
                        End If
                    Finally
                        Marshal.ReleaseComObject(rawItem)
                    End Try
                    If progress IsNot Nothing AndAlso i Mod 100 = 0 Then
                        progress.Report(i)
                    End If
                Next
            Finally
                Marshal.ReleaseComObject(items)
            End Try

            Return result
        End Function

        ' ════════════════════════════════════════════════════════════
        '  メールデータ抽出
        ' ════════════════════════════════════════════════════════════

        ''' <summary>
        ''' MailItem から MessageID のみを軽量に取得する。
        ''' MAPI PropertyAccessor → トランスポートヘッダー → EntryID の順にフォールバック。
        ''' 本文・受信者・添付ファイルなどの重い COM 操作は行わない。
        ''' </summary>
        Public Function ExtractMessageId(mailItem As Outlook.MailItem) As String
            Dim pa As Outlook.PropertyAccessor = mailItem.PropertyAccessor
            Try
                Dim messageId As String = CleanMessageId(GetMAPIString(pa, PropMessageId))
                If Not String.IsNullOrEmpty(messageId) Then Return messageId

                ' MAPI で取れなかった場合はトランスポートヘッダーから補完
                Dim transportHeaders As String = GetMAPIString(pa, PropTransportHeaders)
                If Not String.IsNullOrEmpty(transportHeaders) Then
                    messageId = CleanMessageId(ParseHeaderField(transportHeaders, "Message-ID"))
                    If Not String.IsNullOrEmpty(messageId) Then Return messageId
                End If

                ' MessageId が取れない場合は EntryID を代替として使用
                Return "entryid:" & mailItem.EntryID
            Finally
                Marshal.ReleaseComObject(pa)
            End Try
        End Function

        ''' <summary>
        ''' MailItem から Models.Email を生成して返す。ThreadId は設定しない。
        ''' </summary>
        Public Function ExtractEmailData(mailItem As Outlook.MailItem) As Models.Email
            Dim email As New Models.Email()

            ' ── MAPI プロパティ ──────────────────────────────────
            Dim pa As Outlook.PropertyAccessor = mailItem.PropertyAccessor
            Try
                Dim transportHeaders As String = GetMAPIString(pa, PropTransportHeaders)

                email.MessageId = CleanMessageId(GetMAPIString(pa, PropMessageId))
                email.InReplyTo = CleanMessageId(GetMAPIString(pa, PropInReplyTo))

                ' MAPI で取れなかった場合はトランスポートヘッダーから補完
                If Not String.IsNullOrEmpty(transportHeaders) Then
                    If String.IsNullOrEmpty(email.MessageId) Then
                        email.MessageId = CleanMessageId(ParseHeaderField(transportHeaders, "Message-ID"))
                    End If
                    If String.IsNullOrEmpty(email.InReplyTo) Then
                        email.InReplyTo = CleanMessageId(ParseHeaderField(transportHeaders, "In-Reply-To"))
                    End If
                    email.References = ParseHeaderField(transportHeaders, "References")
                End If

                ' MessageId が取れない場合は EntryID を代替として使用
                If String.IsNullOrEmpty(email.MessageId) Then
                    email.MessageId = "entryid:" & mailItem.EntryID
                End If
            Finally
                Marshal.ReleaseComObject(pa)
            End Try

            ' ── 件名 ────────────────────────────────────────────
            email.EntryId = mailItem.EntryID
            email.Subject = mailItem.Subject
            email.NormalizedSubject = ThreadingService.NormalizeSubject(email.Subject)

            ' ── 差出人 ──────────────────────────────────────────
            email.SenderName = mailItem.SenderName
            email.SenderEmail = GetSenderEmail(mailItem)

            ' ── 受信者 ──────────────────────────────────────────
            Dim recipients As Outlook.Recipients = mailItem.Recipients
            Try
                email.ToRecipients = SerializeRecipients(recipients, Outlook.OlMailRecipientType.olTo)
                email.CcRecipients = SerializeRecipients(recipients, Outlook.OlMailRecipientType.olCC)
                email.BccRecipients = SerializeRecipients(recipients, Outlook.OlMailRecipientType.olBCC)
            Finally
                Marshal.ReleaseComObject(recipients)
            End Try

            ' ── 本文 ────────────────────────────────────────────
            email.BodyText = mailItem.Body
            email.BodyHtml = mailItem.HTMLBody

            ' ── 日時 ────────────────────────────────────────────
            Dim receivedTime As DateTime = mailItem.ReceivedTime
            email.ReceivedAt = If(receivedTime > DateTime.MinValue, receivedTime, DateTime.Now)
            Dim sentOn As DateTime = mailItem.SentOn
            If sentOn > DateTime.MinValue Then email.SentAt = sentOn

            ' ── フォルダ名 ──────────────────────────────────────
            Dim parentObj As Object = mailItem.Parent
            Dim parentFolder As Outlook.MAPIFolder = TryCast(parentObj, Outlook.MAPIFolder)
            If parentFolder IsNot Nothing Then
                email.FolderName = parentFolder.Name
                Marshal.ReleaseComObject(parentFolder)
            Else
                email.FolderName = Nothing
                If parentObj IsNot Nothing Then Marshal.ReleaseComObject(parentObj)
            End If

            ' ── 添付ファイル（OLE 埋め込みを除いた実ファイル数でフラグを設定）──────
            ' ContentID の取得は SaveAttachments で行うため、ここでは OLE のみ除外する。
            ' インライン画像を含む場合も HasAttachments = True になるが、
            ' 添付パネルには IsInline=False のもののみ表示されるため実害はない。
            Dim mailAtts As Outlook.Attachments = mailItem.Attachments
            Try
                If mailAtts.Count = 0 Then
                    email.HasAttachments = False
                Else
                    Dim realAttachCount As Integer = 0
                    For attIdx As Integer = 1 To mailAtts.Count
                        Dim a As Outlook.Attachment = CType(mailAtts.Item(attIdx), Outlook.Attachment)
                        Try
                            If a.Type <> Outlook.OlAttachmentType.olOLE Then
                                realAttachCount += 1
                            End If
                        Finally
                            Marshal.ReleaseComObject(a)
                        End Try
                    Next
                    email.HasAttachments = realAttachCount > 0
                End If
            Finally
                Marshal.ReleaseComObject(mailAtts)
            End Try

            ' ── サイズ ──────────────────────────────────────────
            email.EmailSize = mailItem.Size

            Return email
        End Function

        ' ════════════════════════════════════════════════════════════
        '  添付ファイル保存
        ' ════════════════════════════════════════════════════════════

        ''' <summary>
        ''' MailItem の添付ファイルをファイルシステムに保存し、メタデータのリストを返す。
        ''' 保存先: {attachBaseDir}\{yyyyMMdd_HHmmss}_{shortId}\{filename}
        ''' FilePath には attachBaseDir からの相対パスを格納する。
        ''' </summary>
        Public Function SaveAttachments(mailItem As Outlook.MailItem,
                                        emailId As Integer,
                                        attachBaseDir As String,
                                        Optional saveErrors As List(Of String) = Nothing) As List(Of Models.Attachment)
            Dim result As New List(Of Models.Attachment)()
            Dim atts As Outlook.Attachments = mailItem.Attachments
            If atts.Count = 0 Then
                Marshal.ReleaseComObject(atts)
                Return result
            End If

            Dim dateStr As String = mailItem.ReceivedTime.ToString("yyyyMMdd_HHmmss")
            Dim shortId As String = GetShortId(mailItem.EntryID)
            Dim subDir As String = dateStr & "_" & shortId

            ' Outlook COM の SaveAsFile は相対パスを解釈できないため絶対パスに変換する
            Dim saveDir As String = Path.GetFullPath(Path.Combine(attachBaseDir, subDir))
            Dim dirCreated As Boolean = False

            Try
                For i As Integer = 1 To atts.Count
                    Dim att As Outlook.Attachment = CType(atts.Item(i), Outlook.Attachment)
                    Try
                        ' OLE オブジェクト（埋め込みオブジェクト）はスキップ
                        If att.Type = Outlook.OlAttachmentType.olOLE Then Continue For

                        Dim originalName As String = att.FileName
                        Dim sanitized As String = SanitizeFileName(originalName)
                        Dim contentId As String = GetAttachmentContentId(att)

                        Try
                            ' 保存先ディレクトリを初回ファイル保存直前に作成（空フォルダを残さない）
                            If Not dirCreated Then
                                If Not Directory.Exists(saveDir) Then
                                    Directory.CreateDirectory(saveDir)
                                End If
                                dirCreated = True
                            End If

                            Dim targetPath As String = GetUniqueFilePath(saveDir, sanitized)
                            att.SaveAsFile(targetPath)

                            Dim fi As New FileInfo(targetPath)
                            Dim model As New Models.Attachment()
                            model.EmailId = emailId
                            model.FileName = originalName
                            model.FilePath = Path.Combine(subDir, Path.GetFileName(targetPath))
                            model.FileSize = fi.Length
                            model.ContentId = contentId
                            model.IsInline = Not String.IsNullOrEmpty(contentId)
                            result.Add(model)
                        Catch ex As Exception
                            Dim errMsg As String = "添付保存エラー [" & originalName & "]: " & ex.Message
                            System.Diagnostics.Debug.WriteLine(errMsg)
                            If saveErrors IsNot Nothing Then saveErrors.Add(errMsg)
                        End Try
                    Finally
                        Marshal.ReleaseComObject(att)
                    End Try
                Next
            Finally
                Marshal.ReleaseComObject(atts)
            End Try

            Return result
        End Function

        ' ════════════════════════════════════════════════════════════
        '  MAPI / ヘッダーヘルパー
        ' ════════════════════════════════════════════════════════════

        Private Shared Function GetMAPIString(pa As Outlook.PropertyAccessor,
                                              propTag As String) As String
            Try
                Dim val As Object = pa.GetProperty(propTag)
                If val Is Nothing Then Return Nothing
                If Not TypeOf val Is String Then Return Nothing
                Return CType(val, String)
            Catch ex As COMException
                ' MAPI プロパティが存在しないメールでは例外が発生する（正常動作）
                Return Nothing
            End Try
        End Function

        ''' <summary>RFC 2822 トランスポートヘッダーから指定フィールドの値を取得する。</summary>
        Private Shared Function ParseHeaderField(headers As String,
                                                 fieldName As String) As String
            If String.IsNullOrEmpty(headers) Then Return Nothing

            Dim searchFor As String = fieldName.ToLower() & ":"
            Dim lines As String() = headers.Split(New String() {vbCrLf, vbLf},
                                                  StringSplitOptions.None)
            Dim inField As Boolean = False
            Dim sb As New StringBuilder()

            For Each line As String In lines
                If inField Then
                    ' RFC 2822 折り返し行（先頭がスペースまたはタブ）
                    If line.Length > 0 AndAlso (line(0) = " "c OrElse line(0) = Chr(9)) Then
                        sb.Append(" ").Append(line.Trim())
                    Else
                        Exit For
                    End If
                ElseIf line.ToLower().StartsWith(searchFor) Then
                    inField = True
                    sb.Append(line.Substring(searchFor.Length).Trim())
                End If
            Next

            Dim result As String = sb.ToString().Trim()
            Return If(String.IsNullOrEmpty(result), Nothing, result)
        End Function

        Private Shared Function CleanMessageId(messageId As String) As String
            If String.IsNullOrEmpty(messageId) Then Return messageId
            Dim result As String = messageId.Trim()
            If result.StartsWith("<") AndAlso result.EndsWith(">") Then
                result = result.Substring(1, result.Length - 2)
            End If
            Return result
        End Function

        ' ════════════════════════════════════════════════════════════
        '  差出人・受信者ヘルパー
        ' ════════════════════════════════════════════════════════════

        Private Function GetSenderEmail(mailItem As Outlook.MailItem) As String
            If mailItem.SenderEmailType <> "EX" Then
                Return mailItem.SenderEmailAddress
            End If
            Dim exAddr As String = mailItem.SenderEmailAddress
            Return ResolveExchangeAddress(exAddr, Function() mailItem.Sender)
        End Function

        ''' <summary>
        ''' Exchange アドレスを SMTP アドレスに解決する。キャッシュがあればネットワーク往復を回避する。
        ''' </summary>
        Private Function ResolveExchangeAddress(exAddress As String,
                                                 getAddressEntry As Func(Of Outlook.AddressEntry)) As String
            If String.IsNullOrEmpty(exAddress) Then Return exAddress

            ' キャッシュヒット
            Dim cached As String = Nothing
            If _exchangeSmtpCache.TryGetValue(exAddress, cached) Then Return cached

            ' キャッシュミス → Exchange Server に問い合わせ
            Dim resolved As String = exAddress
            Try
                Dim entry As Outlook.AddressEntry = getAddressEntry()
                If entry IsNot Nothing Then
                    Try
                        Dim exUser As Outlook.ExchangeUser = entry.GetExchangeUser()
                        If exUser IsNot Nothing Then
                            Try
                                resolved = exUser.PrimarySmtpAddress
                            Finally
                                Marshal.ReleaseComObject(exUser)
                            End Try
                        End If
                    Finally
                        Marshal.ReleaseComObject(entry)
                    End Try
                End If
            Catch ex As COMException
                ' Exchange Server への問い合わせ失敗は元のアドレスで代替する
                Logger.Warn("Exchange アドレス解決に失敗: " & exAddress & " - " & ex.Message)
            End Try

            _exchangeSmtpCache(exAddress) = resolved
            Return resolved
        End Function

        ''' <summary>指定タイプの受信者を JSON 配列文字列にシリアライズする。</summary>
        Private Function SerializeRecipients(recipients As Outlook.Recipients,
                                              recipType As Outlook.OlMailRecipientType) As String
            Dim expectedType As Integer = CType(recipType, Integer)
            Dim parts As New List(Of String)()

            For i As Integer = 1 To recipients.Count
                Dim r As Outlook.Recipient = recipients.Item(i)
                Try
                    If r.Type <> expectedType Then Continue For

                    Dim emailAddr As String = r.Address
                    ' Exchange アドレスの場合は SMTP アドレスを解決（キャッシュ付き）
                    Dim addrEntry As Outlook.AddressEntry = r.AddressEntry
                    If addrEntry IsNot Nothing Then
                        Try
                            If addrEntry.Type = "EX" Then
                                emailAddr = ResolveExchangeAddress(r.Address, Function() addrEntry)
                            End If
                        Finally
                            Marshal.ReleaseComObject(addrEntry)
                        End Try
                    End If

                    parts.Add("{""name"":" & JsonStr(r.Name) & ",""email"":" & JsonStr(emailAddr) & "}")
                Finally
                    Marshal.ReleaseComObject(r)
                End Try
            Next

            Return "[" & String.Join(",", parts) & "]"
        End Function

        Private Shared Function JsonStr(s As String) As String
            If s Is Nothing Then Return "null"
            Return """" & s.Replace("\", "\\").
                           Replace("""", "\""").
                           Replace(vbCr, "\r").
                           Replace(vbLf, "\n").
                           Replace(vbTab, "\t") & """"
        End Function

        ' ════════════════════════════════════════════════════════════
        '  添付ファイルヘルパー
        ' ════════════════════════════════════════════════════════════

        ''' <summary>
        ''' 添付ファイルの MIME Content-ID を取得する。
        ''' Content-ID がない（通常の添付ファイル）場合は Nothing を返す。
        ''' </summary>
        Private Shared Function GetAttachmentContentId(att As Outlook.Attachment) As String
            Try
                Dim pa As Outlook.PropertyAccessor = att.PropertyAccessor
                Try
                    Dim val As Object = pa.GetProperty(PropAttachContentId)
                    If val Is Nothing Then Return Nothing
                    If Not TypeOf val Is String Then Return Nothing
                    Dim cid As String = CType(val, String).Trim()
                    ' "<xxx@yyy>" 形式の山括弧を除去して正規化
                    If cid.StartsWith("<") AndAlso cid.EndsWith(">") Then
                        cid = cid.Substring(1, cid.Length - 2)
                    End If
                    Return If(String.IsNullOrEmpty(cid), Nothing, cid)
                Finally
                    Marshal.ReleaseComObject(pa)
                End Try
            Catch ex As COMException
                ' Content-ID が未設定の添付ファイルでは例外が発生する（正常動作）
                Return Nothing
            End Try
        End Function

        ' ════════════════════════════════════════════════════════════
        '  ファイルシステムヘルパー
        ' ════════════════════════════════════════════════════════════

        Private Shared Function GetShortId(entryId As String) As String
            If String.IsNullOrEmpty(entryId) Then Return "00000000"
            Using md5 As MD5 = MD5.Create()
                Dim hash As Byte() = md5.ComputeHash(Encoding.UTF8.GetBytes(entryId))
                Return BitConverter.ToString(hash, 0, 4).Replace("-", "").ToLower()
            End Using
        End Function

        Private Shared Function SanitizeFileName(fileName As String) As String
            If String.IsNullOrEmpty(fileName) Then Return "attachment"
            Dim invalid As Char() = Path.GetInvalidFileNameChars()
            Dim sb As New StringBuilder()
            For Each c As Char In fileName
                If Array.IndexOf(invalid, c) >= 0 Then
                    sb.Append("_"c)
                Else
                    sb.Append(c)
                End If
            Next
            Dim result As String = sb.ToString()
            Return If(String.IsNullOrEmpty(result), "attachment", result)
        End Function

        Private Shared Function GetUniqueFilePath(dir As String, fileName As String) As String
            Dim candidate As String = Path.Combine(dir, fileName)
            If Not File.Exists(candidate) Then Return candidate

            Dim nameOnly As String = Path.GetFileNameWithoutExtension(fileName)
            Dim ext As String = Path.GetExtension(fileName)
            Dim counter As Integer = 1
            Do
                Dim newName As String = nameOnly & "_" & counter.ToString() & ext
                candidate = Path.Combine(dir, newName)
                If Not File.Exists(candidate) Then Return candidate
                counter += 1
            Loop
        End Function

        ' ════════════════════════════════════════════════════════════
        '  IDisposable
        ' ════════════════════════════════════════════════════════════

        Public Sub Dispose() Implements IDisposable.Dispose
            If Not _disposed Then
                If _ns IsNot Nothing Then Marshal.ReleaseComObject(_ns)
                ' Outlook アプリ本体はユーザーが使用中の可能性があるため Release しない
                _disposed = True
            End If
        End Sub

    End Class

End Namespace
