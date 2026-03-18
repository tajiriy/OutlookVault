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

        Private ReadOnly _app As Outlook.Application
        Private ReadOnly _ns As Outlook.NameSpace
        Private _disposed As Boolean

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
        '  フォルダ一覧
        ' ════════════════════════════════════════════════════════════

        ''' <summary>全ストアを検索してフォルダ表示名の一覧を返す。</summary>
        Public Function GetAvailableFolderNames() As List(Of String)
            Dim result As New List(Of String)()
            Dim stores As Outlook.Stores = _ns.Stores
            For i As Integer = 1 To stores.Count
                Dim store As Outlook.Store = stores.Item(i)
                Dim root As Outlook.MAPIFolder = store.GetRootFolder()
                CollectFolderNames(root, result)
            Next
            Return result
        End Function

        Private Sub CollectFolderNames(folder As Outlook.MAPIFolder, result As List(Of String))
            result.Add(folder.Name)
            Dim subFolders As Outlook.Folders = folder.Folders
            For i As Integer = 1 To subFolders.Count
                Dim childFolder As Outlook.MAPIFolder = CType(subFolders.Item(i), Outlook.MAPIFolder)
                CollectFolderNames(childFolder, result)
            Next
        End Sub

        ''' <summary>フォルダ名でフォルダを検索して返す。見つからない場合は Nothing。</summary>
        Public Function FindFolder(folderName As String) As Outlook.MAPIFolder
            Dim stores As Outlook.Stores = _ns.Stores
            For i As Integer = 1 To stores.Count
                Dim store As Outlook.Store = stores.Item(i)
                Dim root As Outlook.MAPIFolder = store.GetRootFolder()
                Dim found As Outlook.MAPIFolder = SearchFolder(root, folderName)
                If found IsNot Nothing Then Return found
            Next
            Return Nothing
        End Function

        Private Function SearchFolder(folder As Outlook.MAPIFolder, name As String) As Outlook.MAPIFolder
            If folder.Name = name Then Return folder
            Dim subFolders As Outlook.Folders = folder.Folders
            For i As Integer = 1 To subFolders.Count
                Dim childFolder As Outlook.MAPIFolder = CType(subFolders.Item(i), Outlook.MAPIFolder)
                Dim found As Outlook.MAPIFolder = SearchFolder(childFolder, name)
                If found IsNot Nothing Then Return found
            Next
            Return Nothing
        End Function

        ''' <summary>フォルダ内のアイテム数を返す（MailItem のみカウント）。</summary>
        Public Function GetMailItemCount(folderName As String) As Integer
            Dim folder As Outlook.MAPIFolder = FindFolder(folderName)
            If folder Is Nothing Then Return 0
            Return folder.Items.Count
        End Function

        ' ════════════════════════════════════════════════════════════
        '  メールデータ抽出
        ' ════════════════════════════════════════════════════════════

        ''' <summary>
        ''' MailItem から Models.Email を生成して返す。ThreadId は設定しない。
        ''' </summary>
        Public Function ExtractEmailData(mailItem As Outlook.MailItem) As Models.Email
            Dim email As New Models.Email()

            ' ── MAPI プロパティ ──────────────────────────────────
            Dim pa As Outlook.PropertyAccessor = mailItem.PropertyAccessor
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

            ' ── 件名 ────────────────────────────────────────────
            email.EntryId = mailItem.EntryID
            email.Subject = mailItem.Subject
            email.NormalizedSubject = ThreadingService.NormalizeSubject(email.Subject)

            ' ── 差出人 ──────────────────────────────────────────
            email.SenderName = mailItem.SenderName
            email.SenderEmail = GetSenderEmail(mailItem)

            ' ── 受信者 ──────────────────────────────────────────
            Dim recipients As Outlook.Recipients = mailItem.Recipients
            email.ToRecipients = SerializeRecipients(recipients, Outlook.OlMailRecipientType.olTo)
            email.CcRecipients = SerializeRecipients(recipients, Outlook.OlMailRecipientType.olCC)
            email.BccRecipients = SerializeRecipients(recipients, Outlook.OlMailRecipientType.olBCC)

            ' ── 本文 ────────────────────────────────────────────
            email.BodyText = mailItem.Body
            email.BodyHtml = mailItem.HTMLBody

            ' ── 日時 ────────────────────────────────────────────
            Dim receivedTime As DateTime = mailItem.ReceivedTime
            email.ReceivedAt = If(receivedTime > DateTime.MinValue, receivedTime, DateTime.Now)
            Dim sentOn As DateTime = mailItem.SentOn
            If sentOn > DateTime.MinValue Then email.SentAt = sentOn

            ' ── フォルダ名 ──────────────────────────────────────
            Dim parentFolder As Outlook.MAPIFolder =
                TryCast(mailItem.Parent, Outlook.MAPIFolder)
            email.FolderName = If(parentFolder IsNot Nothing, parentFolder.Name, Nothing)

            ' ── 添付ファイル（OLE 埋め込み・インライン画像を除いた実ファイル数でフラグを設定）──────
            Dim realAttachCount As Integer = 0
            Dim mailAtts As Outlook.Attachments = mailItem.Attachments
            For attIdx As Integer = 1 To mailAtts.Count
                Dim a As Outlook.Attachment = CType(mailAtts.Item(attIdx), Outlook.Attachment)
                If a.Type = Outlook.OlAttachmentType.olOLE Then Continue For
                Dim cid As String = GetAttachmentContentId(a)
                If Not String.IsNullOrEmpty(cid) Then Continue For  ' インライン画像はカウント外
                realAttachCount += 1
            Next
            email.HasAttachments = realAttachCount > 0

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
            If atts.Count = 0 Then Return result

            Dim dateStr As String = mailItem.ReceivedTime.ToString("yyyyMMdd_HHmmss")
            Dim shortId As String = GetShortId(mailItem.EntryID)
            Dim subDir As String = dateStr & "_" & shortId

            ' Outlook COM の SaveAsFile は相対パスを解釈できないため絶対パスに変換する
            Dim saveDir As String = Path.GetFullPath(Path.Combine(attachBaseDir, subDir))
            Dim dirCreated As Boolean = False

            For i As Integer = 1 To atts.Count
                Dim att As Outlook.Attachment = CType(atts.Item(i), Outlook.Attachment)

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
            Next

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
            Catch
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

        Private Shared Function GetSenderEmail(mailItem As Outlook.MailItem) As String
            If mailItem.SenderEmailType <> "EX" Then
                Return mailItem.SenderEmailAddress
            End If
            Try
                Dim sender As Outlook.AddressEntry = mailItem.Sender
                If sender IsNot Nothing Then
                    Dim exUser As Outlook.ExchangeUser = sender.GetExchangeUser()
                    If exUser IsNot Nothing Then Return exUser.PrimarySmtpAddress
                End If
            Catch
            End Try
            Return mailItem.SenderEmailAddress
        End Function

        ''' <summary>指定タイプの受信者を JSON 配列文字列にシリアライズする。</summary>
        Private Shared Function SerializeRecipients(recipients As Outlook.Recipients,
                                                    recipType As Outlook.OlMailRecipientType) As String
            Dim expectedType As Integer = CType(recipType, Integer)
            Dim parts As New List(Of String)()

            For i As Integer = 1 To recipients.Count
                Dim r As Outlook.Recipient = recipients.Item(i)
                If r.Type <> expectedType Then Continue For

                Dim emailAddr As String = r.Address
                ' Exchange アドレスの場合は SMTP アドレスを解決
                If r.AddressEntry IsNot Nothing AndAlso r.AddressEntry.Type = "EX" Then
                    Try
                        Dim exUser As Outlook.ExchangeUser = r.AddressEntry.GetExchangeUser()
                        If exUser IsNot Nothing Then emailAddr = exUser.PrimarySmtpAddress
                    Catch
                    End Try
                End If

                parts.Add("{""name"":" & JsonStr(r.Name) & ",""email"":" & JsonStr(emailAddr) & "}")
            Next

            Return "[" & String.Join(",", parts) & "]"
        End Function

        Private Shared Function JsonStr(s As String) As String
            If s Is Nothing Then Return "null"
            Return """" & s.Replace("\", "\\").Replace("""", "\""") & """"
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
                Dim val As Object = pa.GetProperty(PropAttachContentId)
                If val Is Nothing Then Return Nothing
                If Not TypeOf val Is String Then Return Nothing
                Dim cid As String = CType(val, String).Trim()
                ' "<xxx@yyy>" 形式の山括弧を除去して正規化
                If cid.StartsWith("<") AndAlso cid.EndsWith(">") Then
                    cid = cid.Substring(1, cid.Length - 2)
                End If
                Return If(String.IsNullOrEmpty(cid), Nothing, cid)
            Catch
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
